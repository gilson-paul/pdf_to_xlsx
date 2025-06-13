import base64
import glob
import os
import shutil
from typing import List, Optional
import io
import numpy as np
import pandas as pd
import pymupdf  # PyMuPDF
from openai import OpenAI
from pydantic import BaseModel, Field
from xlsxwriter.utility import xl_rowcol_to_cell
from llama_cloud_services import LlamaParse


def process_pdf_to_excel_with_images(
    pdf_path: str,
    output_folder: str,
    output_excel_file: str,
    llama_api_key: str,
    openai_api_key: str,
):
    """
    Parses a PDF to extract product information and images, then combines them
    into a single Excel file with embedded images.

    This function performs the following steps:
    1.  Parses the PDF content into markdown using LlamaParse.
    2.  Extracts structured product data from the markdown using OpenAI's GPT model.
    3.  Extracts all images from the PDF using PyMuPDF.
    4.  Creates a pandas DataFrame from the extracted product data.
    5.  Saves the DataFrame to an Excel file.
    6.  Embeds the extracted images into the corresponding rows of the Excel file.

    Args:
        pdf_path (str): The local file path to the input PDF document.
        output_folder (str): The path to the folder where extracted images will be saved.
        output_excel_file (str): The desired name for the output Excel file.
        llama_api_key (str): Your LlamaParse API key.
        openai_api_key (str): Your OpenAI API key.
    """

    # --- 1. Parse PDF to Markdown using LlamaParse ---
    try:
        parser = LlamaParse(api_key=llama_api_key)
        print(f"Parsing PDF: {pdf_path} with LlamaParse...")
        result = parser.parse(pdf_path)

        if not result or not result.pages:
            print("LlamaParse returned no pages or an empty result.")
            return

        all_markdown = "".join(
            page.md + "\n\n" for page in result.pages if page.md
        ).strip()
        print("PDF parsing complete.")

    except FileNotFoundError:
        print(f"Error: The file at {pdf_path} was not found.")
        return
    except Exception as e:
        print(f"An error occurred during LlamaParse: {e}")
        return

    # --- Pydantic Models for Structured Data Extraction ---
    class Product(BaseModel):
        """Represents a single product with its details."""

        style_id: str = Field(
            ...,
            description="The style identifier of the product, e.g., 'NCC-895 S'. Always present.",
        )
        sku: Optional[str] = Field(
            None, alias="SKU", description="The SKU of the product, e.g., 'NCC-895'."
        )
        price: Optional[str] = Field(
            None, description="The price of the product, if available."
        )
        color: Optional[str] = Field(
            None, description="The color of the product, if available."
        )

    class DataExtractor(BaseModel):
        """A container for the list of extracted product data."""

        product_data: List[Product]

    # --- 2. Extract Product Data from Markdown using OpenAI ---
    user_prompt = f""""You are an expert data extraction assistant specializing in parsing structured text from markdown.

    Your task is to analyze the following markdown string, which contains product information organized into multiple tables.
    
    **Instructions:**
    1.  **Identify Data Tables:** Scan the markdown to find all the product tables.
    2.  **Understand the Structure:** In each table, **each column represents a single product**. The data for each product is arranged vertically down the column.
    3.  **Extract Fields from Columns:** For each product column you find:
    
        - The **first data row** (the header row in the table) contains the "style_id" or the "product name", store either under "style_id".
        - The **second data row** contains the "SKU".
        - The **third data row** contains the "Color".
        - If The "Price" field is present in the source data, extract that too, else you must set the value of "Price" to `None`
        - Somtimes the order might be different so makes decisions according to the given data.
    4.  **Maintain Order:** Extract products from left to right within each table. Process the tables sequentially from top to bottom as they appear in the document.
    5.  **Format Output:** Return a single JSON array containing one dictionary for each extracted product. Do not include any explanatory text, markdown formatting, or anything other than the JSON object itself.
    
    **Example Input Snippet:**
    
    | NCC-895 S | NCC-896 S |
    | --------- | --------- |
    | NCC-895   | NCC-896   |
    | SILVER    | SILVER    |
    
    **Correct Output for the Snippet Above:**
    
    ```json
    [
      {{
        "style_id": "NCC-895 S",
        "SKU": "NCC-895",
        "Price": null,
        "Color": "SILVER"
      }},
      {{
        "style_id": "NCC-896 S",
        "SKU": "NCC-896",
        "Price": null,
        "Color": "SILVER"
      }}
    ]

    """

    try:
        client = OpenAI(api_key=openai_api_key)
        print("Sending request to OpenAI API to extract data...")
        response = client.beta.chat.completions.parse(
            model="gpt-4o",
            messages=[
                {
                    "role": "system",
                    "content": [{"type": "text", "text": user_prompt}],
                },
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": f"Mardown text:\n{all_markdown}",
                        }
                    ],
                },
            ],
            response_format=DataExtractor,
        )

        data: DataExtractor = response.choices[0].message.parsed
        print("Data extraction successful.")

        product_list = [
            product.model_dump(by_alias=True) for product in data.product_data
        ]

        if not product_list:
            print("Warning: No products were extracted from the PDF.")
            product_df = pd.DataFrame(
                columns=["Image", "Style ID", "SKU", "Price", "Color"]
            )
        else:
            product_df = pd.DataFrame(product_list)
            # Add the 'product_image' column at the beginning
            product_df.insert(0, "Image", None)
            # Ensure column order and naming matches the requirement
            product_df = product_df.rename(
                columns={
                    "style_id": "Style ID",
                    "sku": "SKU",
                    "price": "Price",
                    "color": "Color",
                }
            )
            final_columns = ["Image", "Style ID", "SKU", "Price", "Color"]
            product_df = product_df.reindex(columns=final_columns)


    except Exception as e:
        print(f"An unexpected error occurred during OpenAI data extraction: {e}")
        product_df = pd.DataFrame(
            columns=["Image", "Style ID", "SKU", "Price", "Color"]
        )

    # --- 3. Extract Images from PDF ---
    def is_single_product_image(image_str: str, api_key: str=openai_api_key) -> bool:
        """
        Analyzes an image using OpenAI's vision model to determine if it's a single
        product image suitable for a catalog.

        This function sends an image to the GPT-4o model and asks it to classify
        it. It is intended to distinguish clean product shots (on white, black, or
        plain backgrounds) from banners, collages, or lifestyle images.

        Args:
            image_path (str): The local file path to the image.
            api_key (str): Your OpenAI API key.

        Returns:
            bool: True if the model identifies the image as a single product shot,
                  False otherwise (e.g., it's a banner, collage, or an error occurred).
        """
            
        try:
            # Initialize the OpenAI client
            client = OpenAI(api_key=openai_api_key)

            # Encode the image in base64
            base64_image = image_str

            # Construct the payload for the API call
            messages = [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text",
                            "text": (
                                """
                                Objective: Analyze an image to filter out banners and collages from a product collection. Respond with only the word 'True' or 'False'.

                                Answer 'False' if the image is:

                                  A Banner or Advertisement: An image with significant marketing text, graphic overlays, or promotional logos that are not part of the product itself.
                                  A Collage: An image that artificially combines multiple, separate products that are not sold together as a single set.

                                Answer 'True' if the image is:

                                  A photo of a product. This includes clean single-product shots, products shown in their packaging, and lifestyle images of the product being worn or used, as long as it is not a banner or a collage.
                                """
                            ),
                        },
                        {
                            "type": "image_url",
                            "image_url": {
                                "url": f"data:image/png;base64,{base64_image}",
                                "detail": "low" # Use 'low' detail for faster, cheaper classification
                            },
                        },
                    ],
                }
            ]

            # Call the OpenAI API
            response = client.chat.completions.create(
                model="gpt-4.1-nano",
                messages=messages,
                max_tokens=5, # Restrict response length
            )

            # Parse the response
            decision = response.choices[0].message.content.strip().lower()
            
            #print(f"Image '{os.path.basename(image_path)}' classified as: {decision}")
            
            return decision == "true"

        except Exception as e:
            print(f"An error occurred while analyzing image {os.path.basename(image_path)}: {e}")
            # Default to False in case of any API or processing error
            return False



    if os.path.exists(output_folder):
        shutil.rmtree(output_folder)
    os.makedirs(output_folder)
    print(f"Output folder '{output_folder}' created/cleaned.")

    try:
        doc = pymupdf.open(pdf_path)
        for page_index in range(len(doc)):
            page = doc[page_index]
            image_list = page.get_images(full=True)
            white_images_skipped = 0

            for image_index, img_info in enumerate(image_list, start=1):
                xref = img_info[0]
                pix = pymupdf.Pixmap(doc, xref)

                if pix.n - pix.alpha > 3:  # CMYK -> RGB
                    pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                
                pil_image = pix.pil_image()
                # Skip pure white images
                if (np.array(pil_image) == 255).all():
                    white_images_skipped += 1
                    continue
                
                buffered = io.BytesIO()
    
                # Save the PIL image to the buffer in the specified format
                pil_image.save(buffered, format="PNG")
                
                # Get the bytes from the buffer
                img_bytes = buffered.getvalue()
                
                # Encode the bytes to base64 and decode to a utf-8 string
                base64_str = base64.b64encode(img_bytes).decode("utf-8")

                if not is_single_product_image(base64_str):
                    continue

                image_filename = f"page_{page_index + 1}-image_{image_index - white_images_skipped}.png"
                pix.save(os.path.join(output_folder, image_filename))
            
            #if image_list:
            #    found_count = len(image_list) - white_images_skipped
            #    print(f"Found {found_count} non-white images on page {page_index + 1}")
            #else:
            #    print(f"No images found on page {page_index + 1}")

        print("All PDF pages processed for images.")
        doc.close()

    except Exception as e:
        print(f"An error occurred during image extraction: {e}")


    # --- 4. Sort Images and Write to Excel ---
    image_files = glob.glob(os.path.join(output_folder, "*.png"))
    # Sort images by page number, then by image number
    images_list = sorted(
        image_files,
        key=lambda x: (
            int(os.path.basename(x).split("_")[1].split("-")[0]),
            int(os.path.basename(x).split("_")[-1].split(".")[0]),
        ),
    )

    try:
        with pd.ExcelWriter(output_excel_file, engine="xlsxwriter") as writer:
            product_df.to_excel(writer, sheet_name="Products", index=False)
            workbook = writer.book
            worksheet = writer.sheets["Products"]

            # --- Set Column and Row Formatting ---
            worksheet.set_column("A:A", 20)  # Width for the 'Image' column
            default_row_height = 80

            # --- Loop Through Images and Insert Them ---
            for index, image_path in enumerate(images_list):
                excel_row = index + 1  # Excel is 1-indexed, +1 to skip header

                if excel_row > len(product_df):
                    print(
                        f"Warning: More images ({len(images_list)}) than data rows ({len(product_df)}). "
                        f"Stopping image insertion."
                    )
                    break
                
                # Set the specific row's height
                worksheet.set_row(excel_row, default_row_height)
                
                # Insert the image, centered in the cell
                worksheet.embed_image(excel_row, 0, image_path, {'object_position': 1})

        print(f"\nSuccess! Your Excel file '{output_excel_file}' has been created.")

    except Exception as e:
        print(f"An error occurred while creating the Excel file: {e}")


# # --- Example Usage ---
# # Make sure to have your API keys available, for example as environment variables
# # or using a secret management tool.
# if __name__ == '__main__':
#     # from google.colab import userdata # Example for Google Colab
#     # LlamaParse API Key
#     # llama_key = userdata.get("llama_parse")
#     # OpenAI API Key
#     # openai_key = userdata.get('OPENAI_API_KEY')
#     
#     # Use placeholders if running locally without a secret manager
#     llama_key = "YOUR_LLAMA_PARSE_API_KEY"
#     openai_key = "YOUR_OPENAI_API_KEY"
#
#     # --- Define Paths ---
#     pdf_file = 'sample_catalog.pdf'  # Replace with your PDF file path
#     image_output_dir = 'extracted_images' # Folder to store images temporarily
#     final_excel_file = 'products_with_images.xlsx' # Name of the final output file
#
#     # --- Run the main function ---
#     process_pdf_to_excel_with_images(
#         pdf_path=pdf_file,
#         output_folder=image_output_dir,
#         output_excel_file=final_excel_file,
#         llama_api_key=llama_key,
#         openai_api_key=openai_key
#     )
