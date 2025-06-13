# prompt: streamlit code for taking in the inputs uploaded pdf file and giving output.xlsx for download

import streamlit as st
import pandas as pd
import os
import tempfile
import base64
from io import BytesIO
from util import process_pdf_to_excel_with_images
# Assuming the functions process_pdf_to_excel_with_images and is_single_product_image
# are defined in your code.
# You might need to adapt the code structure to work within a Streamlit app.

# Define a download button function
def download_button(object_to_download, download_filename, button_text):
    """
    Generates a link to download the given object_to_download.
    In this example, object_to_download is a bytes object (the Excel file).
    """
    if isinstance(object_to_download, bytes):
        b64 = base64.b64encode(object_to_download).decode()
    else:
        raise TypeError("Object to download must be bytes")

    mime_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    href = f'<a href="data:{mime_type};base64,{b64}" download="{download_filename}">{button_text}</a>'
    return href

st.title("PDF Product Catalog Extractor")

st.write("Upload a PDF product catalog and extract product data with images into an Excel file.")

uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

# Define API keys (replace with your actual keys or secure method)
# In a real application, use Streamlit Secrets or environment variables
# st.secrets["llama_parse"] and st.secrets["OPENAI_API_KEY"]
# For this example running in a notebook, we use placeholders.
# You MUST replace these with your actual keys or a secrets management method.
llama_api_key = os.getenv("LLAMA_PARSE_API_KEY") # Replace with your key or secrets mechanism
openai_api_key = os.getenv("OPENAI_API_KEY") # Replace with your key or secrets mechanism

if not llama_api_key or not openai_api_key:
    st.error("API keys not found. Please set your LLAMA_PARSE_API_KEY and OPENAI_API_KEY.")
    st.stop()


if uploaded_file is not None:
    # Create a temporary directory and save the uploaded PDF
    with tempfile.TemporaryDirectory() as tmpdir:
        pdf_path = os.path.join(tmpdir, uploaded_file.name)
        with open(pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.success(f"Uploaded {uploaded_file.name}")

        output_folder = os.path.join(tmpdir, "extracted_images")
        output_excel_file = "extracted_products.xlsx"
        output_excel_path = os.path.join(tmpdir, output_excel_file)

        # Run the extraction process with a loading spinner
        with st.spinner("Extracting data and images from PDF..."):
            process_pdf_to_excel_with_images(
                pdf_path=pdf_path,
                output_folder=output_folder,
                output_excel_file=output_excel_path,
                llama_api_key=llama_api_key,
                openai_api_key=openai_api_key,
            )

        # Check if the Excel file was created
        if os.path.exists(output_excel_path):
            st.success("Extraction complete!")

            # Read the generated Excel file into bytes for download
            with open(output_excel_path, "rb") as f:
                excel_bytes = f.read()

            # Provide the download link
            st.markdown(
                download_button(
                    excel_bytes, output_excel_file, "Download Excel File"
                ),
                unsafe_allow_html=True,
            )

        else:
            st.error("Failed to create the Excel file. Check logs for details.")

