import streamlit as st
import pandas as pd
import os
import io
import base64
from util import process_pdf_to_excel_with_images
import shutil


# Ensure necessary directories exist
if not os.path.exists('temp_uploads'):
    os.makedirs('temp_uploads')
if not os.path.exists('extracted_images_streamlit'):
    os.makedirs('extracted_images_streamlit')
if not os.path.exists('output_excel_streamlit'):
    os.makedirs('output_excel_streamlit')

# --- Streamlit App ---

st.title("PDF Product Catalog Extractor")

st.write(
    "Upload a PDF product catalog, and this app will extract product "
    "information and images, generating an Excel file for download."
)

# File uploader
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

# Process button
process_button = st.button("Process PDF")

# Use secrets for API keys
# In a real Streamlit app, you would manage secrets securely.
# For demonstration in Colab, you might need to get them differently,
# but for a deployed Streamlit app, use st.secrets.toml
try:
    llama_api_key = st.secrets["LLAMA_API_KEY"]
except KeyError:
    llama_api_key = None
    st.warning("LLAMA_API_KEY not found in secrets. Processing may fail.")

try:
    openai_api_key = st.secrets["OPENAI_API_KEY"]
except KeyError:
    openai_api_key = None
    st.warning("OPENAI_API_KEY not found in secrets. Processing may fail.")


if uploaded_file is not None and process_button:
    if llama_api_key is None or openai_api_key is None:
        st.error("API keys are missing. Please configure your secrets.")
    else:
        # Save the uploaded file temporarily
        temp_pdf_path = os.path.join("temp_uploads", uploaded_file.name)
        with open(temp_pdf_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        st.info(f"Uploaded file: {uploaded_file.name}")
        st.info("Processing your PDF. This may take some time...")

        # Define output paths
        image_output_dir = "extracted_images_streamlit"
        final_excel_file = os.path.join("output_excel_streamlit", f"{os.path.splitext(uploaded_file.name)[0]}_extracted.xlsx")

        # Call the processing function
        # Make sure process_pdf_to_excel_with_images is defined or imported
        try:
            # Assuming process_pdf_to_excel_with_images is defined as in the user's previous code
            process_pdf_to_excel_with_images(
                pdf_path=temp_pdf_path,
                output_folder=image_output_dir,
                output_excel_file=final_excel_file,
                llama_api_key=llama_api_key,
                openai_api_key=openai_api_key,
            )
            st.success("Processing complete!")

            # Provide download link for the Excel file
            if os.path.exists(final_excel_file):
                with open(final_excel_file, "rb") as f:
                    excel_bytes = f.read()
                st.download_button(
                    label="Download Excel File",
                    data=excel_bytes,
                    file_name=os.path.basename(final_excel_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.error("Output Excel file was not created.")

        except Exception as e:
            st.error(f"An error occurred during processing: {e}")
            st.write("Please check the logs for more details.")

        finally:
            # Clean up temporary files and folders (optional but good practice)
            # Be careful with shutil.rmtree - ensure you are in the correct directory
            if os.path.exists("temp_uploads"):
                shutil.rmtree("temp_uploads")
            if os.path.exists("extracted_images_streamlit"):
                shutil.rmtree("extracted_images_streamlit")

elif uploaded_file is None and process_button:
    st.warning("Please upload a PDF file before clicking Process.")
