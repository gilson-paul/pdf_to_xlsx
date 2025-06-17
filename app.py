import streamlit as st
import pandas as pd
import os
import io
import base64
from util import process_pdf_to_excel_with_images  # Assuming this is in a util.py file
import shutil
import json

# --- Directory Setup ---
# Ensure necessary directories exist at the start
def setup_directories():
    if not os.path.exists('temp_uploads'):
        os.makedirs('temp_uploads')
    if not os.path.exists('extracted_images_streamlit'):
        os.makedirs('extracted_images_streamlit')
    if not os.path.exists('output_excel_streamlit'):
        os.makedirs('output_excel_streamlit')

setup_directories()

# --- Streamlit App ---

st.title("PDF Product Catalog Extractor 📄➡️📊")

st.write(
    "Upload one or more PDF product catalogs. The app will extract product "
    "information and images, generating a separate Excel file for each PDF."
)

# --- File Uploader and Banner Options ---
uploaded_files = st.file_uploader(
    "Choose PDF files",
    type="pdf",
    accept_multiple_files=True
)

banner_options = {}
if uploaded_files:
    st.markdown("---")
    st.subheader("Uploaded Files & Options")
    for uploaded_file in uploaded_files:
        # Use a unique key for each checkbox based on the file's name
        key = f"banner_checkbox_{uploaded_file.name}"
        banner_options[uploaded_file.name] = st.checkbox(
            f"'{uploaded_file.name}' contains banners (check if yes)",
            key=key
        )
    st.markdown("---")

# Process button
process_button = st.button("Process All PDFs")

# --- API Key Handling ---
# Use secrets for API keys in a deployed Streamlit app via st.secrets
try:
    gemini_api_key = st.secrets["GEMINI_API_KEY"]
except (KeyError, FileNotFoundError):
    gemini_api_key = None
    st.warning("GEMINI_API_KEY not found in secrets. Processing may fail.")

try:
    openai_api_key = st.secrets["OPENAI_API_KEY"]
except (KeyError, FileNotFoundError):
    openai_api_key = None
    st.warning("OPENAI_API_KEY not found in secrets. Processing may fail.")

# --- Processing Logic ---
if process_button and uploaded_files:
    if gemini_api_key is None or openai_api_key is None:
        st.error("API keys are missing. Please configure your Streamlit secrets.")
    else:
        st.info("Starting batch processing...")

        # Process each uploaded file
        for uploaded_file in uploaded_files:
            # Save the uploaded file temporarily
            temp_pdf_path = os.path.join("temp_uploads", uploaded_file.name)
            with open(temp_pdf_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            has_banners = banner_options[uploaded_file.name]
            st.markdown(f"#### Processing `{uploaded_file.name}`")
            st.write(f"This PDF is being processed with the assumption that it **{'contains' if has_banners else 'does not contain'}** banners.")

            # Define output paths for this specific file
            image_output_dir = "extracted_images_streamlit"
            final_excel_file = os.path.join("output_excel_streamlit", f"{os.path.splitext(uploaded_file.name)[0]}_extracted.xlsx")

            # Call the processing function for the current file
            try:
                # NOTE: You may need to modify your `process_pdf_to_excel_with_images`
                # function to accept and use the `has_banners` variable.
                process_pdf_to_excel_with_images(
                    pdf_path=temp_pdf_path,
                    output_folder=image_output_dir,
                    output_excel_file=final_excel_file,
                    contains_banners = has_banners,
                    gemini_api_key=gemini_api_key,
                    openai_api_key=openai_api_key,
                    # has_banners=has_banners  # Example of passing the flag
                )
                st.success(f"Successfully processed `{uploaded_file.name}`!")

                # Provide a download link for the generated Excel file
                if os.path.exists(final_excel_file):
                    with open(final_excel_file, "rb") as f:
                        excel_bytes = f.read()
                    st.download_button(
                        label=f"Download Excel for '{uploaded_file.name}'",
                        data=excel_bytes,
                        file_name=os.path.basename(final_excel_file),
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"download_{uploaded_file.name}" # Unique key for the download button
                    )
                else:
                    st.error(f"Output Excel file was not created for `{uploaded_file.name}`.")

            except Exception as e:
                st.error(f"An error occurred while processing `{uploaded_file.name}`: {e}")
            finally:
                st.markdown("---") # Separator for clarity between files

        # Clean up temporary files after all processing is complete
        st.info("Cleaning up temporary files...")
        if os.path.exists("temp_uploads"):
            shutil.rmtree("temp_uploads")
        if os.path.exists("extracted_images_streamlit"):
            shutil.rmtree("extracted_images_streamlit")
        st.success("All tasks complete!")


elif process_button and not uploaded_files:
    st.warning("Please upload at least one PDF file before clicking 'Process All PDFs'.")
