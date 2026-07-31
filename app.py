import streamlit as st
import os
import json
import io
from google.cloud import vision
from google.oauth2 import service_account
from PIL import Image
import docx

# Set page config
st.set_page_config(page_title="Professional OCR App", layout="wide")

def get_vision_client():
    """
    Authenticate and return a Google Cloud Vision client.
    Tries to use Streamlit secrets first, then falls back to a local JSON file.
    """
    # 1. Try Streamlit Secrets
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        # Convert dictionary to proper credentials object
        credentials = service_account.Credentials.from_service_account_info(creds_dict)
        return vision.ImageAnnotatorClient(credentials=credentials)
    
    # 2. Try Local JSON file
    local_creds_path = "google-credentials.json"
    if os.path.exists(local_creds_path):
        return vision.ImageAnnotatorClient.from_service_account_json(local_creds_path)
    
    # 3. Try standard Google Application Default Credentials (e.g. environment variable)
    if "GOOGLE_APPLICATION_CREDENTIALS" in os.environ:
         return vision.ImageAnnotatorClient()

    st.error("Google Cloud credentials not found. Please configure Streamlit secrets or provide a google-credentials.json file.")
    return None

def perform_ocr(image_bytes, client):
    """
    Perform OCR using Google Cloud Vision API (document_text_detection).
    """
    image = vision.Image(content=image_bytes)
    
    # Use document_text_detection for dense text, pages, complex layouts
    response = client.document_text_detection(image=image)
    
    if response.error.message:
        st.error(f"Error from Google Cloud Vision API: {response.error.message}")
        return ""
        
    return response.full_text_annotation.text

def create_word_doc(text):
    """
    Create a Word document from the given text.
    """
    doc = docx.Document()
    doc.add_paragraph(text)
    
    # Save the document to an in-memory bytes buffer
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def main():
    st.title("Professional OCR App")
    st.markdown("Upload an image to extract text using Google Cloud Vision API.")

    client = get_vision_client()
    if not client:
        st.stop()

    # File uploader
    uploaded_file = st.file_uploader("Choose an image...", type=["jpg", "jpeg", "png"])

    if uploaded_file is not None:
        # Display the image
        image = Image.open(uploaded_file)
        st.image(image, caption='Uploaded Image', use_column_width=True)

        st.markdown("---")
        st.subheader("Extracted Text")

        with st.spinner("Extracting text..."):
            # Read image bytes for the Vision API
            image_bytes = uploaded_file.getvalue()
            
            # Perform OCR
            extracted_text = perform_ocr(image_bytes, client)

        if extracted_text:
            # Editable text area
            edited_text = st.text_area("Edit the extracted text:", value=extracted_text, height=300)

            st.markdown("### Download")
            
            col1, col2 = st.columns(2)
            
            # Download as Word (.docx)
            with col1:
                word_bytes = create_word_doc(edited_text)
                st.download_button(
                    label="Download as Word (.docx)",
                    data=word_bytes,
                    file_name="extracted_text.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            
            # Download as Text (.txt)
            with col2:
                st.download_button(
                    label="Download as Text (.txt)",
                    data=edited_text,
                    file_name="extracted_text.txt",
                    mime="text/plain"
                )
        else:
            st.warning("No text was extracted from the image.")

if __name__ == "__main__":
    main()
