import streamlit as st
import io
import requests
from PIL import Image
import docx

# Set page config
st.set_page_config(page_title="Professional OCR App", layout="wide")

def get_ocr_space_key():
    """
    Retrieve the OCR.space API key from Streamlit secrets.
    """
    if "OCR_SPACE_API_KEY" in st.secrets:
        return st.secrets["OCR_SPACE_API_KEY"]
    st.error("OCR.space API key not found. Please configure Streamlit secrets with 'OCR_SPACE_API_KEY'.")
    return None

def perform_ocr(image_bytes, api_key):
    """
    Perform OCR using OCR.space API.
    """
    url = "https://api.ocr.space/parse/image"
    
    # We can send the image file directly
    files = {"file": ("image.jpg", image_bytes, "image/jpeg")}
    data = {
        "apikey": api_key,
        "language": "eng", 
        "OCREngine": 2 # Engine 2 is usually better for documents
    }
    
    try:
        response = requests.post(url, files=files, data=data)
        response.raise_for_status()
        result = response.json()
        
        if result.get("IsErroredOnProcessing"):
            st.error(f"OCR Error: {result.get('ErrorMessage')}")
            return ""
            
        parsed_results = result.get("ParsedResults", [])
        if not parsed_results:
            return ""
            
        extracted_text = ""
        for res in parsed_results:
            extracted_text += res.get("ParsedText", "") + "\n"
            
        return extracted_text.strip()
    except Exception as e:
        st.error(f"Failed to connect to OCR.space API: {e}")
        return ""

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
    st.markdown("Upload an image to extract text using OCR.space API.")

    api_key = get_ocr_space_key()
    if not api_key:
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
            # Read image bytes
            image_bytes = uploaded_file.getvalue()
            
            # Perform OCR
            extracted_text = perform_ocr(image_bytes, api_key)

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
