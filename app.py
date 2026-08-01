import streamlit as st
import io
import requests
from PIL import Image
import docx
import pandas as pd

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

def perform_ocr(image_bytes, api_key, language, is_table):
    """
    Perform OCR using OCR.space API.
    """
    url = "https://api.ocr.space/parse/image"
    
    # We can send the image file directly
    files = {"file": ("image.jpg", image_bytes, "image/jpeg")}
    
    # Engine 2 supports English and French well, but does not support Arabic.
    # We must use Engine 1 for Arabic to avoid the language parameter error.
    engine = 1 if language == "ara" else 2
    
    data = {
        "apikey": api_key,
        "language": language, 
        "OCREngine": engine,
        "isTable": "true" if is_table else "false"
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

def create_excel_doc(text):
    """
    Create an Excel document (.xlsx) from the text using pandas.
    When isTable=True is enabled in OCR.space, it typically outputs columns separated by tabs.
    """
    # Split text into lines
    lines = text.split('\n')
    # Split each line by tab character to separate columns
    data = [line.split('\t') for line in lines if line.strip()]
    df = pd.DataFrame(data)
    
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False)
    
    return bio.getvalue()

def main():
    st.title("Professional OCR App")
    st.markdown("Upload an image or take a photo to extract text using OCR.space API.")

    api_key = get_ocr_space_key()
    if not api_key:
        st.stop()
        
    st.sidebar.header("OCR Settings")
    language_map = {
        "English": "eng",
        "French": "fre",
        "Arabic": "ara"
    }
    selected_language = st.sidebar.selectbox("Select Document Language", options=list(language_map.keys()))
    language_code = language_map[selected_language]
    
    enable_table_parsing = st.sidebar.checkbox("Enable Table Parsing (isTable=True)", value=True)

    st.markdown("### Input Source")
    input_method = st.radio("Choose input method:", ("Upload an Image", "Take a Photo"))
    
    uploaded_file = None
    if input_method == "Upload an Image":
        # File uploader
        uploaded_file = st.file_uploader("Choose an image...", type=["jpg", "jpeg", "png"])
    else:
        # Camera input
        uploaded_file = st.camera_input("Take a picture")

    if uploaded_file is not None:
        # Display the image
        image = Image.open(uploaded_file)
        st.image(image, caption='Image for OCR', use_column_width=True)

        st.markdown("---")
        st.subheader("Extracted Text")

        with st.spinner("Extracting text..."):
            # Read image bytes
            image_bytes = uploaded_file.getvalue()
            
            # Perform OCR
            extracted_text = perform_ocr(image_bytes, api_key, language_code, enable_table_parsing)

        if extracted_text:
            # Editable text area
            edited_text = st.text_area("Edit the extracted text:", value=extracted_text, height=300)

            st.markdown("### Download")
            
            col1, col2, col3 = st.columns(3)
            
            # Download as Word (.docx)
            with col1:
                word_bytes = create_word_doc(edited_text)
                st.download_button(
                    label="Download as Word (.docx)",
                    data=word_bytes,
                    file_name="extracted_text.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            # Download as Excel (.xlsx)
            with col2:
                excel_bytes = create_excel_doc(edited_text)
                st.download_button(
                    label="Download as Excel (.xlsx)",
                    data=excel_bytes,
                    file_name="extracted_text.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            # Download as Text (.txt)
            with col3:
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
