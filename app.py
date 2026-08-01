import streamlit as st
import io
import requests
from PIL import Image, ImageEnhance
import docx
import pandas as pd

# Set page config
st.set_page_config(page_title="Advanced OCR SaaS", page_icon="📝", layout="wide")

def get_ocr_space_key():
    """
    Retrieve the OCR.space API key from Streamlit secrets.
    """
    if "OCR_SPACE_API_KEY" in st.secrets:
        return st.secrets["OCR_SPACE_API_KEY"]
    st.sidebar.error("OCR.space API key not found. Please configure Streamlit secrets with 'OCR_SPACE_API_KEY'.")
    return None

def perform_ocr(image_bytes, api_key, language, is_table):
    """
    Perform OCR using OCR.space API.
    """
    url = "https://api.ocr.space/parse/image"
    
    # We can send the image file directly
    files = {"file": ("image.jpg", image_bytes, "image/jpeg")}
    
    data = {
        "apikey": api_key,
        "language": language, 
        "OCREngine": 2, # Engine 2 is excellent for Western languages (English, French)
        "isTable": "true" if is_table else "false",
        "scale": "true",
        "detectOrientation": "true"
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

def preprocess_image(image_bytes):
    """
    Preprocess image by converting to grayscale and increasing contrast
    to improve OCR accuracy.
    """
    image = Image.open(io.BytesIO(image_bytes))
    
    # Convert to grayscale
    image = image.convert('L')
    
    # Increase contrast
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    
    # Save back to bytes
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='JPEG')
    return img_byte_arr.getvalue()

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
    # --- Header ---
    st.title("📝 Advanced OCR & Text Extraction")
    st.markdown("Welcome to the premium Optical Character Recognition (OCR) tool. Easily extract structured text from images and documents in English and French, and export them directly to Word or Excel.")
    st.markdown("---")

    api_key = get_ocr_space_key()
    if not api_key:
        st.stop()
        
    # --- Sidebar Controls ---
    st.sidebar.title("Configuration")
    
    st.sidebar.markdown("### 1. Select Language")
    language_map = {
        "English": "eng",
        "French": "fre"
    }
    selected_language = st.sidebar.selectbox("Document Language", options=list(language_map.keys()))
    language_code = language_map[selected_language]
    
    st.sidebar.markdown("### 2. Processing Options")
    enable_table_parsing = st.sidebar.checkbox("Enable Table Parsing", value=True, help="Turn this on to better extract structured data like invoices or tables.")

    st.sidebar.markdown("### 3. Provide Image")
    input_method = st.sidebar.radio("Choose input method:", ("Upload an Image", "Take a Photo"))
    
    uploaded_file = None
    if input_method == "Upload an Image":
        uploaded_file = st.sidebar.file_uploader("Upload Image (JPG, PNG)", type=["jpg", "jpeg", "png"])
    else:
        uploaded_file = st.sidebar.camera_input("Take a picture")

    # --- Main Area ---
    if uploaded_file is not None:
        # Display the image in an expander to keep the UI clean
        with st.expander("View Uploaded Image"):
            image = Image.open(uploaded_file)
            st.image(image, caption='Source Document', use_column_width=True)

        st.subheader("Extracted Data")

        with st.spinner("Analyzing document and extracting text..."):
            # Read image bytes
            image_bytes = uploaded_file.getvalue()
            
            # Preprocess the image for better OCR results
            processed_image_bytes = preprocess_image(image_bytes)
            
            # Perform OCR
            extracted_text = perform_ocr(processed_image_bytes, api_key, language_code, enable_table_parsing)

        if extracted_text:
            st.success("Extraction completed successfully!")
            
            # Editable text area
            edited_text = st.text_area("Review and edit the extracted text:", value=extracted_text, height=350)

            st.markdown("### Export Options")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                word_bytes = create_word_doc(edited_text)
                st.download_button(
                    label="📄 Download as Word",
                    data=word_bytes,
                    file_name="extracted_text.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
            with col2:
                excel_bytes = create_excel_doc(edited_text)
                st.download_button(
                    label="📊 Download as Excel",
                    data=excel_bytes,
                    file_name="extracted_text.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            with col3:
                st.download_button(
                    label="📝 Download as Text",
                    data=edited_text,
                    file_name="extracted_text.txt",
                    mime="text/plain",
                    use_container_width=True
                )
        else:
            st.warning("No text was extracted from the image. Please try a clearer document or adjust the settings.")
    else:
        st.info("👈 Please use the sidebar to upload an image or take a photo to begin.")

if __name__ == "__main__":
    main()
