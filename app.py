import streamlit as st
import io
import requests
from PIL import Image, ImageEnhance
import docx
import pandas as pd

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Advanced OCR SaaS",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# Custom CSS for Premium UI/UX
# -----------------------------------------------------------------------------
st.markdown("""
<style>
/* CSS Animation for the Hero Header */
@keyframes fadeIn {
    0% { opacity: 0; transform: translateY(-20px); }
    100% { opacity: 1; transform: translateY(0); }
}

.hero-text {
    font-size: 2.5rem !important;
    font-weight: 800;
    text-align: center;
    background: -webkit-linear-gradient(45deg, #4b6cb7, #182848);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    animation: fadeIn 1.5s ease-in-out;
    margin-bottom: 0.5rem;
}

.sub-hero {
    text-align: center;
    font-size: 1.2rem;
    color: #555;
    margin-bottom: 2rem;
    animation: fadeIn 2s ease-in-out;
}

/* 3-Step Guide Styling */
.step-card {
    background-color: #f8f9fa;
    border-radius: 10px;
    padding: 1.5rem 1rem;
    text-align: center;
    box-shadow: 0 4px 6px rgba(0,0,0,0.05);
    height: 100%;
}

.step-icon {
    font-size: 2.5rem;
    margin-bottom: 0.5rem;
}

.step-title {
    font-weight: bold;
    color: #333;
    font-size: 1.1rem;
}

/* Premium Button Styling */
div.stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white !important;
    border: none;
    border-radius: 8px;
    padding: 0.6rem 1.2rem;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    width: 100%;
}

div.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(0,0,0,0.3);
}

div.stDownloadButton > button {
    background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    color: white !important;
    border: none;
    border-radius: 8px;
    padding: 0.8rem 1.2rem;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    width: 100%;
    margin-bottom: 0.5rem;
}

div.stDownloadButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(0,0,0,0.3);
}

/* Text Area Styling */
.stTextArea textarea {
    border-radius: 8px;
    border: 2px solid #e0e0e0;
    padding: 1rem;
    font-family: 'Courier New', Courier, monospace;
    font-size: 1rem;
    transition: border-color 0.3s ease;
}

.stTextArea textarea:focus {
    border-color: #667eea;
    box-shadow: 0 0 0 2px rgba(102,126,234,0.2);
}

/* Mobile Padding Adjustments */
@media (max-width: 768px) {
    .block-container {
        padding-top: 2rem;
        padding-left: 1rem;
        padding-right: 1rem;
    }
    .hero-text {
        font-size: 1.8rem !important;
    }
}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Core Functions
# -----------------------------------------------------------------------------
def get_ocr_space_key():
    if "OCR_SPACE_API_KEY" in st.secrets:
        return st.secrets["OCR_SPACE_API_KEY"]
    st.error("OCR.space API key not found. Please configure Streamlit secrets with 'OCR_SPACE_API_KEY'.")
    return None

def preprocess_image(image_bytes):
    image = Image.open(io.BytesIO(image_bytes))
    image = image.convert('L')
    enhancer = ImageEnhance.Contrast(image)
    image = enhancer.enhance(2.0)
    
    img_byte_arr = io.BytesIO()
    image.save(img_byte_arr, format='JPEG')
    return img_byte_arr.getvalue()

def perform_ocr(image_bytes, api_key, language):
    url = "https://api.ocr.space/parse/image"
    files = {"file": ("image.jpg", image_bytes, "image/jpeg")}
    
    data = {
        "apikey": api_key,
        "language": language, 
        "OCREngine": 2, 
        "isTable": "true",
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

def create_word_doc(text):
    doc = docx.Document()
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_excel_doc(text):
    lines = text.split('\n')
    data = [line.split('\t') for line in lines if line.strip()]
    df = pd.DataFrame(data)
    
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, header=False)
    return bio.getvalue()

# -----------------------------------------------------------------------------
# Main App
# -----------------------------------------------------------------------------
def main():
    api_key = get_ocr_space_key()
    if not api_key:
        st.stop()

    # --- Sidebar ---
    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/3752/3752762.png", width=60)
        st.title("Settings")
        st.markdown("Select the primary language of your document.")
        
        language_map = {
            "English (ENG)": "eng",
            "French (FRE)": "fre"
        }
        selected_lang = st.selectbox("Language", options=list(language_map.keys()))
        language_code = language_map[selected_lang]
        
        st.markdown("---")
        st.markdown("**About this App**")
        st.info("This premium tool uses advanced AI to detect tables, scale images, and correct orientation automatically.")

    # --- Hero Section ---
    st.markdown('<p class="hero-text">Turn Invoices & Documents into Editable Files in Seconds</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-hero">Upload an image or snap a photo, and our AI will extract the text, tables, and data instantly.</p>', unsafe_allow_html=True)

    # --- How it Works (3 Steps) ---
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div class="step-card">
            <div class="step-icon">📸</div>
            <div class="step-title">1. Upload or Snap</div>
            <p style="color:#666; font-size:0.9rem;">Provide an image of your document.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col2:
        st.markdown("""
        <div class="step-card">
            <div class="step-icon">⚙️</div>
            <div class="step-title">2. AI Extracts</div>
            <p style="color:#666; font-size:0.9rem;">We auto-detect tables and orientation.</p>
        </div>
        """, unsafe_allow_html=True)
        
    with col3:
        st.markdown("""
        <div class="step-card">
            <div class="step-icon">⬇️</div>
            <div class="step-title">3. Export File</div>
            <p style="color:#666; font-size:0.9rem;">Download securely to Word or Excel.</p>
        </div>
        """, unsafe_allow_html=True)
        
    st.markdown("<br><hr>", unsafe_allow_html=True)

    # --- Main Input Area (Mobile Optimized) ---
    st.subheader("Start Extraction")
    input_method = st.radio("Choose how to provide the document:", ("📂 Upload an Image", "📷 Take a Photo"), horizontal=True)
    
    uploaded_file = None
    if input_method == "📂 Upload an Image":
        uploaded_file = st.file_uploader("Drop your image here (JPG, PNG)", type=["jpg", "jpeg", "png"])
    else:
        uploaded_file = st.camera_input("Take a clear picture of the document")

    # --- Processing & Output ---
    if uploaded_file is not None:
        with st.expander("Preview Document", expanded=False):
            image = Image.open(uploaded_file)
            st.image(image, use_column_width=True)

        st.markdown("### Result")
        with st.spinner("Extracting data... Please wait."):
            image_bytes = uploaded_file.getvalue()
            processed_bytes = preprocess_image(image_bytes)
            extracted_text = perform_ocr(processed_bytes, api_key, language_code)

        if extracted_text:
            st.success("✅ Extraction Successful!")
            
            edited_text = st.text_area("You can edit the text below before downloading:", value=extracted_text, height=300)

            st.markdown("### Export")
            
            # Buttons placed sequentially for perfect mobile rendering
            word_bytes = create_word_doc(edited_text)
            st.download_button(
                label="📄 Download as Word (.docx)",
                data=word_bytes,
                file_name="extracted_text.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
            excel_bytes = create_excel_doc(edited_text)
            st.download_button(
                label="📊 Download as Excel (.xlsx)",
                data=excel_bytes,
                file_name="extracted_text.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            st.download_button(
                label="📝 Download as Text (.txt)",
                data=edited_text,
                file_name="extracted_text.txt",
                mime="text/plain"
            )
        else:
            st.warning("⚠️ No text was extracted. Please try a clearer image.")
            
if __name__ == "__main__":
    main()
