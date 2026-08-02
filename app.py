import streamlit as st
import io
import os
import requests
from PIL import Image
import docx
import pandas as pd
import google.generativeai as genai
from streamlit_lottie import st_lottie

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="VisionAI SaaS - Document Intelligence",
    page_icon="🌌",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# -----------------------------------------------------------------------------
# Custom CSS (Glassmorphism & Premium UI)
# -----------------------------------------------------------------------------
st.markdown("""
<style>
/* Background Gradient */
.stApp {
    background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}

/* Glassmorphism Container */
.glass-container {
    background: rgba(255, 255, 255, 0.7);
    backdrop-filter: blur(10px);
    -webkit-backdrop-filter: blur(10px);
    border-radius: 20px;
    border: 1px solid rgba(255, 255, 255, 0.3);
    padding: 2rem;
    box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.15);
    margin-bottom: 2rem;
}

/* Headings */
h1, h2, h3 {
    color: #1a202c;
    font-weight: 800;
}

.main-title {
    background: -webkit-linear-gradient(45deg, #FF6B6B, #556270);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    text-align: center;
    font-size: 3.5rem !important;
    margin-bottom: 0.5rem;
}

.sub-title {
    text-align: center;
    color: #4a5568;
    font-size: 1.2rem;
    margin-bottom: 2rem;
}

/* Buttons */
.stButton > button {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    color: white !important;
    border: none;
    border-radius: 12px;
    padding: 1rem 2rem;
    font-size: 1.1rem;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    width: 100%;
}
.stButton > button:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 25px rgba(0,0,0,0.3);
}

.stDownloadButton > button {
    background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
    color: white !important;
    border: none;
    border-radius: 12px;
    font-weight: 600;
    transition: all 0.3s ease;
    width: 100%;
}

/* Footer */
.footer {
    text-align: center;
    padding: 2rem;
    color: #718096;
    font-size: 0.95rem;
    margin-top: 4rem;
    border-top: 1px solid rgba(0,0,0,0.05);
}

/* Hide default file uploader text if possible, style the drag-drop zone */
[data-testid="stFileUploadDropzone"] {
    background-color: rgba(255, 255, 255, 0.5);
    border: 2px dashed #a0aec0;
    border-radius: 15px;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Core Functions
# -----------------------------------------------------------------------------
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

def analyze_document_with_vision(image_bytes, api_key):
    """
    Use Google Gemini 3.5 Flash to analyze the document.
    """
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-3.5-flash')
    
    # Prepare the image using PIL (Generative API accepts PIL Image directly)
    image = Image.open(io.BytesIO(image_bytes))
    
    prompt = """
    You are an advanced Document Intelligence AI. Analyze this image and provide:
    1. The Type of Document (e.g., Invoice, Contract, Receipt, Letter).
    2. The Language of the Document (auto-detected).
    3. Key Entities Extracted (e.g., Total Amount, Date, Names of parties involved, Invoice Number).
    4. The Full Extracted Text.
    
    Format your response clearly using Markdown headings (e.g. ### Document Type, ### Key Entities, ### Full Text).
    """

    try:
        response = model.generate_content([prompt, image])
        return response.text
    except Exception as e:
        st.error(f"AI Vision Processing Error: {str(e)}")
        return None

def create_word_doc(text):
    doc = docx.Document()
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_excel_doc(text):
    # Simple parsing to place the structured text into an Excel sheet
    lines = text.split('\n')
    data = [[line] for line in lines if line.strip()]
    df = pd.DataFrame(data, columns=["Extracted Content"])
    
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return bio.getvalue()

# -----------------------------------------------------------------------------
# Main Application
# -----------------------------------------------------------------------------
def main():
    # Retrieve Gemini API Key from environment variable strictly (No st.secrets)
    gemini_api_key = os.environ.get("GEMINI_API_KEY")

    # Hero Section
    st.markdown('<p class="main-title">VisionAI Intelligence</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Turn physical documents into intelligent digital data instantly with AI.</p>', unsafe_allow_html=True)
    
    # Visual Onboarding Animation (Lottie)
    lottie_url = "https://assets3.lottiefiles.com/packages/lf20_qp1q7mct.json"
    lottie_json = load_lottieurl(lottie_url)
    if lottie_json:
        st_lottie(lottie_json, height=250, key="ai_animation")

    st.markdown("---")

    if not gemini_api_key:
        st.warning("⚠️ Gemini API Key is missing. Please set the 'GEMINI_API_KEY' environment variable to use the AI extraction features.")
        st.stop()

    # Input Section: Side-by-side Massive CTA
    st.markdown("### 📥 Provide your Document")
    
    uploaded_file = None
    
    col1, col2 = st.columns(2, gap="large")
    with col1:
        st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
        st.markdown("#### 📂 Upload File")
        upload_input = st.file_uploader("Drag & drop or browse (JPG, PNG)", type=["jpg", "jpeg", "png"], label_visibility="collapsed")
        if upload_input:
            uploaded_file = upload_input
        st.markdown("</div>", unsafe_allow_html=True)
        
    with col2:
        st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
        st.markdown("#### 📷 Snap a Photo")
        camera_input = st.camera_input("Use your device camera", label_visibility="collapsed")
        if camera_input:
            uploaded_file = camera_input
        st.markdown("</div>", unsafe_allow_html=True)

    # Processing & Output Section
    if uploaded_file is not None:
        st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
        st.subheader("🧠 AI Analysis Results")
        
        # Two-column layout for results: Image on left, Data on right
        res_col1, res_col2 = st.columns([1, 1.5], gap="medium")
        
        with res_col1:
            image = Image.open(uploaded_file)
            st.image(image, caption="Source Document", use_container_width=True, clamp=True)
            
        with res_col2:
            with st.spinner("VisionAI is processing your document with Gemini 3.5 Flash..."):
                image_bytes = uploaded_file.getvalue()
                ai_extracted_data = analyze_document_with_vision(image_bytes, gemini_api_key)
                
            if ai_extracted_data:
                st.success("✅ Document Analyzed Successfully!")
                edited_text = st.text_area("Review and Edit Extracted Intelligence:", value=ai_extracted_data, height=400)
                
                st.markdown("### 💾 Export Data")
                btn_col1, btn_col2 = st.columns(2)
                
                with btn_col1:
                    word_bytes = create_word_doc(edited_text)
                    st.download_button(
                        label="📄 Download Word",
                        data=word_bytes,
                        file_name="vision_ai_extraction.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                with btn_col2:
                    excel_bytes = create_excel_doc(edited_text)
                    st.download_button(
                        label="📊 Download Excel",
                        data=excel_bytes,
                        file_name="vision_ai_extraction.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        st.markdown("</div>", unsafe_allow_html=True)

    # Footer
    st.markdown("""
    <div class="footer">
        <strong>VisionAI SaaS</strong><br>
        Transforming physical documents into actionable digital data instantly. 
        We save professionals hours of manual data entry through cutting-edge multimodal AI.
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
