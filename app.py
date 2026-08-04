import streamlit as st
from streamlit_option_menu import option_menu
from streamlit_mic_recorder import speech_to_text
import io
import os
import requests
from PIL import Image
import docx
import pandas as pd
import google.generativeai as genai
from streamlit_lottie import st_lottie
import PyPDF2
import pdfplumber
from pptx import Presentation

# -----------------------------------------------------------------------------
# Configuration & State
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Pixel2Word - AI Invoice & Document Extractor",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "show_camera" not in st.session_state:
    st.session_state.show_camera = False

# -----------------------------------------------------------------------------
# Custom CSS (Enterprise Light Theme - Stripe/Vercel Aesthetic)
# -----------------------------------------------------------------------------
st.markdown("""
<style>
/* Base Theme: Clean White & Light Gray */
.stApp {
    background-color: #fafafa;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
    color: #111827;
}

/* Hide default footer, but KEEP header/hamburger menu */
footer {visibility: hidden;}

/* Enterprise Container */
.glass-container {
    background-color: #ffffff;
    border-radius: 8px;
    border: 1px solid #e5e7eb;
    padding: 2rem;
    box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1), 0 1px 2px 0 rgba(0, 0, 0, 0.06);
    margin-bottom: 2rem;
}

/* Typography */
h1, h2, h3, h4 {
    color: #111827 !important;
    font-weight: 700;
}

.main-title {
    color: #111827;
    text-align: center;
    font-size: 2.5rem !important;
    font-weight: 800;
    letter-spacing: -0.025em;
    margin-bottom: 0.5rem;
}

.sub-title {
    text-align: center;
    color: #6b7280;
    font-size: 1.1rem;
    font-weight: 400;
    margin-bottom: 3rem;
}

/* Premium Buttons */
.stButton > button {
    background-color: #ffffff;
    color: #374151 !important;
    border: 1px solid #d1d5db;
    border-radius: 6px;
    font-size: 0.95rem;
    font-weight: 500;
    padding: 0.5rem 1rem;
    transition: all 0.2s ease;
    box-shadow: 0 1px 2px 0 rgba(0, 0, 0, 0.05);
    width: 100%;
}
.stButton > button:hover {
    background-color: #f9fafb;
    border-color: #9ca3af;
    color: #111827 !important;
}

/* Call to Action Buttons (Download, Action) */
.stDownloadButton > button, div[data-testid="stButton"] > button.primary-btn {
    background-color: #000000;
    color: #ffffff !important;
    border: none;
    border-radius: 6px;
    font-weight: 500;
    transition: all 0.2s ease;
    width: 100%;
}
.stDownloadButton > button:hover, div[data-testid="stButton"] > button.primary-btn:hover {
    background-color: #374151;
    color: #ffffff !important;
}

/* High-Contrast Text Area for Analysis Results */
.stTextArea textarea {
    background-color: #ffffff !important;
    color: #111827 !important;
    border: 1px solid #d1d5db !important;
    border-radius: 6px;
    padding: 1rem;
    font-family: 'Courier New', Courier, monospace;
    font-size: 0.95rem;
    box-shadow: inset 0 1px 2px rgba(0, 0, 0, 0.05);
}
.stTextArea textarea:focus {
    border-color: #3b82f6 !important;
    box-shadow: 0 0 0 1px #3b82f6 !important;
}

/* File Uploader */
[data-testid="stFileUploadDropzone"] {
    background-color: #ffffff;
    border: 1px dashed #d1d5db;
    border-radius: 6px;
    transition: all 0.2s ease;
}
[data-testid="stFileUploadDropzone"]:hover {
    border-color: #6b7280;
    background-color: #f9fafb;
}

/* Footer */
.premium-footer {
    text-align: center;
    padding: 2rem;
    color: #6b7280;
    font-size: 0.85rem;
    margin-top: 4rem;
    border-top: 1px solid #e5e7eb;
}

/* Subpages Container */
.page-container {
    max-width: 800px;
    margin: 0 auto;
    background: #ffffff;
    padding: 3rem;
    border-radius: 8px;
    box-shadow: 0 1px 3px 0 rgba(0, 0, 0, 0.1);
    border: 1px solid #e5e7eb;
}
.page-container h2 {
    color: #111827;
    margin-bottom: 1.5rem;
    font-weight: 700;
    border-bottom: 1px solid #e5e7eb;
    padding-bottom: 0.5rem;
}
.page-container h3 {
    color: #374151;
    margin-top: 2rem;
    margin-bottom: 1rem;
    font-size: 1.2rem;
    font-weight: 600;
}
.page-container p {
    color: #4b5563;
    line-height: 1.6;
    margin-bottom: 1.2rem;
}

/* SEO Section */
.seo-section {
    background-color: #ffffff;
    border-radius: 8px;
    padding: 2.5rem;
    margin-top: 4rem;
    border: 1px solid #e5e7eb;
}
.seo-section h2 {
    color: #111827;
    font-size: 1.3rem;
    margin-bottom: 1rem;
}
.seo-section p, .seo-section ul {
    color: #4b5563;
    line-height: 1.6;
}
</style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Core Extraction Functions
# -----------------------------------------------------------------------------
def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None

def extract_text_from_pdf(file_bytes):
    text = ""
    try:
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            for page in pdf.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
    except Exception:
        try:
            reader = PyPDF2.PdfReader(io.BytesIO(file_bytes))
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
        except Exception as e2:
            st.error(f"Failed to extract PDF: {str(e2)}")
    return text

def extract_text_from_docx(file_bytes):
    text = ""
    try:
        doc = docx.Document(io.BytesIO(file_bytes))
        for para in doc.paragraphs:
            text += para.text + "\n"
    except Exception as e:
        st.error(f"Failed to extract Word document: {str(e)}")
    return text

def extract_text_from_excel(file_bytes):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes))
        return df.to_string()
    except Exception as e:
        st.error(f"Failed to extract Excel document: {str(e)}")
        return ""

def extract_text_from_csv(file_bytes):
    try:
        df = pd.read_csv(io.BytesIO(file_bytes))
        return df.to_string()
    except Exception as e:
        st.error(f"Failed to extract CSV document: {str(e)}")
        return ""

def extract_text_from_pptx(file_bytes):
    text = ""
    try:
        prs = Presentation(io.BytesIO(file_bytes))
        for slide in prs.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    text += shape.text + "\n"
    except Exception as e:
        st.error(f"Failed to extract PowerPoint document: {str(e)}")
    return text

def analyze_document_with_ai(file_content, is_image, api_key):
    genai.configure(api_key=api_key)
    
    prompt = """
    You are an advanced Document Intelligence AI for Pixel2Word. Analyze this input and provide:
    1. The Type of Document (e.g., Invoice, Contract, Receipt, Letter, Spreadsheet).
    2. The Language of the Document (auto-detected).
    3. Key Entities Extracted (e.g., Total Amount, Date, Names of parties involved, Invoice Number, Structured Data).
    4. The Full Extracted Text.
    
    Format your response clearly using Markdown headings (e.g. ### Document Type, ### Key Entities, ### Full Text).
    """

    try:
        model = genai.GenerativeModel('gemini-3.5-flash')
        if is_image:
            image = Image.open(io.BytesIO(file_content))
            response = model.generate_content([prompt, image])
        else:
            full_prompt = f"{prompt}\n\nDocument Text:\n{file_content}"
            response = model.generate_content(full_prompt)
        return response.text
    except Exception as e:
        st.error(f"AI Processing Error: {str(e)}")
        return None

def create_word_doc(text):
    doc = docx.Document()
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_excel_doc(text):
    lines = text.split('\n')
    data = [[line] for line in lines if line.strip()]
    df = pd.DataFrame(data, columns=["Extracted Content"])
    
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return bio.getvalue()

# -----------------------------------------------------------------------------
# Pages
# -----------------------------------------------------------------------------
def render_home(gemini_api_key):
    st.markdown('<p class="main-title">Pixel2Word</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Enterprise document intelligence. Extract structured data instantly.</p>', unsafe_allow_html=True)
    
    lottie_url = "https://assets3.lottiefiles.com/packages/lf20_qp1q7mct.json"
    lottie_json = load_lottieurl(lottie_url)
    if lottie_json:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st_lottie(lottie_json, height=150, key="ai_animation")

    if not gemini_api_key:
        st.warning("⚠️ Configuration Required: Please set the 'GEMINI_API_KEY' environment variable.")
        st.stop()

    st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
    st.markdown("### Process Document")
    
    uploaded_file = None
    file_bytes = None
    is_image = False
    
    col_upload, col_spacer, col_camera = st.columns([1, 0.1, 1])
    
    with col_upload:
        upload_input = st.file_uploader("Upload File (PDF, DOCX, XLSX, Images)", type=["jpg", "jpeg", "png", "webp", "pdf", "docx", "xlsx", "csv", "txt", "pptx"], label_visibility="collapsed")
        if upload_input:
            uploaded_file = upload_input
            st.session_state.show_camera = False

    with col_camera:
        st.write("") 
        st.write("")
        if st.button("📷 Take Photo", key="btn_camera"):
            st.session_state.show_camera = not st.session_state.show_camera
            
        if st.session_state.show_camera:
            st.markdown("<br>", unsafe_allow_html=True)
            camera_input = st.camera_input("Capture Document", label_visibility="collapsed")
            if camera_input:
                uploaded_file = camera_input
                
    st.markdown("</div>", unsafe_allow_html=True)

    # --- Processing & Output Section ---
    if uploaded_file is not None:
        st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
        st.markdown("### Analysis Results")
        
        file_bytes = uploaded_file.getvalue()
        file_extension = uploaded_file.name.split('.')[-1].lower() if uploaded_file.name else 'jpg'
        
        image_formats = ['jpg', 'jpeg', 'png', 'webp']
        if file_extension in image_formats:
            is_image = True
        
        res_col1, res_col2 = st.columns([1, 1.5], gap="medium")
        
        with res_col1:
            if is_image:
                try:
                    image = Image.open(uploaded_file)
                    st.image(image, caption="Source Document", use_container_width=True, clamp=True)
                except:
                    st.warning("Preview not available.")
            else:
                st.info(f"📄 Document Loaded: {uploaded_file.name}")
            
        with res_col2:
            with st.spinner("Processing document..."):
                extracted_raw_text = ""
                
                if not is_image:
                    if file_extension == 'pdf':
                        extracted_raw_text = extract_text_from_pdf(file_bytes)
                    elif file_extension == 'docx':
                        extracted_raw_text = extract_text_from_docx(file_bytes)
                    elif file_extension == 'xlsx':
                        extracted_raw_text = extract_text_from_excel(file_bytes)
                    elif file_extension == 'csv':
                        extracted_raw_text = extract_text_from_csv(file_bytes)
                    elif file_extension == 'pptx':
                        extracted_raw_text = extract_text_from_pptx(file_bytes)
                    elif file_extension == 'txt':
                        extracted_raw_text = file_bytes.decode('utf-8', errors='ignore')
                    
                    if not extracted_raw_text.strip():
                        st.error("Failed to extract raw text from document.")
                        ai_extracted_data = None
                    else:
                        ai_extracted_data = analyze_document_with_ai(extracted_raw_text, False, gemini_api_key)
                else:
                    ai_extracted_data = analyze_document_with_ai(file_bytes, True, gemini_api_key)
                
            if ai_extracted_data:
                st.success("Analysis Complete")
                # Text area styled by CSS to be high-contrast light background
                edited_text = st.text_area("Extracted Intelligence", value=ai_extracted_data, height=350, label_visibility="collapsed")
                
                # --- Voice-to-Text Feature ---
                st.markdown("<p style='font-size: 0.9rem; color: #4b5563; margin-top: 0.5rem;'>🎤 Click to dictate additional notes:</p>", unsafe_allow_html=True)
                dictated_text = speech_to_text(language='en', use_container_width=True, just_once=True, key='STT')
                if dictated_text:
                    edited_text = edited_text + "\n\n### Dictated Notes\n" + dictated_text
                    st.success("Notes appended!")

                st.markdown("<br><b>Export Data</b>", unsafe_allow_html=True)
                btn_col1, btn_col2 = st.columns(2)
                
                with btn_col1:
                    word_bytes = create_word_doc(edited_text)
                    st.download_button(
                        label="📄 Download Word",
                        data=word_bytes,
                        file_name="pixel2word_extraction.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                with btn_col2:
                    excel_bytes = create_excel_doc(edited_text)
                    st.download_button(
                        label="📊 Download Excel",
                        data=excel_bytes,
                        file_name="pixel2word_extraction.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        st.markdown("</div>", unsafe_allow_html=True)

    # --- SEO Optimized Content ---
    st.markdown("""
    <div class="seo-section">
        <h2>Enterprise Document Intelligence</h2>
        <p>
            Welcome to <strong>Pixel2Word</strong>, your B2B platform for automating document extraction. 
            Process receipts, complex PDF invoices, and spreadsheets instantly using advanced AI.
        </p>
        <p>
            <strong>How it Works:</strong> Upload any document. Our AI framework categorizes the file, 
            detects the language, and autonomously extracts critical business entities (amounts, dates, items).
        </p>
        <p>
            <strong>Key Advantages:</strong>
            <ul>
                <li><strong>Eliminate Manual Entry:</strong> Save hours of administrative overhead.</li>
                <li><strong>Universal Format Support:</strong> Ingest PDFs, Word docs, Excel sheets, and Images.</li>
                <li><strong>Instant Export:</strong> Download structured data directly to Microsoft Office formats.</li>
            </ul>
        </p>
    </div>
    """, unsafe_allow_html=True)

def render_privacy():
    st.markdown("""
    <div class="page-container">
        <h2>Privacy Policy</h2>
        <p><em>Effective Date: 2024</em></p>
        <p>Welcome to <strong>www.pixel2word.com</strong>. Enterprise data security is our highest priority.</p>
        
        <h3>1. Data Processing and Security</h3>
        <p>All files, images, and documents uploaded to our platform are processed securely in real-time. 
        <strong>We DO NOT store, log, or train models on your documents.</strong> Following extraction, files are immediately discarded from active memory.</p>
        
        <h3>2. Third-Party Infrastructure</h3>
        <p>Our extraction pipeline relies on secure, enterprise-grade APIs (Google Generative AI). Data transmitted is subject to stringent enterprise security agreements ensuring absolute confidentiality.</p>
        
        <h3>3. Analytics</h3>
        <p>We utilize standard web analytics strictly to monitor platform health. This data remains anonymized and is never linked to the contents of uploaded documents.</p>
    </div>
    """, unsafe_allow_html=True)

def render_terms():
    st.markdown("""
    <div class="page-container">
        <h2>Terms of Service</h2>
        <p>By accessing <strong>www.pixel2word.com</strong>, you agree to comply with these Terms of Service.</p>
        
        <h3>1. Use of Service</h3>
        <p>Pixel2Word provides an advanced AI-powered document extraction tool provided "as is". While we leverage state-of-the-art AI to ensure accuracy, we do not guarantee 100% perfection. Users must verify extracted data prior to professional reliance.</p>
        
        <h3>2. Acceptable Use</h3>
        <p>You agree not to use the platform to upload illegal, explicit, or malicious files. We reserve the right to permanently revoke access for policy violations.</p>
        
        <h3>3. Liability Limitation</h3>
        <p>Pixel2Word and its operators shall not be held liable for direct or indirect damages resulting from data inaccuracies, service outages, or infrastructure limitations.</p>
    </div>
    """, unsafe_allow_html=True)

def render_contact():
    st.markdown("""
    <div class="page-container">
        <h2>Contact Support</h2>
        <p>Have technical questions, require API access, or need enterprise deployment support?</p>
        <p>Reach out directly to our B2B integration team.</p>
        
        <h3>Email Support</h3>
        <p>📧 <strong>contact@pixel2word.com</strong></p>
        
        <p><em>Standard SLA: Responses within 24 business hours.</em></p>
    </div>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Main Application Router
# -----------------------------------------------------------------------------
def main():
    gemini_api_key = os.environ.get("GEMINI_API_KEY")

    with st.sidebar:
        st.markdown("<h2 style='text-align: center; color:#111827; margin-bottom: 2rem;'>Pixel2Word</h2>", unsafe_allow_html=True)
        
        # Premium Navigation using option_menu
        page = option_menu(
            menu_title=None,
            options=["Home / App", "Privacy Policy", "Terms of Service", "Contact Us"],
            icons=["house", "shield-lock", "file-earmark-text", "envelope"],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "transparent", "border": "none"},
                "icon": {"color": "#6b7280", "font-size": "1.1rem"},
                "nav-link": {
                    "font-size": "0.95rem", 
                    "text-align": "left", 
                    "margin": "0.2rem 0", 
                    "color": "#374151",
                    "border-radius": "6px",
                    "font-weight": "500"
                },
                "nav-link-selected": {
                    "background-color": "#f3f4f6", 
                    "color": "#111827",
                    "font-weight": "600"
                },
            }
        )

    # Route Rendering
    if page == "Home / App":
        render_home(gemini_api_key)
    elif page == "Privacy Policy":
        render_privacy()
    elif page == "Terms of Service":
        render_terms()
    elif page == "Contact Us":
        render_contact()

    # Premium Global Footer
    st.markdown("""
    <div class="premium-footer">
        <strong>© 2024 Pixel2Word. All rights reserved.</strong><br>
        Enterprise Document Intelligence
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
