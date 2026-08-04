import streamlit as st
from streamlit_option_menu import option_menu
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
    page_icon="🌌",
    layout="wide",
    initial_sidebar_state="expanded"
)

if "show_camera" not in st.session_state:
    st.session_state.show_camera = False

# -----------------------------------------------------------------------------
# Custom CSS (Minimalist B2B Premium UI)
# -----------------------------------------------------------------------------
st.markdown("""
<style>
/* Hide Streamlit Hamburger Menu and Footer */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}

/* Minimalist Background and Typography */
.stApp {
    background-color: #fcfdfd;
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    color: #2d3748;
}

/* Glassmorphism/Premium Container */
.glass-container {
    background: #ffffff;
    border-radius: 12px;
    border: 1px solid #e2e8f0;
    padding: 2.5rem;
    box-shadow: 0 4px 6px rgba(0,0,0,0.02), 0 1px 3px rgba(0,0,0,0.03);
    margin-bottom: 2rem;
}

/* Headings */
h1, h2, h3, h4 {
    color: #1a202c;
    font-weight: 700;
}

.main-title {
    color: #2d3748;
    text-align: center;
    font-size: 3rem !important;
    font-weight: 800;
    letter-spacing: -1px;
    margin-bottom: 0.5rem;
}

.sub-title {
    text-align: center;
    color: #718096;
    font-size: 1.2rem;
    font-weight: 400;
    margin-bottom: 3rem;
}

/* Universal Button Styling (Download Buttons) */
.stDownloadButton > button {
    background-color: #f8f9fa;
    color: #2d3748 !important;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    font-weight: 500;
    padding: 0.6rem 1.2rem;
    transition: all 0.2s ease;
    width: 100%;
}
.stDownloadButton > button:hover {
    background-color: #edf2f7;
    border-color: #cbd5e0;
    color: #1a202c !important;
}

/* Custom Take Photo Button Styling */
div[data-testid="stButton"] > button {
    background-color: #3182ce;
    color: #ffffff !important;
    border: none;
    border-radius: 8px;
    font-weight: 600;
    padding: 0.6rem 1.2rem;
    transition: all 0.2s ease;
    width: 100%;
    box-shadow: 0 2px 4px rgba(49, 130, 206, 0.2);
}
div[data-testid="stButton"] > button:hover {
    background-color: #2b6cb0;
    box-shadow: 0 4px 6px rgba(49, 130, 206, 0.3);
    transform: translateY(-1px);
}

/* Custom file uploader styling */
[data-testid="stFileUploadDropzone"] {
    background-color: #f7fafc;
    border: 1px dashed #cbd5e0;
    border-radius: 8px;
    transition: all 0.2s ease;
}
[data-testid="stFileUploadDropzone"]:hover {
    border-color: #3182ce;
    background-color: #ebf8ff;
}

/* Custom Footer */
.premium-footer {
    text-align: center;
    padding: 2rem;
    color: #a0aec0;
    font-size: 0.9rem;
    margin-top: 5rem;
    border-top: 1px solid #edf2f7;
}

/* Subpages HTML Styling */
.page-container {
    max-width: 800px;
    margin: 0 auto;
    background: #ffffff;
    padding: 3rem;
    border-radius: 12px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    border: 1px solid #e2e8f0;
}
.page-container h2 {
    color: #2d3748;
    margin-bottom: 1.5rem;
    font-weight: 700;
}
.page-container h3 {
    color: #4a5568;
    margin-top: 2rem;
    margin-bottom: 1rem;
    font-size: 1.3rem;
    font-weight: 600;
}
.page-container p {
    color: #4a5568;
    line-height: 1.7;
    margin-bottom: 1.2rem;
}

/* SEO Text Styling */
.seo-section {
    background-color: #f7fafc;
    border-radius: 12px;
    padding: 2.5rem;
    margin-top: 4rem;
    border: 1px solid #e2e8f0;
}
.seo-section h2 {
    color: #2d3748;
    font-size: 1.5rem;
    margin-bottom: 1.5rem;
}
.seo-section p {
    color: #4a5568;
    line-height: 1.7;
    margin-bottom: 1rem;
}
.seo-section ul {
    color: #4a5568;
    line-height: 1.7;
    margin-left: 1.5rem;
}
</style>

<!-- SEO Meta Tags for Google AdSense -->
<meta name="description" content="Pixel2Word is an advanced AI document intelligence platform. Automatically extract data from invoices, PDFs, and images into structured Excel and Word formats.">
<meta name="keywords" content="AI OCR, Invoice Extractor, Data Entry Automation, PDF to Word, PDF to Excel, Gemini AI OCR, Document Intelligence, Pixel2Word">
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
    st.markdown('<p class="main-title">Pixel2Word Intelligence</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Transform physical documents and digital files into intelligent, structured data instantly.</p>', unsafe_allow_html=True)
    
    # Optional Onboarding Animation
    lottie_url = "https://assets3.lottiefiles.com/packages/lf20_qp1q7mct.json"
    lottie_json = load_lottieurl(lottie_url)
    if lottie_json:
        col1, col2, col3 = st.columns([1,2,1])
        with col2:
            st_lottie(lottie_json, height=200, key="ai_animation")

    if not gemini_api_key:
        st.warning("⚠️ Configuration Required: Please set the 'GEMINI_API_KEY' environment variable.")
        st.stop()

    st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
    st.markdown("### Process Document")
    st.markdown("<p style='color:#718096; margin-bottom:1.5rem;'>Upload a file (PDF, DOCX, XLSX, CSV, JPG, PNG) or capture an image via camera.</p>", unsafe_allow_html=True)
    
    uploaded_file = None
    file_bytes = None
    is_image = False
    
    # Perfectly aligned layout for Upload and Camera
    col_upload, col_spacer, col_camera = st.columns([1, 0.1, 1])
    
    with col_upload:
        upload_input = st.file_uploader("Upload File", type=["jpg", "jpeg", "png", "webp", "pdf", "docx", "xlsx", "csv", "txt", "pptx"], label_visibility="collapsed")
        if upload_input:
            uploaded_file = upload_input
            st.session_state.show_camera = False

    with col_camera:
        # Align button vertically with the uploader dropzone
        st.write("") 
        st.write("")
        if st.button("📸 Open Camera"):
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
        st.markdown("### AI Analysis Results")
        
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
                    st.warning("Preview not available for this image format.")
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
                        st.error("Failed to extract raw text from document, or document is empty.")
                        ai_extracted_data = None
                    else:
                        ai_extracted_data = analyze_document_with_ai(extracted_raw_text, False, gemini_api_key)
                else:
                    ai_extracted_data = analyze_document_with_ai(file_bytes, True, gemini_api_key)
                
            if ai_extracted_data:
                st.success("Analysis Complete")
                edited_text = st.text_area("Extracted Intelligence", value=ai_extracted_data, height=400, label_visibility="collapsed")
                
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
        <h2>Automate Data Entry with AI Invoice & Document Extraction</h2>
        <p>
            Welcome to <strong>Pixel2Word</strong>, your premium B2B document intelligence solution. 
            Whether you are dealing with hundreds of physical receipts, complex PDF invoices, or messy Excel sheets, 
            our advanced Multimodal AI handles it instantly with enterprise-grade accuracy.
        </p>
        <p>
            <strong>How it Works:</strong> Simply upload an image, PDF, or Word document. Our AI scans the contents, 
            identifies the document type, and autonomously extracts critical entities like amounts, names, and dates. 
        </p>
        <p>
            <strong>The Benefits of Automation:</strong>
            <ul>
                <li><strong>Save Hours of Time:</strong> Eliminate tedious manual data entry and human error.</li>
                <li><strong>Universal Format Support:</strong> From JPGs and PDFs to XLSX and PPTX, Pixel2Word reads everything seamlessly.</li>
                <li><strong>Seamless Exports:</strong> Download the perfectly structured intelligence straight to Microsoft Word or Excel for immediate use.</li>
            </ul>
        </p>
    </div>
    """, unsafe_allow_html=True)

def render_privacy():
    st.markdown("""
    <div class="page-container">
        <h2>Privacy Policy</h2>
        <p><em>Effective Date: 2024</em></p>
        <p>Welcome to <strong>www.pixel2word.com</strong>. Your privacy and data security are the foundation of our platform.</p>
        
        <h3>1. Data Processing and Security</h3>
        <p>We explicitly state that all files, images, and documents uploaded to our platform are processed securely in real-time. 
        <strong>We DO NOT store, log, or train our models on your sensitive documents.</strong> Once the extraction is complete and the file is returned to you, it is immediately and permanently discarded from our active memory.</p>
        
        <h3>2. Third-Party API Usage</h3>
        <p>Our intelligent extraction relies on secure, enterprise-grade third-party APIs (such as Google Generative AI). Data sent to these APIs is subject to strict enterprise security agreements ensuring absolute confidentiality and compliance.</p>
        
        <h3>3. Analytics and Cookies</h3>
        <p>We may use standard web analytics to monitor platform performance and improve user experience. This data is strictly anonymized and never linked directly to the contents of your uploaded documents.</p>
    </div>
    """, unsafe_allow_html=True)

def render_terms():
    st.markdown("""
    <div class="page-container">
        <h2>Terms of Service</h2>
        <p>By accessing or using the <strong>www.pixel2word.com</strong> platform, you agree to comply with and be strictly bound by these Terms of Service.</p>
        
        <h3>1. Use of Service</h3>
        <p>Pixel2Word provides an advanced AI-powered document extraction tool. The service is provided "as is". While we leverage state-of-the-art AI to ensure the highest extraction accuracy, we do not guarantee 100% perfection in OCR or text extraction. Users must review extracted data before relying on it in professional, financial, or legal contexts.</p>
        
        <h3>2. Acceptable Use</h3>
        <p>You agree not to use the platform to upload illegal, explicit, or maliciously infected files. We maintain zero tolerance for abuse and reserve the right to permanently block IP addresses violating this policy without notice.</p>
        
        <h3>3. Liability Limitation</h3>
        <p>Pixel2Word and its operators shall not be held liable for any direct or indirect damages resulting from data inaccuracies, temporary service outages, or API limitations.</p>
    </div>
    """, unsafe_allow_html=True)

def render_contact():
    st.markdown("""
    <div class="page-container">
        <h2>Contact Us</h2>
        <p>Have questions, require feature integrations, or need dedicated enterprise support?</p>
        <p>We are here to help. Please reach out directly to our dedicated B2B support team.</p>
        
        <h3>Support Channels</h3>
        <p>📧 Email us at: <strong>contact@pixel2word.com</strong></p>
        
        <p><em>We aim to respond to all inquiries within 24 business hours.</em></p>
    </div>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Main Application Router
# -----------------------------------------------------------------------------
def main():
    gemini_api_key = os.environ.get("GEMINI_API_KEY")

    with st.sidebar:
        st.markdown("<div style='text-align: center; margin-bottom: 2rem;'>", unsafe_allow_html=True)
        st.image("https://cdn-icons-png.flaticon.com/512/3752/3752762.png", width=70)
        st.markdown("<h2 style='color:#2d3748; margin-top:0.5rem;'>Pixel2Word</h2>", unsafe_allow_html=True)
        st.markdown("</div>", unsafe_allow_html=True)
        
        # Premium Navigation using option_menu
        page = option_menu(
            menu_title=None,
            options=["Home / App", "Privacy Policy", "Terms of Service", "Contact Us"],
            icons=["house", "shield-lock", "file-earmark-text", "envelope"],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "transparent"},
                "icon": {"color": "#4a5568", "font-size": "1.1rem"},
                "nav-link": {
                    "font-size": "1rem", 
                    "text-align": "left", 
                    "margin": "0.2rem 0", 
                    "color": "#4a5568",
                    "border-radius": "8px"
                },
                "nav-link-selected": {
                    "background-color": "#ebf8ff", 
                    "color": "#3182ce",
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
        Transforming documents into actionable digital intelligence for the modern enterprise. 
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
