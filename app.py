import streamlit as st
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
# Custom CSS (Glassmorphism, Premium UI, & SEO Meta)
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
h1, h2, h3, h4 {
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
    padding: 0.8rem 1.5rem;
    font-size: 1rem;
    font-weight: 600;
    transition: all 0.3s ease;
    box-shadow: 0 4px 15px rgba(0,0,0,0.2);
    width: 100%;
}
.stButton > button:hover {
    transform: translateY(-2px);
    box-shadow: 0 6px 20px rgba(0,0,0,0.3);
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

/* Custom file uploader styling */
[data-testid="stFileUploadDropzone"] {
    background-color: rgba(255, 255, 255, 0.5);
    border: 2px dashed #a0aec0;
    border-radius: 15px;
}

/* SEO Text Styling */
.seo-section {
    background-color: rgba(255,255,255,0.85);
    border-radius: 15px;
    padding: 2rem;
    margin-top: 3rem;
    color: #333;
    line-height: 1.6;
}
.seo-section h2 {
    color: #2c5282;
    border-bottom: 2px solid #e2e8f0;
    padding-bottom: 0.5rem;
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
    st.markdown('<p class="main-title">Pixel2Word AI Intelligence</p>', unsafe_allow_html=True)
    st.markdown('<p class="sub-title">Turn any physical or digital document into intelligent data instantly with AI.</p>', unsafe_allow_html=True)
    
    lottie_url = "https://assets3.lottiefiles.com/packages/lf20_qp1q7mct.json"
    lottie_json = load_lottieurl(lottie_url)
    if lottie_json:
        st_lottie(lottie_json, height=250, key="ai_animation")

    st.markdown("---")

    if not gemini_api_key:
        st.warning("⚠️ Gemini API Key is missing. Please set the 'GEMINI_API_KEY' environment variable to use the AI extraction features.")
        st.stop()

    # --- Sleek Input Section ---
    st.markdown("### 📥 Provide your Document")
    st.markdown("Supported formats: Images (PNG, JPG, WEBP) & Documents (PDF, DOCX, XLSX, CSV, PPTX, TXT)")
    
    uploaded_file = None
    file_bytes = None
    is_image = False
    
    st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
    
    col_upload, col_camera = st.columns(2)
    with col_upload:
        upload_input = st.file_uploader("📂 Upload File", type=["jpg", "jpeg", "png", "webp", "pdf", "docx", "xlsx", "csv", "txt", "pptx"])
        if upload_input:
            uploaded_file = upload_input
            st.session_state.show_camera = False

    with col_camera:
        st.markdown("<br>", unsafe_allow_html=True) 
        if st.button("📷 Take Photo"):
            st.session_state.show_camera = not st.session_state.show_camera
            
        if st.session_state.show_camera:
            camera_input = st.camera_input("Capture Document")
            if camera_input:
                uploaded_file = camera_input
                
    st.markdown("</div>", unsafe_allow_html=True)

    # --- Processing & Output Section ---
    if uploaded_file is not None:
        st.markdown("<div class='glass-container'>", unsafe_allow_html=True)
        st.subheader("🧠 AI Analysis Results")
        
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
                st.info(f"📄 Document Uploaded: {uploaded_file.name}")
            
        with res_col2:
            with st.spinner("VisionAI is processing your document..."):
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
                st.success("✅ Document Analyzed Successfully!")
                edited_text = st.text_area("Review and Edit Extracted Intelligence:", value=ai_extracted_data, height=400)
                
                st.markdown("### 💾 Export Data")
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
            Welcome to <strong>Pixel2Word</strong>, your ultimate document intelligence solution. 
            Whether you are dealing with hundreds of physical receipts, complex PDF invoices, or messy Excel sheets, 
            our advanced Multimodal AI handles it instantly. 
        </p>
        <p>
            <strong>How it Works:</strong> Simply upload an image, PDF, or Word document. Our Gemini-powered AI scans the contents, 
            identifies the document type, and autonomously extracts critical entities like amounts, names, and dates. 
        </p>
        <p>
            <strong>The Benefits of Automation:</strong>
            <ul>
                <li><strong>Save Hours of Time:</strong> Eliminate tedious manual data entry and human error.</li>
                <li><strong>Universal Format Support:</strong> From JPGs and PDFs to XLSX and PPTX, Pixel2Word reads everything.</li>
                <li><strong>Seamless Exports:</strong> Download the perfectly structured and extracted intelligence straight to Microsoft Word or Excel.</li>
            </ul>
        </p>
    </div>
    """, unsafe_allow_html=True)

def render_privacy():
    st.title("🔒 Privacy Policy")
    st.markdown("""
    <div class="glass-container">
        <h3>Effective Date: 2024</h3>
        <p>Welcome to <strong>www.pixel2word.com</strong>. Your privacy is critically important to us.</p>
        
        <h4>1. Data Processing and Security</h4>
        <p>We explicitly state that all files, images, and documents uploaded to our platform are processed securely in real-time. 
        <strong>We DO NOT store, log, or train our models on your sensitive documents.</strong> Once the extraction is complete and the file is returned to you, it is immediately discarded from our active memory.</p>
        
        <h4>2. Third-Party API Usage</h4>
        <p>Our intelligent extraction relies on secure third-party APIs (such as Google Generative AI). Data sent to these APIs is subject to strict enterprise security agreements ensuring confidentiality.</p>
        
        <h4>3. Analytics and Cookies</h4>
        <p>We may use standard web analytics to improve user experience. This data is anonymized and never linked directly to the contents of your uploaded documents.</p>
    </div>
    """, unsafe_allow_html=True)

def render_terms():
    st.title("📜 Terms of Service")
    st.markdown("""
    <div class="glass-container">
        <h3>Welcome to Pixel2Word</h3>
        <p>By accessing or using <strong>www.pixel2word.com</strong>, you agree to comply with and be bound by these Terms of Service.</p>
        
        <h4>1. Use of Service</h4>
        <p>Pixel2Word provides an AI-powered document extraction tool. The service is provided "as is," and while we strive for the highest accuracy, we do not guarantee 100% perfection in OCR or AI text extraction. Users are expected to review extracted data before using it in professional or legal contexts.</p>
        
        <h4>2. Acceptable Use</h4>
        <p>You agree not to use the platform to upload illegal, explicit, or maliciously infected files. We reserve the right to block IP addresses violating this policy.</p>
        
        <h4>3. Liability</h4>
        <p>Pixel2Word shall not be held liable for any damages resulting from data inaccuracies, temporary service outages, or API limitations.</p>
    </div>
    """, unsafe_allow_html=True)

def render_contact():
    st.title("✉️ Contact Us")
    st.markdown("""
    <div class="glass-container">
        <p>Have questions, feature requests, or need enterprise support?</p>
        <p>We would love to hear from you. Please reach out to our dedicated support team.</p>
        
        <h4>Support Email</h4>
        <p>📧 <strong>contact@pixel2word.com</strong></p>
        
        <p><em>We aim to respond to all inquiries within 24-48 business hours.</em></p>
    </div>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Main Application Router
# -----------------------------------------------------------------------------
def main():
    gemini_api_key = os.environ.get("GEMINI_API_KEY")

    with st.sidebar:
        st.image("https://cdn-icons-png.flaticon.com/512/3752/3752762.png", width=60)
        st.title("Navigation")
        
        page = st.radio("Go to:", [
            "🏠 Home / App", 
            "🔒 Privacy Policy", 
            "📜 Terms of Service", 
            "✉️ Contact Us"
        ])
        
        st.markdown("---")
        st.markdown("**Pixel2Word SaaS**")
        st.info("Automate your document intelligence.")

    if page == "🏠 Home / App":
        render_home(gemini_api_key)
    elif page == "🔒 Privacy Policy":
        render_privacy()
    elif page == "📜 Terms of Service":
        render_terms()
    elif page == "✉️ Contact Us":
        render_contact()

    # Shared Footer across all pages
    st.markdown("""
    <div class="footer">
        <strong>© 2024 Pixel2Word. All rights reserved.</strong><br>
        Transforming physical documents and digital files into actionable data instantly. 
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
