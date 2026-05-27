import streamlit as st
import google.generativeai as genai
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import io
import os
import json

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="AI Resume-to-Slide Pro", page_icon="🎨", layout="wide")

st.title("🎨 Professional Resume Layout Engine")
st.markdown("Optimized for **Streamlit Cloud Deployment**, **Tight Bullet Indents**, and **Zero Text Omissions**.")

# --- 1. INITIALIZE GEMINI API ---
api_key = os.environ.get("GEMINI_API_KEY")
if not api_key:
    st.sidebar.error("❌ API Key Missing. Set GEMINI_API_KEY in environment.")
else:
    st.sidebar.success("🔑 API Connected")
    genai.configure(api_key=api_key)

# --- SESSION STATE INITIALIZATION ---
if "ppt_buffer" not in st.session_state:
    st.session_state.ppt_buffer = None
if "file_name_out" not in st.session_state:
    st.session_state.file_name_out = "Resume.pptx"

# --- 2. AUTO-MODEL DETECTION ---
def get_best_model():
    try:
        available = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        for m in ['models/gemini-1.5-flash', 'models/gemini-1.5-flash-latest', 'models/gemini-1.5-pro']:
            if m in available: return m
        return available[0]
    except: 
        return 'models/gemini-1.5-flash'

# --- 3. AI STRUCTURE EXTRACTION WITH STRUCTURAL JSON ENFORCEMENT ---
def analyze_resume(img, model_name):
    model = genai.GenerativeModel(model_name)
    prompt = """
    Extract ALL text elements from this resume image. Do not summarize, truncate, or omit any section.
    Ensure you capture 'AWARDS', 'HONORS', or 'ACHIEVEMENTS' completely.
    
    Organize data strictly into the following JSON schema:
    {
      "name": "Full Name String",
      "title": "Professional Title String",
      "contact": ["Phone", "Email", "LinkedIn URL", "Location"],
      "about_me": ["Sentence line 1", "Sentence line 2"],
      "skills": ["Skill 1", "Skill 2"],
      "tools": ["Tool 1", "Tool 2"],
      "education": [{"deg": "Degree/Major", "inst": "University/Institution"}],
      "experience": [
        {
          "company": "Company Name",
          "location": "City/Country",
          "role": "Job Title",
          "period": "Years Active",
          "bullets": ["Achievement 1", "Achievement 2"]
        }
      ],
      "awards": ["Award/Honor 1", "Award/Honor 2"]
    }
    """
    response = model.generate_content(
        [prompt, img],
        generation_config={"response_mime_type": "application/json"}
    )
    return json.loads(response.text)

# --- 4. PPTX BUILDER ENGINE ---
def build_tight_slide(data, paper_size, orientation, theme_choice):
    prs = Presentation()
    
    # Universal Color Definitions
    COLOR_TEXT_DARK = RGBColor(30, 30, 30)
    COLOR_MUTED_GRAY = RGBColor(110, 110, 110)

    # Configure Base Paper Dimensions
    if paper_size == "A4 Page Size": base_w, base_h = 8.27, 11.69
    elif paper_size == "Letter": base_w, base_h = 8.5, 11.0
    elif paper_size == "Image Size": base_w, base_h = 7.5, 13.33
    else: base_w, base_h = 13.333, 7.5 # Standard Widescreen (16:9)
    
    if orientation == "Landscape":
        sw, sh = max(base_w, base_h), min(base_w, base_h)
        left_col_width = Inches(sw * 0.32)
        right_col_left = Inches(sw * 0.35)
        right_col_width = Inches(sw * 0.61)
    else: # Portrait
        sw, sh = min(base_w, base_h), max(base_w, base_h)
        left_col_width = Inches(sw * 0.34)
        right_col_left = Inches(sw * 0.38)
        right_col_width = Inches(sw * 0.58)
        
    prs.slide_width = Inches(sw)
    prs.slide_height = Inches(sh)
    
    themes_to_run = ["Classic Blue Banner", "Minimal Stark Dark", "Modern Slate"] if theme_choice == "All three options" else [theme_choice]
        
    for theme in themes_to_run:
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        
        # Set dynamic theme parameters inside the loop mapping
        if theme == "Minimal Stark Dark":
            COLOR_PRIMARY = RGBColor(40, 44, 52)
            COLOR_SIDEBAR_BG = RGBColor(230, 233, 238)
        elif theme == "Modern Slate":
            COLOR_PRIMARY = RGBColor(74, 85, 104)
            COLOR_SIDEBAR_BG = RGBColor(226, 232, 240)
        else: # Classic Blue Banner
            COLOR_PRIMARY = RGBColor(0, 90, 160)
            COLOR_SIDEBAR_BG = RGBColor(242, 244, 247)
        
        # Draw LHS Sidebar
        sidebar_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), left_col_width, prs.slide_height)
        sidebar_bg.fill.solid()
        sidebar_bg.fill.fore_color.rgb = COLOR_SIDEBAR_BG
        sidebar_bg.line.fill.background()

        # Draw Top Header Banner
        banner_height = Inches(1.4)
        header_banner = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, banner_height)
        header_banner.fill.solid()
        header_banner.fill.fore_color.rgb = COLOR_PRIMARY
        header_banner.line.fill.background()

        # SPLIT-HEADER: Name & Title (Top Left)
        name_box = slide.shapes.add_textbox(Inches(0.2), Inches(0), left_col_width - Inches(0.2), banner_height)
        name_tf = name_box.text_frame
        name_tf.word_wrap = True
        name_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        p_name = name_tf.paragraphs[0]
        p_name.text = data.get("name", "PROFILE NAME").upper()
        p_name.font.bold = True
        p_name.font.size = Pt(22)
        p_name.font.color.rgb = RGBColor(255, 255, 255)
        
        p_title = name_tf.add_paragraph()
        p_title.text = data.get("title", "").upper()
        p_title.font.bold = True
        p_title.font.size = Pt(11)
        p_title.font.color.rgb = RGBColor(220, 235, 255)
        p_title.space_before = Pt(2)

        # SPLIT-HEADER: Contact Details (Top Right Aligned + Font Size Increased)
        contact_box = slide.shapes.add_textbox(right_col_left, Inches(0), right_col_width, banner_height)
        contact_tf = contact_box.text_frame
        contact_tf.word_wrap = True
        contact_tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        
        contact_items = data.get("contact", [])
        if contact_items:
            p_contact = contact_tf.paragraphs[0]
            p_contact.alignment = PP_ALIGN.RIGHT
            p_contact.text = "  |  ".join(contact_items)
            p_contact.font.size = Pt(12)  
            p_contact.font.color.rgb = RGBColor(255, 255, 255)

        # Fixed Tight Spacing Rule: Uses spaces instead of tabs to completely bypass large tab gaps safely
        def format_tight_bullet(p, text):
            p.text = f"•  {text}"
            p.left_indent = Inches(0.25)          # Left margin block alignment
            p.first_line_indent = Inches(-0.15)   # Keeps the bullet snug and close to the text
            p.font.size = Pt(11)
            p.font.color.rgb = COLOR_TEXT_DARK

        # --- LHS COLUMN CONTENT FILL ---
        lhs_box = slide.shapes.add_textbox(Inches(0.2), banner_height + Inches(0.2), left_col_width - Inches(0.35), prs.slide_height - banner_height - Inches(0.4))
        lhs_tf = lhs_box.text_frame
        lhs_tf.word_wrap = True
        lhs_tf.margin_top = Inches(0)
        lhs_tf.margin_left = Inches(0)

        # About Me (Bullet Form Enforced)
        if "about_me" in data and data["about_me"]:
            p_head = lhs_tf.paragraphs[0] if not lhs_tf.text else lhs_tf.add_paragraph()
            p_head.text = "ABOUT ME"
            p_head.font.bold = True
            p_head.font.size = Pt(15)
            p_head.font.color.rgb = COLOR_PRIMARY
            p_head.space_before = Pt(8)
            p_head.space_after = Pt(4)
            
            for sentence in data["about_me"]:
                p = lhs_tf.add_paragraph()
                format_tight_bullet(p, sentence)
                p.space_after = Pt(3)

        # Core Skills
        if "skills" in data and data["skills"]:
            p_head = lhs_tf.add_paragraph()
            p_head.text = "CORE SKILLS"
            p_head.font.bold = True
            p_head.font.size = Pt(15)
            p_head.font.color.rgb = COLOR_PRIMARY
            p_head.space_before = Pt(12)
            p_head.space_after = Pt(4)
            
            for skill in data["skills"]:
                p = lhs_tf.add_paragraph()
                format_tight_bullet(p, skill)
                p.space_after = Pt(3)

        # Tools
        if "tools" in data and data["tools"]:
            p_head = lhs_tf.add_paragraph()
            p_head.text = "TOOLS"
            p_head.font.bold = True
            p_head.font.size = Pt(15)
            p_head.font.color.rgb = COLOR_PRIMARY
            p_head.space_before = Pt(12)
            p_head.space_after = Pt(4)
            
            p = lhs_tf.add_paragraph()
            p.text = ", ".join(data["tools"]) if isinstance(data["tools"], list) else data["tools"]
            p.font.size = Pt(11)
            p.font.color.rgb = COLOR_TEXT_DARK
            p.space_after = Pt(3)

        # Education (Stacked: Bold Degree Title Above Institution)
        if "education" in data and data["education"]:
            p_head = lhs_tf.add_paragraph()
            p_head.text = "EDUCATION"
            p_head.font.bold = True
            p_head.font.size = Pt(15)
            p_head.font.color.rgb = COLOR_PRIMARY
            p_head.space_before = Pt(12)
            p_head.space_after = Pt(4)
            
            for edu in data["education"]:
                p_deg = lhs_tf.add_paragraph()
                p_deg.text = edu.get('deg', '')
                p_deg.font.bold = True
                p_deg.font.size = Pt(11.5)
                p_deg.font.color.rgb = COLOR_TEXT_DARK
                p_deg.space_before = Pt(4)
                
                p_inst = lhs_tf.add_paragraph()
                p_inst.text = edu.get('inst', '')
                p_inst.font.size = Pt(10.5)
                p_inst.font.color.rgb = COLOR_MUTED_GRAY
                p_inst.space_after = Pt(4)

        # --- RHS COLUMN CONTENT FILL ---
        rhs_box = slide.shapes.add_textbox(right_col_left, banner_height + Inches(0.2), right_col_width, prs.slide_height - banner_height - Inches(0.4))
        rhs_tf = rhs_box.text_frame
        rhs_tf.word_wrap = True
        rhs_tf.margin_top = Inches(0)
        rhs_tf.margin_left = Inches(0)

        # Experience Section
        if "experience" in data and data["experience"]:
            p_exp_head = rhs_tf.paragraphs[0] if not rhs_tf.text else rhs_tf.add_paragraph()
            p_exp_head.text = "EXPERIENCE"
            p_exp_head.font.bold = True
            p_exp_head.font.size = Pt(16)
            p_exp_head.font.color.rgb = COLOR_PRIMARY
            p_exp_head.space_after = Pt(4)
            
            for job in data["experience"]:
                p_job = rhs_tf.add_paragraph()
                p_job.space_before = Pt(8)
                p_job.space_after = Pt(4)
                
                # Company Name (Blue & Bold)
                r_comp = p_job.add_run()
                r_comp.text = f"{job.get('company', '')} ({job.get('location', '')}) "
                r_comp.font.bold = True
                r_comp.font.size = Pt(13)
                r_comp.font.color.rgb = COLOR_PRIMARY
                
                r_div = p_job.add_run()
                r_div.text = " | "
                r_div.font.size = Pt(13)
                r_div.font.color.rgb = COLOR_MUTED_GRAY
                
                # Role Title (Blue, Bold, and Italicized)
                r_role = p_job.add_run()
                r_role.text = f" {job.get('role', '')} "
                r_role.font.bold = True
                r_role.font.italic = True
                r_role.font.size = Pt(13)
                r_role.font.color.rgb = COLOR_PRIMARY
                
                r_time = p_job.add_run()
                r_time.text = f" — {job.get('period', '')}"
                r_time.font.size = Pt(11.5)
                r_time.font.color.rgb = COLOR_MUTED_GRAY

                for bullet in job.get("bullets", []):
                    p_blt = rhs_tf.add_paragraph()
                    format_tight_bullet(p_blt, bullet)
                    p_blt.space_after = Pt(2.5)

        # Awards & Achievements Section
        if "awards" in data and data["awards"]:
            p_awd_head = rhs_tf.add_paragraph()
            p_awd_head.text = "AWARDS & ACHIEVEMENTS"
            p_awd_head.font.bold = True
            p_awd_head.font.size = Pt(16)
            p_awd_head.font.color.rgb = COLOR_PRIMARY
            p_awd_head.space_before = Pt(14)
            p_awd_head.space_after = Pt(4)
            
            for award in data["awards"]:
                p_awd = rhs_tf.add_paragraph()
                format_tight_bullet(p_awd, award)
                p_awd.space_after = Pt(2.5)

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

# --- 5. SIDEBAR CONTROLS ---
with st.sidebar:
    st.header("1. Layout Control")
    p_size = st.selectbox("Paper Format", ["Standard Widescreen (16:9)", "A4 Page Size", "Letter", "Image Size"], index=1)
    orient = st.radio("Page Flow", ["Portrait", "Landscape"], index=0)
    
    st.header("2. Creative Theme")
    theme_choice = st.selectbox("Select Slide Design", [
        "Classic Blue Banner", 
        "Minimal Stark Dark", 
        "Modern Slate",
        "All three options"
    ], index=0)

# --- 6. MAIN WORKFLOW CONTROLLER ---
up = st.file_uploader("Upload Resume Screenshot", type=['png', 'jpg', 'jpeg'])

if up:
    img = Image.open(up).convert('RGB')
    st.markdown("### 🖼️ Uploaded Resume Preview")
    st.image(img, use_container_width=True)
    
    if st.button("Fix Layout & Generate Design", type="primary"):
        model = get_best_model()
        
        with st.spinner("Processing document data structure and executing strict structural layout tweaks..."):
            try:
                parsed_data = analyze_resume(img, model)
                st.session_state.ppt_buffer = build_tight_slide(parsed_data, p_size, orient, theme_choice)
                
                mode_name = "MultiTheme" if theme_choice == "All three options" else theme_choice.split()[0]
                st.session_state.file_name_out = f"Resume_{mode_name}_{orient}_{p_size.split()[0]}.pptx"
                
                st.success("✨ Production template generated completely! Ready for production deployment.")
                st.balloons()
                
            except Exception as e:
                st.error(f"Generation failed: {e}")

# Persistent Downloader Segment
if st.session_state.ppt_buffer is not None:
    st.markdown("---")
    st.download_button(
        label="📥 Download Final PPTX File",
        data=st.session_state.ppt_buffer,
        file_name=st.session_state.file_name_out,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )