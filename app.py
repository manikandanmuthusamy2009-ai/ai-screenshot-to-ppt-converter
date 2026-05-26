import streamlit as st
import google.generativeai as genai
from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_AUTO_SIZE
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
import io
import os
import json
import re

# --- PAGE CONFIGURATION ---
st.set_page_config(page_title="AI Resume-to-Slide Pro", page_icon="🎨", layout="wide")

st.title("🎨 Professional Resume Layout Engine")
st.markdown("Optimized for **Perfect Fit**, **Zero-Overlap**, and **Creative Presentations**.")

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
    except: return 'models/gemini-1.5-flash'

# --- 3. AI ANALYSIS (STRICT LAYOUT) ---
def analyze_resume(img, model_name):
    model = genai.GenerativeModel(model_name)
    prompt = """
    Extract resume text and group it into logical blocks. 
    Strictly identify 'sidebar' (Contact, Skills, Languages) vs 'main' (Summary, Experience, Education).
    Output RAW JSON array: [{"section": "sidebar"|"main"|"header", "type": "title"|"heading"|"body", "text": "content"}]
    Keep related bullet points in a SINGLE 'body' block to minimize gaps.
    """
    response = model.generate_content([prompt, img])
    clean = re.sub(r'```json|```', '', response.text).strip()
    return json.loads(clean)

# --- 4. PPTX BUILDER (THEMES & ADVANCED STACKING) ---
def build_tight_slide(data, paper_size, orientation, theme_choice):
    prs = Presentation()
    
    # 1. Page Dimensions (Strict Portrait Baselines)
    if paper_size == "A4": base_w, base_h = 8.27, 11.69
    elif paper_size == "Letter": base_w, base_h = 8.5, 11.0
    else: base_w, base_h = 7.5, 13.33 # "Image Size" standard portrait ratio
    
    # Apply Orientation
    if orientation == "Landscape":
        sw, sh = base_h, base_w
    else:
        sw, sh = base_w, base_h
        
    prs.slide_width = Inches(sw)
    prs.slide_height = Inches(sh)
    
    # 2. Determine which themes to generate
    if theme_choice == "4. All three options":
        themes_to_run = ["Standard Professional", "Creative Modern (Blue Accent)", "Creative Minimalist (Dark Sidebar)"]
    else:
        themes_to_run = [theme_choice]
        
    for theme in themes_to_run:
        slide = prs.slides.add_slide(prs.slide_layouts[6]) # Blank slide
        
        # 3. Margins & Dynamic Widths
        margin_top = 0.2
        margin_side = 0.3
        
        if orientation == "Portrait":
            sidebar_w = sw * 0.35 
            main_w = sw * 0.55
        else:
            sidebar_w = sw * 0.30
            main_w = sw * 0.60
            
        gap_x = 0.2
        
        # Default Trackers
        y_side = margin_top
        y_main = margin_top
        y_head = margin_top
        
        # 4. Apply Creative Backgrounds
        if theme == "Creative Modern (Blue Accent)":
            header_h = 0.8 # Reduced from 1.2 to save bottom space
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(sw), Inches(header_h))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(0, 80, 136) 
            shape.line.fill.background()
            y_side = header_h + 0.1
            y_main = header_h + 0.1
            y_head = 0.1
            
        elif theme == "Creative Minimalist (Dark Sidebar)":
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(margin_side + sidebar_w + (gap_x/2)), Inches(sh))
            shape.fill.solid()
            shape.fill.fore_color.rgb = RGBColor(43, 50, 56) 
            shape.line.fill.background()

        # Dynamic spacing controls
        v_gap_head = 0.15 # Larger space after a heading
        v_gap_body = 0.05 # Tighter space after body text

        # 5. Insert Text Blocks
        for item in data:
            sec, ctype, txt = item.get("section", "main"), item.get("type", "body"), item.get("text", "").strip()
            if not txt: continue

            # Alignment Logic
            if sec == "header": 
                tx, tw, ty = margin_side, sw-(margin_side*2), y_head
            elif sec == "sidebar": 
                tx, tw, ty = margin_side, sidebar_w, y_side
            else: 
                tx, tw, ty = margin_side + sidebar_w + gap_x, main_w, y_main

            box = slide.shapes.add_textbox(Inches(tx), Inches(ty), Inches(tw), Inches(0.2))
            tf = box.text_frame
            tf.word_wrap = True
            
            p = tf.paragraphs[0]
            p.text = txt
            p.font.name = "Calibri"
            
            # Theme Colors
            if theme == "Creative Minimalist (Dark Sidebar)" and sec == "sidebar":
                p.font.color.rgb = RGBColor(255, 255, 255)
            elif theme == "Creative Modern (Blue Accent)" and sec == "header":
                p.font.color.rgb = RGBColor(255, 255, 255)
            elif ctype == "heading" and theme == "Creative Modern (Blue Accent)":
                p.font.color.rgb = RGBColor(0, 80, 136)

            # Sizing Hierarchy
            if ctype == "title": 
                p.font.size, p.font.bold = Pt(18), True
            elif ctype == "heading": 
                p.font.size, p.font.bold = Pt(11), True
            else: 
                p.font.size = Pt(9) 

            # Advanced Height Calculation (Prevents overlapping boxes)
            chars_per_inch = 22 * (9 / p.font.size.pt) 
            chars_per_line = tw * chars_per_inch
            estimated_lines = (len(txt) / chars_per_line) + txt.count('\n') + 1.1
            real_h = estimated_lines * (p.font.size.pt * 1.25 / 72.0)
            
            next_y = ty + real_h
            
            # Apply specific spacing
            if ctype == "heading":
                next_y += v_gap_head
            else:
                next_y += v_gap_body
                
            if sec == "header": y_head = y_side = y_main = next_y
            elif sec == "sidebar": y_side = next_y
            else: y_main = next_y

    out = io.BytesIO()
    prs.save(out)
    return out.getvalue()

# --- 5. UI ---
with st.sidebar:
    st.header("1. Layout Control")
    p_size = st.selectbox("Paper Format", ["A4", "Letter", "Image Size"])
    orient = st.radio("Page Flow", ["Portrait", "Landscape"])
    
    st.header("2. Creative Theme")
    theme_choice = st.selectbox("Select Slide Design", [
        "Standard Professional", 
        "Creative Modern (Blue Accent)", 
        "Creative Minimalist (Dark Sidebar)",
        "4. All three options"
    ])

up = st.file_uploader("Upload Resume Screenshot", type=['png', 'jpg', 'jpeg'])

# --- IMAGE PREVIEW ---
if up:
    img = Image.open(up).convert('RGB')
    st.markdown("### 🖼️ Uploaded Resume Preview")
    st.image(img, use_container_width=True)
    
    # Process Button
    if st.button("Fix Layout & Generate Design", type="primary"):
        model = get_best_model()
        
        with st.spinner(f"Analyzing and applying layout..."):
            try:
                data = analyze_resume(img, model)
                
                st.session_state.ppt_buffer = build_tight_slide(data, p_size, orient, theme_choice)
                
                # Naming the file dynamically
                mode_name = "MultiTheme" if "4" in theme_choice else theme_choice.split()[0]
                st.session_state.file_name_out = f"Resume_{mode_name}_{orient}_{p_size}.pptx"
                
                st.success("✨ Presentation Generated Successfully! Ready to download below.")
                st.balloons() # Triggers audience engagement animation!
                
            except Exception as e:
                st.error(f"Generation failed: {e}")

# Download Button (Must remain outside the button loop for IDM support)
if st.session_state.ppt_buffer is not None:
    st.download_button(
        label="📥 Download Final PPTX File",
        data=st.session_state.ppt_buffer,
        file_name=st.session_state.file_name_out,
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
