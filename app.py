import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import os
import re
import time
import altair as alt
import feedparser
import urllib.parse
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN

# 1. Page Configuration
st.set_page_config(page_title="RivalRadar", layout="wide")

# --- CUSTOM CSS FOR SANY LOOK & RADAR ---
st.markdown("""
<style>
    @keyframes sweep {
        0% { transform: rotate(0deg); }
        100% { transform: rotate(360deg); }
    }
    .radar-container {
        display: flex;
        justify-content: center;
        align-items: center;
        margin-bottom: 5px;
    }
    .radar {
        position: relative;
        width: 60px;
        height: 60px;
        border-radius: 50%;
        border: 2px solid #E3000F;
        background: radial-gradient(circle, rgba(227,0,15,0.1) 0%, rgba(26,28,30,0) 70%);
        margin-right: 20px;
        overflow: hidden;
        cursor: pointer;
    }
    .radar::before {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        width: 30px;
        height: 2px;
        background-color: #E3000F;
        transform-origin: 0% 50%;
        animation: sweep 2s linear infinite;
    }
    .radar::after {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        transform: translate(-50%, -50%);
        width: 6px;
        height: 6px;
        background-color: #ffffff;
        border-radius: 50%;
    }
    .radar:hover {
        transform: scale(1.05);
        transition: transform 0.2s ease-in-out;
    }
    .title-text {
        font-size: 4.5rem;
        font-weight: 800;
        letter-spacing: 2px;
        color: #ffffff;
        margin: 0;
        text-transform: uppercase;
        cursor: pointer;
    }
    .subtitle-text {
        text-align: center; 
        color: #E3000F;
        font-weight: 600;
        letter-spacing: 1.5px;
        margin-top: 5px;
        font-size: 1.2rem;
    }
    .stButton>button {
        font-weight: bold;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    h3 {
        text-transform: uppercase;
        letter-spacing: 1px;
        border-bottom: 1px solid #444;
        padding-bottom: 10px;
    }
</style>
""", unsafe_allow_html=True)

# 2. Modern Header with Home Link workaround
if st.button("RESET TO SCANNER (HOME)", key="home_btn", help="Click to reload the page"):
    st.rerun()

st.markdown("""
<div class="radar-container" onclick="document.getElementById('home_btn').click();">
    <div class="radar"></div>
    <h1 class="title-text">RivalRadar</h1>
</div>
""", unsafe_allow_html=True)
st.markdown("<h4 class='subtitle-text'>[ AI-POWERED COMPETITOR ANALYSIS SYSTEM ]</h4>", unsafe_allow_html=True)
st.markdown("---")

# --- MACHINE DATA DEFINITIONS ---

MACHINE_TYPES = ["Tracked Excavator", "Wheeled Excavator", "Wheel Loader"]

CATEGORIES = {
    "Tracked Excavator": [
        "SY10U", "SY16C", "SY18U", "SY18C", "SY19E", "SY20C", "SY26U", "SY26C", 
        "SY35U", "SY35C", "SY50U", "SY60U", "SY60C", "SY75C", "SY80U", "SY95C", 
        "SY135C", "SY155U", "SY215C", "SY235C", "SY265C", "SY305C", "SY365H", 
        "SY390H", "SY500H", "SY750H", "SY980H"
    ],
    "Wheeled Excavator": [
        "SY155W", "SY175W", "7t Placeholder", "10t Placeholder"
    ],
    "Wheel Loader": [
        "SW305K", "SW405K"
    ]
}

PARAMS = {
    "Tracked Excavator": [
        "Engine Make/Model", "Emission Standard", "Gross Power (kW)", "Net Power (kW)", "Max Torque (Nm)", "Displacement (L)", "Number of Cylinders",
        "Operating Weight (kg)", "Transport Weight (kg)", "Counterweight (kg)",
        "Breakout Force - Bucket (kN)", "Tearout Force - Stick (kN)", "Drawbar Pull / Traction Force (kN)", "Max Travel Speed H/L (km/h)", "Swing Speed (rpm)", "Gradeability (%)",
        "Main Pump Type", "Max Flow (l/min)", "System Pressure (bar)", "AUX 1 Flow (l/min)", "AUX 2 Flow (l/min)",
        "Boom Length (mm)", "Standard Stick Length (mm)", "Max Digging Depth (mm)", "Max Digging Reach on Ground (mm)", "Max Dump Height (mm)", "Tail Swing Radius (mm)", "Ground Clearance (mm)", "Track Gauge (mm)", "Track Shoe Width (mm)",
        "Fuel Tank (l)", "Hydraulic System (l)", "Engine Oil (l)", "DEF/AdBlue Tank (l)"
    ],
    "Wheeled Excavator": [
        "Engine Make/Model", "Emission Standard", "Gross Power (kW)", "Net Power (kW)", "Max Torque (Nm)",
        "Operating Weight with Blade (kg)", "Operating Weight with Outriggers (kg)",
        "Breakout Force (kN)", "Tearout Force (kN)", "Swing Speed (rpm)",
        "Transmission Type", "Travel Speed Creep (km/h)", "Travel Speed Low (km/h)", "Travel Speed High (km/h)", "Tire Size/Type", "Steering Radius (mm)", "Front Axle Oscillation Angle (°)",
        "Main Pump Max Flow (l/min)", "System Pressure (bar)", "AUX 1 Flow (l/min)",
        "Max Digging Depth (mm)", "Max Reach (mm)", "Max Dump Height (mm)", "Tail Swing Radius (mm)", "Wheelbase (mm)", "Overall Width (mm)",
        "Fuel Tank (l)", "Hydraulic Tank (l)", "DEF/AdBlue Tank (l)"
    ],
    "Wheel Loader": [
        "Engine Make/Model", "Emission Standard", "Gross Power (kW)", "Max Torque (Nm)", "Displacement (L)",
        "Operating Weight (kg)",
        "Rated Payload (kg)", "Static Tipping Load - Straight (kg)", "Static Tipping Load - Full Turn (kg)", "Bucket Breakout Force (kN)", "Standard Bucket Capacity Heaped (m3)",
        "Raise Time (s)", "Dump Time (s)", "Lower Time (s)", "Total Cycle Time (s)",
        "Transmission Type", "Gears Forward/Reverse", "Max Travel Speed Forward (km/h)", "Articulation Angle (°)", "Turning Radius outside tires (mm)", "Tire Size",
        "Linkage Type (Z-Bar / Parallel)", "Hinge Pin Height at Max Lift (mm)", "Dump Clearance at Max Lift (mm)", "Reach at Max Lift (mm)", "Overall Length (mm)", "Wheelbase (mm)", "Overall Width (mm)",
        "Fuel Tank (l)", "Hydraulic Tank (l)", "DEF/AdBlue Tank (l)"
    ]
}

# --- DATABASE SETUP & GITHUB SYNC ---
DB_FILE = "machine_database.json"

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_db(data):
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=4)
        
    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
        try:
            from github import Github
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(st.secrets["GITHUB_REPO"])
            json_str = json.dumps(data, indent=4)
            try:
                contents = repo.get_contents(DB_FILE)
                repo.update_file(contents.path, "Auto-Sync: Database updated", json_str, contents.sha)
            except:
                repo.create_file(DB_FILE, "Auto-Sync: Database created", json_str)
        except Exception as e:
            print(f"GitHub Sync Error: {e}")

db = load_db()

# --- HELPER FOR CHARTS & PPT ---
def extract_number(val):
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        match = re.search(r"[-+]?\d*\.\d+|\d+", val.replace(',', '.'))
        if match:
            return float(match.group())
    return None

def get_template_path():
    if os.path.exists("sany_template.pptx"):
        return "sany_template.pptx"
    for root, dirs, files in os.walk("."):
        for file in files:
            if "sany" in file.lower() and file.endswith(".pptx") and not file.startswith("._") and not file.startswith("."):
                return os.path.join(root, file)
    return None

def create_pitch_deck(baseline_name, competitor_names, ai_text, df_battle):
    template_path = get_template_path()
    
    if not template_path:
        st.error("CRITICAL ERROR: SANY Template (.pptx) not found on server! Creating blank fallback deck.")
        prs = Presentation() 
    else:
        try:
            prs = Presentation(template_path)
        except Exception as e:
            st.error(f"Error loading template '{template_path}': {e}")
            prs = Presentation()
        
    # FOLIE 1: TITEL
    try:
        if len(prs.slides) > 0:
            slide1 = prs.slides[0]
            replaced_title = False
            for shape in slide1.shapes:
                if shape.has_text_frame and "XXX" in shape.text_frame.text:
                    shape.text_frame.text = shape.text_frame.text.replace("XXX", f"Tactical Product Comparison\n{baseline_name} vs. {', '.join(competitor_names)}")
                    replaced_title = True
                    break
            if not replaced_title:
                text_shapes_1 = [s for s in slide1.shapes if s.has_text_frame]
                text_shapes_1.sort(key=lambda s: s.width * s.height if s.width and s.height else 0, reverse=True)
                if len(text_shapes_1) > 0:
                    text_shapes_1[0].text_frame.text = f"Tactical Product Comparison\n{baseline_name} vs. {', '.join(competitor_names)}"
    except Exception as e:
        print(f"Slide 1 Error: {e}")
        
    # FOLIE 1.5: DATENTABELLE 
    try:
        table_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
        table_slide = prs.slides.add_slide(table_layout)
        
        slide_id_list = prs.slides._sldIdLst
        new_slide_id = slide_id_list[-1]
        slide_id_list.remove(new_slide_id)
        slide_id_list.insert(1, new_slide_id)
        
        title_set = False
        for shape in table_slide.shapes:
            if shape.has_text_frame and "XXXX" in shape.text_frame.text:
                shape.text_frame.text = shape.text_frame.text.replace("XXXX", "Technical Specifications Matrix")
                title_set = True
                break
        
        if not title_set and table_slide.shapes.title:
            table_slide.shapes.title.text = "Technical Specifications Matrix"
            
        for shape in table_slide.shapes:
            if shape.has_text_frame and "XXX" in shape.text_frame.text:
                shape.text_frame.text = ""

        rows = len(df_battle.columns)
        cols = len(df_battle) + 1
        
        left = Inches(0.5)
        top = Inches(1.8)
        width = Inches(9.0)
        height = Inches(0.4 * rows) 
        
        shape = table_slide.shapes.add_table(rows, cols, left, top, width, height)
        table = shape.table
        
        table.columns[0].width = Inches(2.5)
        for i in range(1, cols):
            table.columns[i].width = Inches(6.5 / (cols - 1))
            
        for row_idx, col_name in enumerate(df_battle.columns):
            cell = table.cell(row_idx, 0)
            cell.text = str(col_name)
            cell.text_frame.paragraphs[0].font.bold = True
            cell.text_frame.paragraphs[0].font.size = Pt(11)
            
        for m_idx, row_data in df_battle.iterrows():
            col_idx = m_idx + 1
            for r_idx, col_name in enumerate(df_battle.columns):
                cell = table.cell(r_idx, col_idx)
                val = str(row_data[col_name])
                if val == "nan" or val == "None": val = "-"
                cell.text = val
                cell.text_frame.paragraphs[0].font.size = Pt(11)
                cell.text_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
                if r_idx == 0: 
                    cell.text_frame.paragraphs[0].font.bold = True
    except Exception as e:
        print(f"Table Slide Error: {e}")

    # FOLIE 2: AI ANALYSIS 
    if ai_text:
        try:
            if len(prs.slides) > 2:
                slide3 = prs.slides[2]
            else:
                content_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
                slide3 = prs.slides.add_slide(content_layout)
                
            tf_title = None
            tf_content = None
            
            for shape in slide3.shapes:
                if shape.has_text_frame:
                    text = shape.text_frame.text
                    if "XXXX" in text:
                        shape.text_frame.text = text.replace("XXXX", "Competitive Analysis & Pitch")
                        tf_title = shape.text_frame
                    elif "XXX" in text:
                        tf_content = shape.text_frame
            
            if not tf_title or not tf_content:
                text_shapes_3 = [s for s in slide3.shapes if s.has_text_frame]
                text_shapes_3.sort(key=lambda s: s.width * s.height if s.width and s.height else 0, reverse=True)
                if len(text_shapes_3) > 0 and not tf_content:
                    tf_content = text_shapes_3[0].text_frame
                if len(text_shapes_3) > 1 and not tf_title:
                    tf_title = text_shapes_3[1].text_frame
                    
            if tf_title and "XXXX" in tf_title.text:
                tf_title.text = "Competitive Analysis & Pitch"
                
            if tf_content:
                tf_content.text = "" 
                tf_content.word_wrap = True
                try: tf_content.auto_size = 1 
                except: pass
                
                lines = ai_text.split('\n')
                first_paragraph = True
                
                for line in lines:
                    line = line.strip()
                    if not line or line.startswith('|') or line.startswith('---') or line.startswith('='):
                        continue
                    clean_line = line.replace('**', '').replace('### ', '').replace('## ', '')
                    if first_paragraph:
                        p = tf_content.paragraphs[0]
                        first_paragraph = False
                    else:
                        p = tf_content.add_paragraph()
                        
                    if clean_line.startswith('* ') or clean_line.startswith('- '):
                        p.text = clean_line[2:]
                        p.level = 1 
                        p.font.size = Pt(14)
                    else:
                        p.text = clean_line
                        p.level = 0
                        p.font.bold = True
                        p.font.size = Pt(16)
        except Exception as e:
            print(f"PPT Error: {e}")
            
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- DYNAMIC STATE VARIABLES ---
if "custom_params" not in st.session_state:
    st.session_state.custom_params = {}
    for mt in MACHINE_TYPES:
        st.session_state.custom_params[mt] = []

if "ai_analysis_cache" not in st.session_state:
    st.session_state.ai_analysis_cache = ""

# --- SIDEBAR ---
st.sidebar.markdown("### NAVIGATION")
app_mode = st.sidebar.radio("Navigate to:", ["Scanner", "Database", "Product Comparison", "News Radar"])
st.sidebar.markdown("---")

if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.error("API-Key loaded successfully.") 
else:
    st.sidebar.markdown("### SYSTEM SETTINGS")
    api_key = st.sidebar.text_input("Gemini API Key (Manual)", type="password")

selected_model_name = "gemini-3.1-flash-lite"
if api_key:
    genai.configure(api_key=api_key)
    st.sidebar.markdown("### AI MODEL")
    try:
        available_models = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        if available_models:
            default_idx = 0
            if "gemini-3.1-flash-lite" in available_models:
                default_idx = available_models.index("gemini-3.1-flash-lite")
            elif any("flash-lite" in m for m in available_models):
                default_idx = next(i for i, m in enumerate(available_models) if "flash-lite" in m)
            selected_model_name = st.sidebar.selectbox("Active Model:", available_models, index=default_idx)
    except Exception:
        selected_model_name = st.sidebar.text_input("Manual Model Input:", value="gemini-3.1-flash-lite")

st.sidebar.markdown("---")

# ================= VIEW 1: SCANNER =================
if app_mode == "Scanner":
    st.markdown("### UPLOAD DATASHEETS")
    
    st.markdown("#### SELECT MACHINE TYPE")
    selected_machine_type = st.radio("Configure AI Brain For:", MACHINE_TYPES, horizontal=True)
    
    current_categories = CATEGORIES[selected_machine_type]
    current_base_params = PARAMS[selected_machine_type]
    all_available_params = current_base_params + st.session_state.custom_params[selected_machine_type]

    with st.expander(f"CONFIGURE PARAMS ({selected_machine_type})", expanded=False):
        st.error(f"The system recognizes {len(all_available_params)} parameters for {selected_machine_type}. The scanner automatically extracts all known metrics.")
        new_param_input = st.text_input("Add custom parameter:", placeholder="e.g., Track width (mm)", key=f"new_param_{selected_machine_type}")
        if st.button("Save Parameter", use_container_width=True):
            if new_param_input:
                clean_param = new_param_input.strip()
                if clean_param not in all_available_params:
                    st.session_state.custom_params[selected_machine_type].append(clean_param)
                    st.success(f"'{clean_param}' added to {selected_machine_type}.")
                    st.rerun()
    
    uploaded_files = st.file_uploader("Drop brochures or datasheets here (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

    machine_configs = {}
    if uploaded_files:
        if not api_key:
            st.error("ACCESS DENIED: Please provide a valid API Key.")
        else:
            for file in uploaded_files:
                sub_col1, sub_col2 = st.columns(2)
                with sub_col1:
                    default_name = file.name.rsplit('.', 1)[0]
                    m_name = st.text_input(f"Machine Name:", value=default_name, key=f"name_{file.name}")
                with sub_col2:
                    m_cat = st.selectbox(f"Baseline Class ({selected_machine_type}):", options=current_categories, key=f"cat_{file.name}")
                machine_configs[file.name] = {"name": m_name, "category": m_cat}

    if st.button("INITIATE AI SCAN (EXTRACT ALL DATA)", type="primary", use_container_width=True):
        if uploaded_files:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for index, file in enumerate(uploaded_files):
                current_machine = machine_configs[file.name]["name"]
                current_category = machine_configs[file.name]["category"]
                status_text.text(f"Scanning '{current_machine}' as {selected_machine_type}...")
                
                file.seek(0)
                file_part = {"mime_type": file.type, "data": file.read()}
                prompt = f"""You are a precise technical data extraction assistant analyzing a {selected_machine_type}. Focus EXCLUSIVELY on: "{current_machine}". Extract exact values for: {json.dumps(all_available_params)}. 1. Valid JSON object only. 2. Exact keys. 3. Use "?" if missing. 4. Convert imperial to metric. ONLY JSON."""
                
                try:
                    model = genai.GenerativeModel(selected_model_name)
                    response = model.generate_content(contents=[file_part, prompt], generation_config={"response_mime_type": "application/json"})
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("
http://googleusercontent.com/immersive_entry_chip/0
http://googleusercontent.com/immersive_entry_chip/1
http://googleusercontent.com/immersive_entry_chip/2
