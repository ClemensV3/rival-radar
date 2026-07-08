import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import os
import re
import time
import altair as alt
from pptx import Presentation
from pptx.util import Inches, Pt

# 1. Page Configuration
st.set_page_config(page_title="RivalRadar", layout="wide", page_icon="📡")

# 2. Modern Header
st.markdown("<h1 style='text-align: center;'>📡 RivalRadar</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: #10b981; font-weight: normal;'>[ AI-POWERED COMPETITOR ANALYSIS FOR SALES ]</h4>", unsafe_allow_html=True)
st.markdown("---")

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
    # Lokal speichern
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=4)
        
    # Auf GitHub spiegeln
    if "GITHUB_TOKEN" in st.secrets and "GITHUB_REPO" in st.secrets:
        try:
            from github import Github
            g = Github(st.secrets["GITHUB_TOKEN"])
            repo = g.get_repo(st.secrets["GITHUB_REPO"])
            json_str = json.dumps(data, indent=4)
            try:
                contents = repo.get_contents(DB_FILE)
                repo.update_file(contents.path, "🤖 Auto-Sync: Database updated", json_str, contents.sha)
            except:
                repo.create_file(DB_FILE, "🤖 Auto-Sync: Database created", json_str)
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

def create_pitch_deck(baseline_name, competitor_names, ai_text):
    # Versuche das SANY Template zu laden, ansonsten nimm ein leeres Template
    try:
        prs = Presentation("sany_template.pptx")
    except Exception:
        prs = Presentation()
        
    # Slide 1: Title Slide
    try:
        title_layout = prs.slide_layouts[0]
        slide1 = prs.slides.add_slide(title_layout)
        title = slide1.shapes.title
        subtitle = slide1.placeholders[1]
        
        title.text = "Tactical Product Comparison"
        subtitle.text = f"{baseline_name} vs. {', '.join(competitor_names)}"
    except Exception as e:
        pass # Robust fallback falls das Template anders aufgebaut ist
        
    # Slide 2: AI Analysis
    if ai_text:
        try:
            content_layout = prs.slide_layouts[1]
            slide2 = prs.slides.add_slide(content_layout)
            
            title2 = slide2.shapes.title
            if title2:
                title2.text = "AI Competitive Analysis & Sales Pitch"
                
            # Erstelle eine Textbox für den generierten KI Text
            left = Inches(0.5)
            top = Inches(1.5)
            width = Inches(9)
            height = Inches(5.5)
            txBox = slide2.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.word_wrap = True
            
            # Bereinige Markdown Formatierungen (**, #) für sauberen PowerPoint Text
            clean_text = ai_text.replace('**', '').replace('### ', '').replace('## ', '')
            
            p = tf.paragraphs[0]
            p.text = clean_text
            p.font.size = Pt(14)
        except Exception as e:
            pass
            
    # In den Speicher laden (damit wir es in Streamlit herunterladen können)
    ppt_stream = io.BytesIO()
    prs.save(ppt_stream)
    ppt_stream.seek(0)
    return ppt_stream

# --- DYNAMIC STATE VARIABLES ---
if "custom_params" not in st.session_state:
    st.session_state.custom_params = []

if "ai_analysis_cache" not in st.session_state:
    st.session_state.ai_analysis_cache = ""

# --- BASE DATA ---
CATEGORIES = [
    "SY10U", "SY16C", "SY18U", "SY18C", "SY19E", "SY20C", "SY26U", "SY26C", 
    "SY35U", "SY35C", "SY50U", "SY60U", "SY60C", "SY75C", "SY80U", "SY95C", 
    "SY135C", "SY155U", "SY215C", "SY235C", "SY265C", "SY305C", "SY365H", 
    "SY390H", "SY500H", "SY750H", "SY980H"
]

BASE_PARAMS = [
    "Operating weight (kg)", "Engine Power STD (kW)", "Engine Power OPT (kW)",
    "Engine STD", "Engine OPT", "Cylinder", "Displacement", "STD Bucket Capacity",
    "Breakout force (kN)", "Swing force (kNm)", "U/min (Swing)", "Traction Force (kN)",
    "Fuel Tank (l)", "Hydraulik Oil tank (l)", "Hydraulik Oil System (l)", "AdBlue Tank (l)",
    "Max Cab height", "Max width", "Swing radius", "Ground clearance", "Transport lengths",
    "Max Digging depth", "Vertical reach", "Standard Stick", "Long Stick", "AUX 1 Flow",
    "AUX 2 Flow", "Quick coupler line", "Inside Cab (ISO 6396)", "Outside Cab (2000/14/EG)",
    "STD Speed (kmh)", "OPT Speed (kmh)"
]

# --- SIDEBAR ---
st.sidebar.markdown("### 🧭 NAVIGATION")
app_mode = st.sidebar.radio("Go to:", ["📡 Scanner", "📚 Database", "📊 Product Comparison"])
st.sidebar.markdown("---")

if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("🔑 API-Key loaded!")
else:
    st.sidebar.markdown("### ⚙️ SYSTEM SETTINGS")
    api_key = st.sidebar.text_input("🔑 Gemini API Key (Manual)", type="password")

selected_model_name = "gemini-1.5-flash"
if api_key:
    genai.configure(api_key=api_key)
    st.sidebar.markdown("### 🧠 AI MODEL")
    try:
        available_models = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        if available_models:
            default_idx = 0
            for i, m in enumerate(available_models):
                if "flash" in m:
                    default_idx = i
                    break
            selected_model_name = st.sidebar.selectbox("Active Model:", available_models, index=default_idx)
    except Exception:
        selected_model_name = st.sidebar.text_input("Manual Model Input:", value="gemini-1.5-flash")

st.sidebar.markdown("---")

with st.sidebar.expander("⚙️ CONFIGURE PARAMETERS", expanded=False):
    st.markdown("#### ➕ Add new metric")
    new_param_input = st.text_input("Parameter name:", placeholder="e.g., Track width (mm)", key="new_param_field")
    if st.button("Save parameter", use_container_width=True):
        if new_param_input:
            clean_param = new_param_input.strip()
            if clean_param not in BASE_PARAMS and clean_param not in st.session_state.custom_params:
                st.session_state.custom_params.append(clean_param)
                st.success(f"'{clean_param}' added!")
                st.rerun()
    st.markdown("---")
    st.markdown("#### 🎯 Active Scan Metrics")
    all_available_params = BASE_PARAMS + st.session_state.custom_params
    selected_parameters = st.multiselect("Selected Parameters:", options=all_available_params, default=all_available_params)

# ================= VIEW 1: SCANNER =================
if app_mode == "📡 Scanner":
    st.markdown("### 📥 UPLOAD DATASHEETS")
    uploaded_files = st.file_uploader("Drop brochures or datasheets here (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

    machine_configs = {}
    if uploaded_files:
        if not api_key:
            st.error("ACCESS DENIED: Please provide an API Key.")
        else:
            for file in uploaded_files:
                sub_col1, sub_col2 = st.columns(2)
                with sub_col1:
                    default_name = file.name.rsplit('.', 1)[0]
                    m_name = st.text_input(f"Machine Name:", value=default_name, key=f"name_{file.name}")
                with sub_col2:
                    m_cat = st.selectbox(f"Baseline Class (Sany):", options=CATEGORIES, key=f"cat_{file.name}")
                machine_configs[file.name] = {"name": m_name, "category": m_cat}

    if st.button("🚀 INITIATE AI SCAN (EXTRACT DATA)", type="primary", use_container_width=True):
        if uploaded_files and selected_parameters:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for index, file in enumerate(uploaded_files):
                current_machine = machine_configs[file.name]["name"]
                current_category = machine_configs[file.name]["category"]
                status_text.text(f"Scanning '{current_machine}'...")
                
                file.seek(0)
                file_part = {"mime_type": file.type, "data": file.read()}
                prompt = f"""You are a precise technical data extraction assistant. Focus EXCLUSIVELY on: "{current_machine}". Extract exact values for: {json.dumps(selected_parameters)}. 1. Valid JSON object only. 2. Exact keys. 3. Use "?" if missing. 4. Convert imperial to metric. ONLY JSON."""
                
                try:
                    model = genai.GenerativeModel(selected_model_name)
                    response = model.generate_content(contents=[file_part, prompt], generation_config={"response_mime_type": "application/json"})
                    
                    raw_text = response.text.strip()
                    if raw_text.startswith("
