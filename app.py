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
    try:
        prs = Presentation("sany_template.pptx")
    except Exception:
        prs = Presentation() # Fallback
        
    # Slide 1: Title Slide
    try:
        title_layout = prs.slide_layouts[0]
        slide1 = prs.slides.add_slide(title_layout)
        if slide1.shapes.title:
            slide1.shapes.title.text = "Tactical Product Comparison"
        if len(slide1.placeholders) > 1:
            slide1.placeholders[1].text = f"{baseline_name} vs. {', '.join(competitor_names)}"
    except Exception:
        pass 
        
    # Slide 2: AI Analysis (CORPORATE BRANDING & AUTO-FIT HACK)
    if ai_text:
        try:
            # Wir zwingen ihn, das zweite Layout (meist "Titel & Inhalt") zu nehmen
            content_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
            slide2 = prs.slides.add_slide(content_layout)
            
            if slide2.shapes.title:
                slide2.shapes.title.text = "Competitive Analysis & Pitch"
                
            # Wir suchen EXAKT den Platzhalter für den Inhalt (idx 1 ist meistens der Body)
            tf = None
            for shape in slide2.placeholders:
                if shape != slide2.shapes.title and shape.has_text_frame:
                    tf = shape.text_frame
                    break
            
            # Falls das Template verhunzt ist, machen wir eine Not-Box
            if tf is None:
                txBox = slide2.shapes.add_textbox(Inches(0.5), Inches(1.5), Inches(9), Inches(5.5))
                tf = txBox.text_frame
                
            tf.word_wrap = True
            
            # --- AUTO FIT MAGIC ---
            # Wir sagen PowerPoint: "Pass die Schriftgröße an, wenn es nicht passt!"
            tf.auto_size = 1 # 1 = MSO_AUTO_SIZE_TEXT_TO_FIT_SHAPE
            
            lines = ai_text.split('\n')
            first_paragraph = True
            
            for line in lines:
                line = line.strip()
                # Schmutz rausfiltern
                if not line or line.startswith('|') or line.startswith('---') or line.startswith('='):
                    continue
                    
                clean_line = line.replace('**', '').replace('### ', '').replace('## ', '')
                
                if first_paragraph:
                    p = tf.paragraphs[0]
                    first_paragraph = False
                else:
                    p = tf.add_paragraph()
                    
                # Strukturieren!
                if clean_line.startswith('* ') or clean_line.startswith('- '):
                    p.text = clean_line[2:]
                    p.level = 1 
                    # Base Font Size setzen, PowerPoint verkleinert dann automatisch
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
                    if raw_text.startswith("```json"): 
                        raw_text = raw_text[7:]
                    elif raw_text.startswith("```"): 
                        raw_text = raw_text[3:]
                    if raw_text.endswith("```"): 
                        raw_text = raw_text[:-3]
                        
                    extracted_json = json.loads(raw_text.strip())
                    extracted_json["Machine"] = current_machine
                    extracted_json["Category"] = current_category
                    
                    db[f"{current_machine} ({current_category})"] = extracted_json
                    save_db(db)
                except Exception as e:
                    st.error(f"❌ Scan failed for '{current_machine}'. Error: {e}")
                
                progress_bar.progress((index + 1) / len(uploaded_files))
            
            status_text.text("✅ Data extraction complete. Machines synced to GitHub.")
            time.sleep(2)
            st.rerun()

# ================= VIEW 2: DATABASE =================
elif app_mode == "📚 Database":
    st.markdown("### 📚 MACHINE DATABASE")
    if db:
        col_filter, col_delete = st.columns(2)
        with col_filter:
            lib_filter = st.selectbox("Filter database by Sany Class:", ["All"] + CATEGORIES, key="lib_filter_select")
        with col_delete:
            delete_options = ["None"] + list(db.keys())
            to_delete = st.selectbox("🗑️ Remove faulty record:", options=delete_options)
            if st.button("🧨 DELETE RECORD", use_container_width=True):
                if to_delete != "None":
                    del db[to_delete]
                    save_db(db)
                    st.rerun()
                    
        st.markdown("---")
        df_lib = pd.DataFrame(list(db.values())).rename(columns={'Category': 'Class'})
        cols = ['Machine', 'Class'] + [c for c in df_lib.columns if c not in ['Machine', 'Class']]
        df_lib = df_lib[cols]
        if lib_filter != "All":
            df_lib = df_lib[df_lib['Class'] == lib_filter]
            
        st.dataframe(df_lib, use_container_width=True)
        
        excel_buffer_lib = io.BytesIO()
        with pd.ExcelWriter(excel_buffer_lib, engine='openpyxl') as writer:
            df_lib.to_excel(writer, index=False)
        st.download_button("📥 DOWNLOAD DATABASE (.xlsx)", excel_buffer_lib.getvalue(), "Database.xlsx", use_container_width=True)
    else:
        st.info("The database is currently empty.")

# ================= VIEW 3: PRODUCT COMPARISON =================
elif app_mode == "📊 Product Comparison":
    st.markdown("### 📊 PRODUCT COMPARISON (COMPETITIVE ANALYSIS)")
    if not db:
        st.info("No data available yet.")
    else:
        db_keys = list(db.keys())
        arena_col1, arena_col2 = st.columns(2)
        with arena_col1:
            baseline_sel = st.selectbox("🟢 Own Product (Sany Baseline):", options=db_keys)
        with arena_col2:
            competitors_sel = st.multiselect("🔴 Competitor Models:", options=[k for k in db_keys if k != baseline_sel])
            
        st.markdown("#### ✨ ACTIVATE AI ANALYSES")
        pwr_charts = st.checkbox("📊 Visual Performance Comparison (Charts)")
        pwr_ampel = st.checkbox("🚦 Strengths/Weaknesses Profile")
        pwr_pitch = st.checkbox("💬 Sales Arguments (Pitch)")
        
        if st.button("⚖️ GENERATE COMPARISON", type="primary", use_container_width=True):
            if competitors_sel:
                battle_data = [db[baseline_sel]] + [db[k] for k in competitors_sel]
                df_battle = pd.DataFrame(battle_data).drop(columns=['Category'], errors='ignore')
                
                st.markdown("---")
                st.write("### 🗄️ Raw Data Overview")
                st.dataframe(df_battle.set_index("Machine").T, use_container_width=True)
                
                if pwr_charts:
                    st.markdown("---")
                    st.write("### 📊 Visual Performance Comparison")
                    chart_metrics = ["Operating weight (kg)", "Engine Power STD (kW)", "Max Digging depth", "Breakout force (kN)"]
                    chart_cols = st.columns(2)
                    col_idx = 0
                    for metric in chart_metrics:
                        if metric in df_battle.columns:
                            chart_data = [{"Machine": row['Machine'], metric: extract_number(row.get(metric))} for _, row in df_battle.iterrows() if extract_number(row.get(metric)) is not None]
                            if chart_data:
                                chart = alt.Chart(pd.DataFrame(chart_data)).mark_bar().encode(
                                    x=alt.X('Machine', title=None, axis=alt.Axis(labelAngle=-45)),
                                    y=alt.Y(metric, title=None),
                                    color=alt.Color('Machine', legend=None),
                                    tooltip=['Machine', metric]
                                ).properties(height=300, title=metric)
                                with chart_cols[col_idx % 2]:
                                    st.altair_chart(chart, use_container_width=True)
                                col_idx += 1

                if pwr_ampel or pwr_pitch:
                    st.markdown("---")
                    st.write("### 🧠 AI Competitive Analysis")
                    with st.spinner("The AI is analyzing the data..."):
                        baseline_name = db[baseline_sel]['Machine']
                        competitor_names = [db[k]['Machine'] for k in competitors_sel]
                        
                        # --- HIER IST DER VERBESSERTE PROMPT FÜR DIE KI ---
                        sys_prompt = f"You are a Senior Sales Strategist. English only. Baseline: '{baseline_name}'. Competitors: {', '.join(competitor_names)}. Data: {json.dumps(battle_data)}.\n\n"
                        sys_prompt += "CRITICAL: NEVER USE MARKDOWN TABLES! NO '|' SYMBOLS. DO NOT DRAW TABLES.\n"
                        sys_prompt += "Write the competitive analysis strictly as plain text paragraphs and short bullet points starting with '*'. Keep the text concise to fit on a presentation slide.\n"
                        
                        if pwr_ampel: sys_prompt += "- Objective assessment (🟢 Superior, 🟡 Tie, 🔴 Competitor superior) formatted as SHORT text bullets.\n"
                        if pwr_pitch: sys_prompt += "- Short, punchy sales arguments (Bullet points max 1 sentence each).\n"
                        
                        try:
                            model = genai.GenerativeModel(selected_model_name)
                            response = model.generate_content(sys_prompt)
                            st.session_state.ai_analysis_cache = response.text
                            st.markdown(response.text)
                        except Exception as e:
                            st.error(f"AI Analysis failed: {e}")

        # --- PPT EXPORT BUTTON ---
        if st.session_state.get("ai_analysis_cache") and competitors_sel:
            st.markdown("---")
            st.markdown("### 📥 EXPORT PITCH DECK")
            baseline_name = db[baseline_sel]['Machine']
            competitor_names = [db[k]['Machine'] for k in competitors_sel]
            
            ppt_file = create_pitch_deck(baseline_name, competitor_names, st.session_state.ai_analysis_cache)
            
            st.download_button(
                label="🚀 DOWNLOAD PITCH DECK (.pptx)",
                data=ppt_file.getvalue(),
                file_name=f"SANY_Pitch_{baseline_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary"
            )
