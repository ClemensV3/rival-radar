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
from pptx.enum.text import PP_ALIGN

# 1. Page Configuration
st.set_page_config(page_title="RivalRadar", layout="wide")

# --- CUSTOM CSS FOR SANY LOOK & RADAR ---
st.markdown("""
<style>
    /* Custom Radar Animation */
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
        width: 40px;
        height: 40px;
        border-radius: 50%;
        border: 2px solid #E3000F; /* SANY Red */
        background: radial-gradient(circle, rgba(227,0,15,0.1) 0%, rgba(26,28,30,0) 70%);
        margin-right: 15px;
        overflow: hidden;
    }
    .radar::before {
        content: '';
        position: absolute;
        top: 50%;
        left: 50%;
        width: 20px;
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
        width: 4px;
        height: 4px;
        background-color: #ffffff;
        border-radius: 50%;
    }
    
    /* Typography Overrides */
    .title-text {
        font-size: 3rem;
        font-weight: 800;
        letter-spacing: 2px;
        color: #ffffff;
        margin: 0;
        text-transform: uppercase;
    }
    .subtitle-text {
        text-align: center; 
        color: #E3000F; /* SANY Red */
        font-weight: 600;
        letter-spacing: 1.5px;
        margin-top: 5px;
        font-size: 1.1rem;
    }
    
    /* Global Adjustments */
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

# 2. Modern Header
st.markdown("""
<div class="radar-container">
    <div class="radar"></div>
    <h1 class="title-text">RivalRadar</h1>
</div>
""", unsafe_allow_html=True)
st.markdown("<h4 class='subtitle-text'>[ AI-POWERED COMPETITOR ANALYSIS SYSTEM ]</h4>", unsafe_allow_html=True)
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
            st.success(f"Template successfully loaded from: {template_path}")
        except Exception as e:
            st.error(f"Error loading template '{template_path}': {e}")
            prs = Presentation()
        
    # --- FOLIE 1: TITEL (Platzhalter "XXX") ---
    try:
        if len(prs.slides) > 0:
            slide1 = prs.slides[0]
            replaced_title = False
            
            for shape in slide1.shapes:
                if shape.has_text_frame:
                    if "XXX" in shape.text_frame.text:
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
        
    # --- FOLIE 1.5: DATENTABELLE (NEUE FOLIE) ---
    try:
        table_layout = prs.slide_layouts[1] if len(prs.slide_layouts) > 1 else prs.slide_layouts[0]
        table_slide = prs.slides.add_slide(table_layout)
        
        # XML-Hack: Verschiebe diese Folie auf Position 2 (Index 1)
        slide_id_list = prs.slides._sldIdLst
        new_slide_id = slide_id_list[-1]
        slide_id_list.remove(new_slide_id)
        slide_id_list.insert(1, new_slide_id)
        
        # Platziere den Titel (versuche XXXX Platzhalter zu nutzen)
        title_set = False
        for shape in table_slide.shapes:
            if shape.has_text_frame and "XXXX" in shape.text_frame.text:
                shape.text_frame.text = shape.text_frame.text.replace("XXXX", "Technical Specifications Matrix")
                title_set = True
                break
        
        if not title_set and table_slide.shapes.title:
            table_slide.shapes.title.text = "Technical Specifications Matrix"
            
        # Optional: XXX Platzhalter löschen, da wir eine Tabelle zeichnen
        for shape in table_slide.shapes:
            if shape.has_text_frame and "XXX" in shape.text_frame.text:
                shape.text_frame.text = ""

        # Tabelle Zeichnen
        rows = len(df_battle.columns)
        cols = len(df_battle) + 1
        
        left = Inches(0.5)
        top = Inches(1.8)
        width = Inches(9.0)
        height = Inches(0.4 * rows) # Dynamische Höhe
        
        shape = table_slide.shapes.add_table(rows, cols, left, top, width, height)
        table = shape.table
        
        # Spaltenbreiten justieren (Parameter-Name breiter)
        table.columns[0].width = Inches(2.5)
        for i in range(1, cols):
            table.columns[i].width = Inches(6.5 / (cols - 1))
            
        # Zelle befüllen
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

    # --- FOLIE 2 (Jetzt Folie 3): AI ANALYSIS ---
    if ai_text:
        try:
            # Da wir eine Folie eingefügt haben, ist die AI-Folie jetzt auf Index 2
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
                
                try:
                    tf_content.auto_size = 1 
                except:
                    pass
                
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
st.sidebar.markdown("### NAVIGATION")
app_mode = st.sidebar.radio("Navigate to:", ["Scanner", "Database", "Product Comparison"])
st.sidebar.markdown("---")

if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("API-Key loaded successfully.")
else:
    st.sidebar.markdown("### SYSTEM SETTINGS")
    api_key = st.sidebar.text_input("Gemini API Key (Manual)", type="password")

# Model Selection Logic for gemini-3.1-flash-lite
selected_model_name = "gemini-3.1-flash-lite"
if api_key:
    genai.configure(api_key=api_key)
    st.sidebar.markdown("### AI MODEL")
    try:
        available_models = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        if available_models:
            # Force gemini-3.1-flash-lite as default
            default_idx = 0
            if "gemini-3.1-flash-lite" in available_models:
                default_idx = available_models.index("gemini-3.1-flash-lite")
            elif any("flash-lite" in m for m in available_models):
                default_idx = next(i for i, m in enumerate(available_models) if "flash-lite" in m)
                
            selected_model_name = st.sidebar.selectbox("Active Model:", available_models, index=default_idx)
    except Exception:
        selected_model_name = st.sidebar.text_input("Manual Model Input:", value="gemini-3.1-flash-lite")

st.sidebar.markdown("---")

with st.sidebar.expander("CONFIGURE DATABASE PARAMS", expanded=False):
    st.info("The system recognizes these parameters. The scanner automatically extracts all known metrics in the background.")
    new_param_input = st.text_input("Parameter Name:", placeholder="e.g., Track width (mm)", key="new_param_field")
    if st.button("Save Parameter", use_container_width=True):
        if new_param_input:
            clean_param = new_param_input.strip()
            if clean_param not in BASE_PARAMS and clean_param not in st.session_state.custom_params:
                st.session_state.custom_params.append(clean_param)
                st.success(f"'{clean_param}' added successfully.")
                st.rerun()
                
all_available_params = BASE_PARAMS + st.session_state.custom_params

# ================= VIEW 1: SCANNER =================
if app_mode == "Scanner":
    st.markdown("### UPLOAD DATASHEETS")
    st.info("The AI scanner automatically extracts all known metrics from the provided documents. You can select specific parameters to compare in the 'Product Comparison' tab.")
    
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
                    m_cat = st.selectbox(f"Baseline Class (SANY):", options=CATEGORIES, key=f"cat_{file.name}")
                machine_configs[file.name] = {"name": m_name, "category": m_cat}

    if st.button("INITIATE AI SCAN (EXTRACT ALL DATA)", type="primary", use_container_width=True):
        if uploaded_files:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for index, file in enumerate(uploaded_files):
                current_machine = machine_configs[file.name]["name"]
                current_category = machine_configs[file.name]["category"]
                status_text.text(f"Scanning '{current_machine}'...")
                
                file.seek(0)
                file_part = {"mime_type": file.type, "data": file.read()}
                prompt = f"""You are a precise technical data extraction assistant. Focus EXCLUSIVELY on: "{current_machine}". Extract exact values for: {json.dumps(all_available_params)}. 1. Valid JSON object only. 2. Exact keys. 3. Use "?" if missing. 4. Convert imperial to metric. ONLY JSON."""
                
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
                    st.error(f"Scan failed for '{current_machine}'. Error: {e}")
                
                progress_bar.progress((index + 1) / len(uploaded_files))
            
            status_text.text("Data extraction complete. Machines synced to database.")
            time.sleep(2)
            st.rerun()

# ================= VIEW 2: DATABASE =================
elif app_mode == "Database":
    st.markdown("### MACHINE DATABASE")
    if db:
        col_filter, col_delete = st.columns(2)
        with col_filter:
            lib_filter = st.selectbox("Filter database by SANY Class:", ["All"] + CATEGORIES, key="lib_filter_select")
        with col_delete:
            delete_options = ["None"] + list(db.keys())
            to_delete = st.selectbox("Remove faulty record:", options=delete_options)
            if st.button("DELETE RECORD", use_container_width=True):
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
        st.download_button("DOWNLOAD DATABASE (.xlsx)", excel_buffer_lib.getvalue(), "Database.xlsx", use_container_width=True)
    else:
        st.info("The database is currently empty.")

# ================= VIEW 3: PRODUCT COMPARISON =================
elif app_mode == "Product Comparison":
    st.markdown("### PRODUCT COMPARISON (THE ARENA)")
    if not db:
        st.info("No data available yet.")
    else:
        db_keys = list(db.keys())
        arena_col1, arena_col2 = st.columns(2)
        with arena_col1:
            baseline_sel = st.selectbox("Own Product (SANY Baseline):", options=db_keys)
        with arena_col2:
            competitors_sel = st.multiselect("Competitor Models:", options=[k for k in db_keys if k != baseline_sel])
            
        st.markdown("---")
        st.markdown("### MATCH SETUP")
        
        default_params = ["Operating weight (kg)", "Engine Power STD (kW)", "Max Digging depth", "Breakout force (kN)", "AUX 1 Flow"]
        compare_params = st.multiselect("Select parameters for Matrix & AI Analysis:", options=all_available_params, default=[p for p in default_params if p in all_available_params])
        
        st.markdown("### AI BIOS SWITCH")
        ai_persona = st.radio("Select AI Persona:", 
                              ["Sales Mode (Punchy, Strategic, ROI & Sales Focus)", 
                               "R&D Mode (Technical, Analytical, Engineering & Structure)"])
        
        st.markdown("### OUTPUT GENERATION")
        pwr_charts = st.checkbox("Visual Performance Comparison (Charts in App)")
        pwr_ampel = st.checkbox("Strengths/Weaknesses Profile (For Pitch Deck)")
        pwr_pitch = st.checkbox("Sales/Tech Arguments (For Pitch Deck)")
        
        if st.button("GENERATE COMPARISON", type="primary", use_container_width=True):
            if competitors_sel and compare_params:
                battle_data = [db[baseline_sel]] + [db[k] for k in competitors_sel]
                df_battle_full = pd.DataFrame(battle_data).drop(columns=['Category'], errors='ignore')
                
                cols_to_keep = ['Machine'] + [p for p in compare_params if p in df_battle_full.columns]
                df_battle_filtered = df_battle_full[cols_to_keep]
                
                st.session_state.current_df_battle = df_battle_filtered
                
                st.markdown("---")
                st.write("### Raw Data Matrix (Filtered)")
                st.dataframe(df_battle_filtered.set_index("Machine").T, use_container_width=True)
                
                if pwr_charts:
                    st.markdown("---")
                    st.write("### Visual Performance Comparison")
                    chart_cols = st.columns(2)
                    col_idx = 0
                    for metric in compare_params:
                        if metric in df_battle_filtered.columns:
                            chart_data = [{"Machine": row['Machine'], metric: extract_number(row.get(metric))} for _, row in df_battle_filtered.iterrows() if extract_number(row.get(metric)) is not None]
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
                    st.write("### AI Competitive Analysis")
                    with st.spinner("The AI is analyzing the data..."):
                        baseline_name = db[baseline_sel]['Machine']
                        competitor_names = [db[k]['Machine'] for k in competitors_sel]
                        
                        if "Sales" in ai_persona:
                            sys_prompt = f"You are a Senior Sales Strategist. English only. Baseline: '{baseline_name}'. Competitors: {', '.join(competitor_names)}.\n\n"
                            sys_prompt += f"Data: {df_battle_filtered.to_dict(orient='records')}\n\n"
                            sys_prompt += "CRITICAL: NEVER USE MARKDOWN TABLES! NO '|' SYMBOLS. DO NOT DRAW TABLES.\n"
                            sys_prompt += "Write the competitive analysis strictly as plain text paragraphs and short bullet points starting with '*'. Keep the text concise to fit on a presentation slide.\n"
                            if pwr_ampel: sys_prompt += "- Objective assessment (Use only text, no emojis if possible, outline Superior, Tie, Competitor superior) formatted as SHORT text bullets.\n"
                            if pwr_pitch: sys_prompt += "- Short, punchy sales arguments (Max 1 sentence each) focusing on productivity, ROI, and why the customer should buy SANY.\n"
                        else:
                            sys_prompt = f"You are a Senior R&D Engineer. English only. Baseline: '{baseline_name}'. Competitors: {', '.join(competitor_names)}.\n\n"
                            sys_prompt += f"Data: {df_battle_filtered.to_dict(orient='records')}\n\n"
                            sys_prompt += "CRITICAL: NEVER USE MARKDOWN TABLES! NO '|' SYMBOLS. DO NOT DRAW TABLES.\n"
                            sys_prompt += "Write the technical analysis strictly as plain text paragraphs and short bullet points starting with '*'. Keep the text concise to fit on a presentation slide.\n"
                            if pwr_ampel: sys_prompt += "- Objective technical assessment (Use only text, no emojis if possible, outline Superior, Tie, Competitor superior) formatted as SHORT text bullets.\n"
                            if pwr_pitch: sys_prompt += "- Deep dive into engineering tradeoffs, mechanical advantages, hydraulic efficiency, and structural integrity. Use highly technical terminology (Max 1 sentence each).\n"
                            
                        try:
                            model = genai.GenerativeModel(selected_model_name)
                            response = model.generate_content(sys_prompt)
                            st.session_state.ai_analysis_cache = response.text
                            st.markdown(response.text)
                        except Exception as e:
                            st.error(f"AI Analysis failed: {e}")

        # --- PPT EXPORT BUTTON ---
        if st.session_state.get("ai_analysis_cache") and competitors_sel and "current_df_battle" in st.session_state:
            st.markdown("---")
            st.markdown("### EXPORT PITCH DECK")
            
            baseline_name = db[baseline_sel]['Machine']
            competitor_names = [db[k]['Machine'] for k in competitors_sel]
            
            ppt_file = create_pitch_deck(baseline_name, competitor_names, st.session_state.ai_analysis_cache, st.session_state.current_df_battle)
            
            st.download_button(
                label="DOWNLOAD PITCH DECK (.pptx)",
                data=ppt_file.getvalue(),
                file_name=f"SANY_Pitch_{baseline_name}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
                type="primary"
            )
