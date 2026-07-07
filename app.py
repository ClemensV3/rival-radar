import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import os
import re
import time
import altair as alt

# 1. Page Configuration
st.set_page_config(page_title="RivalRadar", layout="wide", page_icon="📡")

# 2. Modern Header
st.markdown("<h1 style='text-align: center;'>📡 RivalRadar</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: #10b981; font-weight: normal;'>[ AI-POWERED COMPETITOR ANALYSIS FOR SALES ]</h4>", unsafe_allow_html=True)
st.markdown("---")

# --- DATABASE SETUP ---
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

db = load_db()

# --- HELPER FOR CHARTS ---
def extract_number(val):
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        match = re.search(r"[-+]?\d*\.\d+|\d+", val.replace(',', '.'))
        if match:
            return float(match.group())
    return None

# --- DYNAMIC STATE VARIABLES FOR NEW PARAMETERS ---
if "custom_params" not in st.session_state:
    st.session_state.custom_params = []

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

# --- SIDEBAR NAVIGATION & SETTINGS ---
st.sidebar.markdown("### 🧭 NAVIGATION")
app_mode = st.sidebar.radio("Go to:", ["📡 Scanner", "📚 Database", "📊 Product Comparison"])
st.sidebar.markdown("---")

# --- API KEY HANDLING & DYNAMIC MODEL SELECTOR ---
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
            st.sidebar.caption("Recommendation: Use a 'flash' model for fast PDF scans.")
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
    st.info("💡 The scanner extracts pure facts from the PDF. Qualitative AI analysis (strengths/weaknesses) is activated in the 'Product Comparison' tab.")
    
    uploaded_files = st.file_uploader("Drop brochures or datasheets here (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

    machine_configs = {}
    if uploaded_files:
        if not api_key:
            st.error("ACCESS DENIED: Please provide an API Key.")
        else:
            st.success("Files detected. Please assign models.")
            for file in uploaded_files:
                sub_col1, sub_col2 = st.columns(2)
                with sub_col1:
                    default_name = file.name.rsplit('.', 1)[0]
                    m_name = st.text_input(f"Machine Name:", value=default_name, key=f"name_{file.name}")
                with sub_col2:
                    m_cat = st.selectbox(f"Baseline Class (Sany):", options=CATEGORIES, key=f"cat_{file.name}")
                machine_configs[file.name] = {"name": m_name, "category": m_cat}

    if st.button("🚀 INITIATE AI SCAN (EXTRACT DATA)", type="primary", use_container_width=True):
        if not uploaded_files:
            st.error("ERROR: No files selected for scanning.")
        elif not selected_parameters:
            st.error("ERROR: No scan parameters selected.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            has_error = False
            
            for index, file in enumerate(uploaded_files):
                current_config = machine_configs[file.name]
                current_machine = current_config["name"]
                current_category = current_config["category"]
                
                status_text.text(f"Scanning data for '{current_machine}' ({current_category})...")
                
                file.seek(0)
                file_bytes = file.read()
                file_part = {"mime_type": file.type, "data": file_bytes}
                
                prompt = f"""
                You are a precise technical data extraction assistant.
                Focus EXCLUSIVELY on the technical data for the specific model/series: "{current_machine}".
                Extract exact values for: {json.dumps(selected_parameters, ensure_ascii=False)}
                Requirements:
                1. Valid JSON object only.
                2. Exact matching keys.
                3. Use "?" if missing.
                4. UNIT CONVERSION: Convert imperial to metric (kW, kg, liters, mm). No imperial units.
                Respond ONLY with raw JSON format.
                """
                
                try:
                    model = genai.GenerativeModel(selected_model_name)
                    
                    safety_settings = [
                        {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
                        {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"}
                    ]
                    
                    response = model.generate_content(
                        contents=[file_part, prompt],
                        generation_config={"response_mime_type": "application/json"},
                        safety_settings=safety_settings
                    )
                    
                    if not response.candidates:
                        raise Exception("Google Safety Filter blocked the response! Try again.")
                        
                    raw_text = response.text.strip()
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    elif raw_text.startswith("```"):
                        raw_text = raw_text[3:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                        
                    raw_text = raw_text.strip()
                    extracted_json = json.loads(raw_text)
                    
                    extracted_json["Machine"] = current_machine
                    extracted_json["Category"] = current_category
                    
                    unique_id = f"{current_machine} ({current_category})"
                    db[unique_id] = extracted_json
                    save_db(db)
                    
                except Exception as e:
                    error_msg = str(e)
                    if "429" in error_msg:
                        match = re.search(r"seconds:\s*(\d+)", error_msg)
                        wait_seconds = int(match.group(1)) + 2 if match else 45
                        
                        countdown_box = st.empty()
                        for i in range(wait_seconds, 0, -1):
                            countdown_box.error(f"🛑 API limit reached! Scanner cooling down... Please wait: {i} seconds.")
                            time.sleep(1)
                        countdown_box.success("🟢 System ready! You can initiate the scan again.")
                        has_error = True
                    else:
                        st.error(f"❌ Scan failed for '{current_machine}'. Error: {error_msg}")
                        has_error = True
                
                progress_bar.progress((index + 1) / len(uploaded_files))
            
            if not has_error:
                status_text.text("✅ Data extraction complete. Machines saved to database.")
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
        
        if db: 
            db_rows = list(db.values())
            df_lib = pd.DataFrame(db_rows)
            df_lib = df_lib.rename(columns={'Category': 'Class'})
            
            cols = ['Machine', 'Class'] + [c for c in df_lib.columns if c not in ['Machine', 'Class']]
            df_lib = df_lib[cols]
            
            if lib_filter != "All":
                df_lib = df_lib[df_lib['Class'] == lib_filter]
                
            st.dataframe(df_lib, use_container_width=True)
            
            excel_buffer_lib = io.BytesIO()
            with pd.ExcelWriter(excel_buffer_lib, engine='openpyxl') as writer:
                df_lib.to_excel(writer, index=False, sheet_name='Master_Database')
            st.download_button("📥 DOWNLOAD DATABASE AS EXCEL (.xlsx)", excel_buffer_lib.getvalue(), "RivalRadar_Database.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.info("The database is currently empty. Scan datasheets first.")

# ================= VIEW 3: PRODUCT COMPARISON =================
elif app_mode == "📊 Product Comparison":
    st.markdown("### 📊 PRODUCT COMPARISON (COMPETITIVE ANALYSIS)")
    if not db:
        st.info("No data available yet. Load machines into the database via the scanner first!")
    else:
        db_keys = list(db.keys())
        
        arena_col1, arena_col2 = st.columns(2)
        with arena_col1:
            baseline_sel = st.selectbox("🟢 Own Product (Sany Baseline):", options=db_keys)
        with arena_col2:
            competitors_sel = st.multiselect("🔴 Competitor Models:", options=[k for k in db_keys if k != baseline_sel])
            
        st.markdown("#### ✨ ACTIVATE AI ANALYSES")
        pwr_charts = st.checkbox("📊 Visual Performance Comparison (Generate bar charts)")
        pwr_ampel = st.checkbox("🚦 Strengths/Weaknesses Profile (Traffic light system)")
        pwr_pitch = st.checkbox("💬 Sales Arguments (Generate Elevator Pitch)")
        
        if st.button("⚖️ GENERATE COMPARISON", type="primary", use_container_width=True):
            if not competitors_sel:
                st.warning("Please select at least one competitor model to start the comparison!")
            else:
                battle_roster = [baseline_sel] + competitors_sel
                battle_data = [db[k] for k in battle_roster]
                
                df_battle = pd.DataFrame(battle_data)
                if 'Category' in df_battle.columns:
                    df_battle = df_battle.drop(columns=['Category'])
                df_battle_t = df_battle.set_index("Machine").T
                
                st.markdown("---")
                st.write("### 🗄️ Raw Data Overview")
                st.dataframe(df_battle_t, use_container_width=True)
                
                if pwr_charts:
                    st.markdown("---")
                    st.write("### 📊 Visual Performance Comparison")
                    
                    chart_metrics = ["Operating weight (kg)", "Engine Power STD (kW)", "Max Digging depth", "Breakout force (kN)"]
                    
                    chart_cols = st.columns(2)
                    col_idx = 0
                    
                    for metric in chart_metrics:
                        if metric in df_battle.columns:
                            chart_data = []
                            for _, row in df_battle.iterrows():
                                val = extract_number(row.get(metric))
                                if val is not None:
                                    chart_data.append({"Machine": row['Machine'], metric: val})
                            
                            if chart_data:
                                df_chart = pd.DataFrame(chart_data)
                                
                                chart = alt.Chart(df_chart).mark_bar().encode(
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
                        
                        sys_prompt = f"You are a Senior Sales Strategist for construction machinery. Analyze this comparison in English. Our own product (baseline) is '{db[baseline_sel]['Machine']}'. The competitors are: {', '.join([db[k]['Machine'] for k in competitors_sel])}. Here is the data as JSON: {json.dumps(battle_data)}."
                        
                        reqs = []
                        if pwr_ampel:
                            reqs.append("- Provide an objective assessment of the most important parameters compared to the competition (Use the traffic light system: 🟢 We are superior, 🟡 Tie/Similar, 🔴 Competitor is superior).")
                        if pwr_pitch:
                            reqs.append("- Formulate punchy sales arguments (Elevator Pitch): Why should the customer choose our machine in a direct comparison? Where do sales reps need to be careful in their argumentation?")
                            
                        sys_prompt += "\n\nRequirements:\n" + "\n".join(reqs)
                        
                        try:
                            model = genai.GenerativeModel(selected_model_name)
                            response = model.generate_content(sys_prompt)
                            st.markdown(response.text)
                        except Exception as e:
                            st.error(f"AI Analysis failed: {str(e)}")
