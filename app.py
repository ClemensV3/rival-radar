import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import os

# 1. Der Tab-Reiter im Browser
st.set_page_config(page_title="RivalRadar", layout="wide", page_icon="📡")

# 2. Der moderne, cleane Header
st.markdown("<h1 style='text-align: center;'>📡 RIVAL-RADAR</h1>", unsafe_allow_html=True)
st.markdown("<h4 style='text-align: center; color: #10b981; font-weight: normal;'>[ TACTICAL COMPETITOR INTELLIGENCE SCANNER ]</h4>", unsafe_allow_html=True)
st.markdown("---")

# --- DATABASE SETUP ---
DB_FILE = "machine_database.json"

def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r") as f:
            return json.load(f)
    return {}

def save_db(data):
    with open(DB_FILE, "w") as f:
        json.dump(data, f, indent=4)

db = load_db()
CATEGORIES = ["Small HEX", "Medium HEX", "Large HEX", "WHEX", "WL"]

# --- SIDEBAR (Aufgeräumt & Clean) ---
st.sidebar.markdown("### 📡 SYSTEM CONTROL")
api_key = st.sidebar.text_input("🔑 Gemini API Key (Auth)", type="password", help="Insert credentials to unlock scanning")

if api_key:
    genai.configure(api_key=api_key)

with st.sidebar.expander("⚙️ CONFIGURE SCAN PARAMETERS"):
    default_params = "Operating weight (kg)\nEngine Power STD (kW)\nCylinder\nDisplacement\nBreakout force (kN)\nFuel Tank (l)\nMax Cab height\nStandard Stick"
    params_input = st.text_area("Target Metrics:", value=default_params, height=200)
    parameters = [p.strip() for p in params_input.split("\n") if p.strip()]

with st.sidebar.expander("✨ ACTIVATE AI SUPERPOWERS"):
    use_ampel = st.checkbox("🚦 Threat Level Scoring (🟢/🟡/🔴)")
    use_pitch = st.checkbox("💬 Tactical Elevator Pitch")

# --- MAIN AREA ---
col1, col2 = st.columns([1, 1])

with col1:
    st.markdown("### 📥 [1] UPLOAD INTERCEPTED DATA")
    uploaded_files = st.file_uploader("Drop competitor blueprints here (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

    machine_configs = {}
    if uploaded_files:
        if not api_key:
            st.error("ACCESS DENIED: Please authenticate with API Key in the sidebar.")
        else:
            st.success("Files intercepted. Assign target IDs.")
            for file in uploaded_files:
                sub_col1, sub_col2 = st.columns(2)
                with sub_col1:
                    default_name = file.name.rsplit('.', 1)[0]
                    m_name = st.text_input(f"Target ID:", value=default_name, key=f"name_{file.name}")
                with sub_col2:
                    m_cat = st.selectbox(f"Class:", options=CATEGORIES, key=f"cat_{file.name}")
                machine_configs[file.name] = {"name": m_name, "category": m_cat}

with col2:
    st.markdown("### 🗄️ [2] ACCESS RADAR DATABASE")
    selected_db_machines = []
    if db:
        filter_cat = st.selectbox("Filter Database by Class:", ["All"] + CATEGORIES)
        available_machines = []
        for m_id, m_data in db.items():
            if filter_cat == "All" or m_data.get("Category") == filter_cat:
                available_machines.append(f"{m_data['Machine']} ({m_data.get('Category', 'Unknown')})")
                
        selected_labels = st.multiselect("Select known targets to include in scan:", available_machines)
        for label in selected_labels:
            clean_name = label.split(" (")[0]
            selected_db_machines.append(clean_name)
    else:
        st.info("Radar database is empty. Waiting for targets...")

st.markdown("---")

# --- EXECUTION ---
if st.button("🚀 INITIATE RADAR SWEEP & GENERATE MATRIX", type="primary", use_container_width=True):
    if not uploaded_files and not selected_db_machines:
        st.error("ERROR: No targets selected for scanning.")
    elif not api_key and uploaded_files:
        st.error("ERROR: API Auth missing.")
    else:
        all_rows = []
        
        # Load from DB
        for m_name in selected_db_machines:
            for m_id, m_data in db.items():
                if m_data["Machine"] == m_name:
                    all_rows.append(m_data)
                    break
                    
        # Process new
        if uploaded_files:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for index, file in enumerate(uploaded_files):
                current_config = machine_configs[file.name]
                current_machine = current_config["name"]
                current_category = current_config["category"]
                
                status_text.text(f"Scanning target '{current_machine}' ({current_category})...")
                
                file_bytes = file.read()
                file_part = {"mime_type": file.type, "data": file_bytes}
                
                prompt = f"""
                You are a precise technical data extraction assistant.
                Focus EXCLUSIVELY on the technical data for the specific model/series: "{current_machine}".
                Extract exact values for: {json.dumps(parameters, ensure_ascii=False)}
                Requirements:
                1. Valid JSON object only.
                2. Exact matching keys.
                3. Use "?" if missing.
                4. UNIT CONVERSION: Convert imperial (HP, lbs, gallons, inches) to metric (kW, kg, liters, mm). No imperial units in output.
                """
                
                if use_ampel:
                    prompt += "\n5. TRAFFIC LIGHT SCORING: Evaluate values compared to industry standard. Append emoji (🟢 for strong, 🟡 for average, 🔴 for weak)."
                if use_pitch:
                    prompt += "\n6. ELEVATOR PITCH: Add '🌟 Tactical Advantage' and '🎯 Vulnerability'."
                
                prompt += "\nRespond ONLY with raw JSON."
                
                try:
                    model = genai.GenerativeModel('gemini-2.5-flash')
                    response = model.generate_content(
                        contents=[file_part, prompt],
                        generation_config={"response_mime_type": "application/json"}
                    )
                    
                    extracted_json = json.loads(response.text)
                    extracted_json["Machine"] = current_machine
                    extracted_json["Category"] = current_category
                    
                    all_rows.append(extracted_json)
                    db[f"{current_machine}_{current_category}"] = extracted_json
                    save_db(db)
                    
                except Exception as e:
                    st.error(f"❌ Scan failed for '{current_machine}': {str(e)}")
                
                progress_bar.progress((index + 1) / len(uploaded_files))
            
            status_text.text("✅ Radar Sweep Completed. Targets Acquired.")

        # Matrix Generation
        if all_rows:
            df = pd.DataFrame(all_rows)
            if 'Category' in df.columns:
                df = df.drop(columns=['Category'])
                
            df_transposed = df.set_index("Machine").T
            df_transposed.reset_index(inplace=True)
            df_transposed.rename(columns={'index': 'Target Specs'}, inplace=True)
            
            st.write("### 📊 INTERCEPTED MATRIX")
            st.dataframe(df_transposed, use_container_width=True)
            
            excel_buffer = io.BytesIO()
            with pd.ExcelWriter(excel_buffer, engine='openpyxl') as writer:
                df_transposed.to_excel(writer, index=False, sheet_name='Radar Matrix')
            
            st.download_button(
                label="📥 DOWNLOAD INTEL (.xlsx)",
                data=excel_buffer.getvalue(),
                file_name="RivalRadar_Intel.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
