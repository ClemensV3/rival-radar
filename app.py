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

# --- DEINE HARTE FLOTTEN-LISTE ---
CATEGORIES = [
    "SY10U", "SY16C", "SY18U", "SY18C", "SY19E", "SY20C", "SY26U", "SY26C", 
    "SY35U", "SY35C", "SY50U", "SY60U", "SY60C", "SY75C", "SY80U", "SY95C", 
    "SY135C", "SY155U", "SY215C", "SY235C", "SY265C", "SY305C", "SY365H", 
    "SY390H", "SY500H", "SY750H", "SY980H"
]

# --- API KEY HANDLING ---
if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("🔑 API-Key automatisch geladen!")
else:
    st.sidebar.markdown("### 📡 SYSTEM CONTROL")
    api_key = st.sidebar.text_input("🔑 Gemini API Key (Manuell)", type="password", help="Insert credentials to unlock scanning")

if api_key:
    genai.configure(api_key=api_key)

# --- SIDEBAR: PARAMETER ---
with st.sidebar.expander("⚙️ CONFIGURE SCAN PARAMETERS", expanded=False):
    default_params = """Operating weight (kg)
Engine Power STD (kW)
Engine Power OPT (kW)
Engine STD
Engine OPT
Cylinder
Displacement
STD Bucket Capacity
Breakout force (kN)
Swing force (kNm)
U/min (Swing)
Traction Force (kN)
Fuel Tank (l)
Hydraulik Oil tank (l)
Hydraulik Oil System (l)
AdBlue Tank (l)
Max Cab height
Max width
Swing radius
Ground clearance
Transport lengths
Max Digging depth
Vertical reach
Standard Stick
Long Stick
AUX 1 Flow
AUX 2 Flow
Quick coupler line
Inside Cab (ISO 6396)
Outside Cab (2000/14/EG)
STD Speed (kmh)
OPT Speed (kmh)"""
    params_input = st.text_area("Target Metrics:", value=default_params, height=400)
    parameters = [p.strip() for p in params_input.split("\n") if p.strip()]

with st.sidebar.expander("✨ ACTIVATE AI SUPERPOWERS"):
    use_ampel = st.checkbox("🚦 Threat Level Scoring (🟢/🟡/🔴)")
    use_pitch = st.checkbox("💬 Tactical Elevator Pitch")

# --- DAS REGISTER-SYSTEM (Tabs) ---
tab_scanner, tab_library = st.tabs(["📡 Scanner & Live Matrix", "📚 Hangar / Bibliothek"])

# ================= TAB 1: SCANNER =================
with tab_scanner:
    col1, col2 = st.columns([1, 1])

    with col1:
        st.markdown("### 📥 [1] UPLOAD INTERCEPTED DATA")
        uploaded_files = st.file_uploader("Drop competitor blueprints here (PDF, PNG, JPG)", type=["pdf", "png", "jpg", "jpeg"], accept_multiple_files=True)

        machine_configs = {}
        if uploaded_files:
            if not api_key:
                st.error("ACCESS DENIED: Please authenticate with API Key.")
            else:
                st.success("Files intercepted. Assign target IDs.")
                for file in uploaded_files:
                    sub_col1, sub_col2 = st.columns(2)
                    with sub_col1:
                        default_name = file.name.rsplit('.', 1)[0]
                        m_name = st.text_input(f"Target ID:", value=default_name, key=f"name_{file.name}")
                    with sub_col2:
                        m_cat = st.selectbox(f"Baseline Class:", options=CATEGORIES, key=f"cat_{file.name}")
                    machine_configs[file.name] = {"name": m_name, "category": m_cat}

    with col2:
        st.markdown("### 🗄️ [2] ACCESS RADAR DATABASE")
        selected_db_machines = []
        if db:
            filter_cat = st.selectbox("Filter Database by Baseline:", ["All"] + CATEGORIES)
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
            
            # Aus DB laden
            for m_name in selected_db_machines:
                for m_id, m_data in db.items():
                    if m_data["Machine"] == m_name:
                        all_rows.append(m_data)
                        break
                        
            # Neue PDFs scannen
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
                st.rerun()

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

# ================= TAB 2: BIBLIOTHEK =================
with tab_library:
    st.markdown("### 📚 GEHEIMER HANGAR (ALLE GESCANNTEN MASCHINEN)")
    if db:
        # --- FILTER & LÖSCH-BEREICH ---
        col_filter, col_delete = st.columns(2)
        
        with col_filter:
            lib_filter = st.selectbox("Bibliothek filtern nach Baseline:", ["All"] + CATEGORIES, key="lib_filter_select")
            
        with col_delete:
            delete_options = ["Nichts löschen"] + list(db.keys())
            
            def format_delete_option(opt):
                if opt == "Nichts löschen":
                    return "--- Bitte wählen ---"
                return f"{db[opt]['Machine']} (Baseline: {db[opt].get('Category', 'Unbekannt')})"
                
            to_delete = st.selectbox("🗑️ Fehlerhaften Datensatz löschen:", options=delete_options, format_func=format_delete_option)
            
            if st.button("🧨 DATENSATZ VERNICHTEN", use_container_width=True):
                if to_delete != "Nichts löschen":
                    del db[to_delete]
                    save_db(db)
                    st.rerun() # Lädt die Seite sofort neu und räumt die UI auf
                    
        st.markdown("---")
        
        # --- TABELLEN-ANZEIGE ---
        if db: # Nochmal prüfen, falls gerade der letzte Datensatz gelöscht wurde
            db_rows = list(db.values())
            df_lib = pd.DataFrame(db_rows)
            
            # Sortierung der Spalten
            cols = ['Machine', 'Category'] + [c for c in df_lib.columns if c not in ['Machine', 'Category']]
            df_lib = df_lib[cols]
            
            # Filter anwenden
            if lib_filter != "All":
                df_lib = df_lib[df_lib['Category'] == lib_filter]
                
            st.dataframe(df_lib, use_container_width=True)
            
            # Master-Download
            excel_buffer_lib = io.BytesIO()
            with pd.ExcelWriter(excel_buffer_lib, engine='openpyxl') as writer:
                df_lib.to_excel(writer, index=False, sheet_name='Gesamte_Bibliothek')
                
            st.download_button(
                label="📥 GESAMTE DATENBANK ALS EXCEL DOWNLOADEN (.xlsx)",
                data=excel_buffer_lib.getvalue(),
                file_name="RivalRadar_Master_Datenbank.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("Der Hangar ist leer. Scanne zuerst ein Prospekt im ersten Reiter!")
