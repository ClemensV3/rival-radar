import streamlit as st
import google.generativeai as genai
import pandas as pd
import json
import io
import os
import re

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

# --- HELPER FÜR DIAGRAMME ---
def extract_number(val):
    """Zieht die nackte Zahl aus Strings wie '75 kW' oder '10320 kg' für die Diagramme"""
    if isinstance(val, (int, float)):
        return float(val)
    if isinstance(val, str):
        match = re.search(r"[-+]?\d*\.\d+|\d+", val.replace(',', '.'))
        if match:
            return float(match.group())
    return None

# --- DYNAMISCHE STATE-VARIABLEN FÜR NEUE PARAMETER ---
if "custom_params" not in st.session_state:
    st.session_state.custom_params = []

# --- BASIS DATEN ---
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

# --- API KEY HANDLING ---
if "GEMINI_API_KEY" in st.secrets and st.secrets["GEMINI_API_KEY"]:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("🔑 API-Key automatisch geladen!")
else:
    st.sidebar.markdown("### 📡 SYSTEM CONTROL")
    api_key = st.sidebar.text_input("🔑 Gemini API Key (Manuell)", type="password", help="Insert credentials to unlock scanning")

if api_key:
    genai.configure(api_key=api_key)

# --- SIDEBAR: DYNAMISCHE PARAMETER-KONFIGURATION ---
with st.sidebar.expander("⚙️ CONFIGURE SCAN PARAMETERS", expanded=False):
    st.markdown("#### ➕ Neue Metrik hinzufügen")
    new_param_input = st.text_input("Name des Parameters:", placeholder="z.B. Track width (mm)", key="new_param_field")
    
    if st.button("Parameter speichern", use_container_width=True):
        if new_param_input:
            clean_param = new_param_input.strip()
            if clean_param not in BASE_PARAMS and clean_param not in st.session_state.custom_params:
                st.session_state.custom_params.append(clean_param)
                st.success(f"'{clean_param}' hinzugefügt!")
                st.rerun()
                
    st.markdown("---")
    st.markdown("#### 🎯 Aktive Scan-Metriken")
    all_available_params = BASE_PARAMS + st.session_state.custom_params
    selected_parameters = st.multiselect(
        "Ausgewählte Parameter:",
        options=all_available_params,
        default=all_available_params,
        help="Entferne die Tags, die für den reinen Daten-Scan nicht benötigt werden."
    )

# --- DAS REGISTER-SYSTEM (Tabs) ---
tab_scanner, tab_library, tab_arena = st.tabs(["📡 Data Scanner", "📚 Hangar / Bibliothek", "⚔️ THE ARENA (Vergleich & KI)"])

# ================= TAB 1: SCANNER =================
with tab_scanner:
    st.markdown("### 📥 UPLOAD INTERCEPTED DATA")
    st.info("💡 Der Scanner liest ab sofort nur noch die reinen Fakten aus, um maximale Geschwindigkeit zu garantieren. Die taktische Analyse zündest du im 'Arena'-Reiter!")
    
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

    if st.button("🚀 INITIATE FAST DATA SWEEP", type="primary", use_container_width=True):
        if not uploaded_files:
            st.error("ERROR: No targets selected for scanning.")
        elif not selected_parameters:
            st.error("ERROR: No scan parameters selected.")
        else:
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            for index, file in enumerate(uploaded_files):
                current_config = machine_configs[file.name]
                current_machine = current_config["name"]
                current_category = current_config["category"]
                
                status_text.text(f"Scanning pure data for '{current_machine}' ({current_category})...")
                
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
                    # Fix 1: Wir nutzen das extrem stabile 1.5-flash Modell
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    response = model.generate_content(
                        contents=[file_part, prompt],
                        generation_config={"response_mime_type": "application/json"}
                    )
                    
                    # Fix 2: Der "Staubsauger", falls die KI Markdown mitgibt
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
                    st.error(f"❌ Scan failed for '{current_machine}'. Grund: {str(e)}")
                
                progress_bar.progress((index + 1) / len(uploaded_files))
            
            status_text.text("✅ Data Sweep Completed. Targets stored in Hangar.")
            st.rerun()

# ================= TAB 2: BIBLIOTHEK =================
with tab_library:
    st.markdown("### 📚 GEHEIMER HANGAR (ALLE MASCHINEN)")
    if db:
        col_filter, col_delete = st.columns(2)
        with col_filter:
            lib_filter = st.selectbox("Bibliothek filtern nach Baseline:", ["All"] + CATEGORIES, key="lib_filter_select")
        with col_delete:
            delete_options = ["Nichts löschen"] + list(db.keys())
            to_delete = st.selectbox("🗑️ Fehlerhaften Datensatz löschen:", options=delete_options)
            if st.button("🧨 DATENSATZ VERNICHTEN", use_container_width=True):
                if to_delete != "Nichts löschen":
                    del db[to_delete]
                    save_db(db)
                    st.rerun()
                    
        st.markdown("---")
        
        if db: 
            db_rows = list(db.values())
            df_lib = pd.DataFrame(db_rows)
            cols = ['Machine', 'Category'] + [c for c in df_lib.columns if c not in ['Machine', 'Category']]
            df_lib = df_lib[cols]
            
            if lib_filter != "All":
                df_lib = df_lib[df_lib['Category'] == lib_filter]
                
            st.dataframe(df_lib, use_container_width=True)
            
            excel_buffer_lib = io.BytesIO()
            with pd.ExcelWriter(excel_buffer_lib, engine='openpyxl') as writer:
                df_lib.to_excel(writer, index=False, sheet_name='Gesamte_Bibliothek')
            st.download_button("📥 HANGAR ALS EXCEL DOWNLOADEN (.xlsx)", excel_buffer_lib.getvalue(), "RivalRadar_Hangar.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
    else:
        st.info("Der Hangar ist leer.")

# ================= TAB 3: THE ARENA =================
with tab_arena:
    st.markdown("### ⚔️ THE ARENA: HEAD-TO-HEAD BATTLEGROUND")
    if not db:
        st.info("Die Arena ist geschlossen. Du musst zuerst Maschinen im Scanner in den Hangar laden!")
    else:
        db_keys = list(db.keys())
        
        arena_col1, arena_col2 = st.columns(2)
        with arena_col1:
            baseline_sel = st.selectbox("🟢 Deine Baseline-Maschine (Home):", options=db_keys)
        with arena_col2:
            competitors_sel = st.multiselect("🔴 Die Gegner (Away):", options=[k for k in db_keys if k != baseline_sel])
            
        st.markdown("#### ✨ ACTIVATE AI SUPERPOWERS")
        pwr_charts = st.checkbox("📊 Taktische Diagramme generieren (Performance Charts)")
        pwr_ampel = st.checkbox("🚦 KI Threat Level Scoring (Stärken/Schwächen-Ampel)")
        pwr_pitch = st.checkbox("💬 KI Tactical Elevator Pitch (Vertriebs-Argumente)")
        
        if st.button("⚔️ LET THEM FIGHT (Vergleich zünden)", type="primary", use_container_width=True):
            if not competitors_sel:
                st.warning("Du musst mindestens einen Gegner auswählen, um den Kampf zu starten!")
            else:
                battle_roster = [baseline_sel] + competitors_sel
                battle_data = [db[k] for k in battle_roster]
                
                # Nackte Fakten Matrix
                df_battle = pd.DataFrame(battle_data)
                if 'Category' in df_battle.columns:
                    df_battle = df_battle.drop(columns=['Category'])
                df_battle_t = df_battle.set_index("Machine").T
                
                st.markdown("---")
                st.write("### 🗄️ Nackte Daten-Matrix")
                st.dataframe(df_battle_t, use_container_width=True)
                
                # --- SUPERPOWER 1: DIAGRAMME ---
                if pwr_charts:
                    st.markdown("---")
                    st.write("### 📊 Taktischer Leistungsvergleich")
                    
                    chart_metrics = ["Operating weight (kg)", "Engine Power STD (kW)", "Max Digging depth", "Breakout force (kN)"]
                    
                    chart_cols = st.columns(2)
                    col_idx = 0
                    
                    for metric in chart_metrics:
                        if metric in df_battle.columns:
                            chart_data = {}
                            for _, row in df_battle.iterrows():
                                val = extract_number(row.get(metric))
                                if val is not None:
                                    chart_data[row['Machine']] = val
                            
                            if chart_data:
                                df_chart = pd.DataFrame(list(chart_data.items()), columns=['Machine', metric]).set_index('Machine')
                                with chart_cols[col_idx % 2]:
                                    st.write(f"**{metric}**")
                                    st.bar_chart(df_chart)
                                col_idx += 1

                # --- SUPERPOWER 2 & 3: KI REFEREE ---
                if pwr_ampel or pwr_pitch:
                    st.markdown("---")
                    st.write("### 🧠 KI Ringrichter-Analyse")
                    with st.spinner("Die KI analysiert den Kampf..."):
                        
                        sys_prompt = f"Du bist ein Chief Strategy Officer für Baumaschinen. Analysiere diesen Vergleich. Deine Baseline/Eigenes Produkt ist '{db[baseline_sel]['Machine']}'. Die Konkurrenten sind: {', '.join([db[k]['Machine'] for k in competitors_sel])}. Hier sind die Daten als JSON: {json.dumps(battle_data)}."
                        
                        reqs = []
                        if pwr_ampel:
                            reqs.append("- Gib eine harte Einschätzung (Threat Level) zu den wichtigsten Parametern (Ampel-System: 🟢 Wir sind besser, 🟡 Gleichstand, 🔴 Konkurrenz ist besser).")
                        if pwr_pitch:
                            reqs.append("- Schreibe einen knackigen 'Tactical Elevator Pitch' für den Vertrieb: Warum gewinnt unsere Baseline-Maschine in diesem Line-Up? Wo sind wir angreifbar?")
                            
                        sys_prompt += "\n\nAnforderungen:\n" + "\n".join(reqs)
                        
                        try:
                            # Für die komplexe Text-Analyse nehmen wir das smarte "Pro" Modell
                            model = genai.GenerativeModel('gemini-1.5-pro')
                            response = model.generate_content(sys_prompt)
                            st.markdown(response.text)
                        except Exception as e:
                            st.error(f"KI Analyse fehlgeschlagen: {str(e)}")
