###############################################################
#  JADWAL.PY — MAIN STREAMLIT APP (FINAL VERSION)
###############################################################

import os
import sys
import traceback

# ============================================================
# 1. FIX PYTHON PATH
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

print("=== DEBUG START ===")
print("BASE_DIR:", BASE_DIR)
print("Current files:", os.listdir(BASE_DIR))

# ============================================================
# STREAMLIT IMPORT
# ============================================================

import streamlit as st

# ============================================================
# APP MODULE IMPORT dengan error handling
# ============================================================

try:
    from app.config import Config
    from app.core.scheduler import Scheduler
    from app.core.cleaner import DataCleaner
    from app.core.excel_writer import ExcelWriter
    from app.core.time_parser import TimeParser
    from app.core.validator import Validator
    from app.core.analyzer import ErrorAnalyzer

    from app.ui.sidebar import render_sidebar
    from app.ui.tab_upload import render_upload_tab
    from app.ui.tab_analyzer import render_analyzer_tab
    from app.ui.tab_visualization import render_visualization_tab
    from app.ui.tab_settings import render_settings_tab
    from app.ui.tab_kanban_drag import render_drag_kanban
    
    print("✅ All modules imported successfully")
    
except ImportError as e:
    st.error(f"❌ Import Error: {e}")
    st.code(traceback.format_exc())
    st.stop()

# ============================================================
# SESSION CONFIG INIT
# ============================================================

if "config" not in st.session_state:
    st.session_state["config"] = Config()

config = st.session_state["config"]

# ============================================================
# CORE OBJECT INITIALIZATION dengan debug
# ============================================================

try:
    print(f"🕐 Initializing TimeParser: start={config.start_hour}:{config.start_minute}, interval={config.interval_minutes}")
    
    # Inisialisasi TimeParser HANYA SEKALI
    time_parser = TimeParser(
        start_hour=config.start_hour,
        start_minute=config.start_minute,
        interval_minutes=config.interval_minutes
    )
    
    # Test TimeParser dengan generate slots
    slots = time_parser.generate_slot_strings()
    print(f"🕐 Generated {len(slots)} time slots: {slots[:3]}...")
    
    # Inisialisasi objek lainnya
    cleaner = DataCleaner()
    print("✅ DataCleaner initialized")
    
    analyzer = ErrorAnalyzer()
    print("✅ ErrorAnalyzer initialized")
    
    scheduler = Scheduler(
        parser=time_parser,
        cleaner=cleaner,
        config=config
    )
    print("✅ Scheduler initialized")
    
    writer = ExcelWriter(config=config)
    print("✅ ExcelWriter initialized")
    
    validator = Validator()
    print("✅ Validator initialized")
    
    print("✅ All core objects initialized successfully")
    
except Exception as e:
    print(f"❌ Error initializing core objects: {e}")
    st.error(f"Error initializing application: {e}")
    st.code(traceback.format_exc())
    st.stop()

# ============================================================
# PAGE SETUP
# ============================================================

st.set_page_config(
    page_title="Jadwal Dokter",
    layout="wide",
    page_icon="🗓️"
)

# ============================================================
# SIDEBAR
# ============================================================

render_sidebar(config)

# ============================================================
# MAIN CONTENT
# ============================================================

st.title("🗓️ Sistem Jadwal Dokter")
st.caption("Aplikasi untuk mengelola jadwal dokter reguler dan poleks")

# ============================================================
# TAB SYSTEM
# ============================================================

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📤 Upload & Proses",
    "🔍 Analyzer",
    "📊 Visualisasi",
    "🛠️ Settings",
    "📌 Kanban"
])

with tab1:
    # PASS VALIDATOR ke render_upload_tab
    render_upload_tab(scheduler, writer, analyzer, validator, config)

with tab2:
    render_analyzer_tab(analyzer, config)

with tab3:
    render_visualization_tab(config)

with tab4:
    render_settings_tab(config)

with tab5:
    render_drag_kanban()

# ============================================================
# FOOTER & DEBUG INFO
# ============================================================

st.divider()

with st.expander("ℹ️ Debug Information"):
    col1, col2 = st.columns(2)
    
    with col1:
        st.write("**Session State Keys:**")
        for key in st.session_state.keys():
            st.write(f"- {key}")
    
    with col2:
        st.write("**Configuration:**")
        st.write(f"- Start Time: {config.start_hour:02d}:{config.start_minute:02d}")
        st.write(f"- Interval: {config.interval_minutes} menit")
        st.write(f"- Max Poleks per Slot: {config.max_poleks_per_slot}")
        st.write(f"- Auto Fix Errors: {config.auto_fix_errors}")
        st.write(f"- Enable Sabtu: {config.enable_sabtu}")
        st.write(f"- Hari List: {config.hari_list}")

# ============================================================
# STYLE CUSTOMIZATION
# ============================================================

st.markdown("""
<style>
    /* Main container */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    
    /* Tabs styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        height: 50px;
        white-space: pre-wrap;
        background-color: #f0f2f6;
        border-radius: 8px 8px 0px 0px;
        gap: 8px;
        padding-top: 10px;
        padding-bottom: 10px;
    }
    
    .stTabs [aria-selected="true"] {
        background-color: #ffffff;
        font-weight: bold;
    }
    
    /* Button styling */
    .stButton button {
        border-radius: 8px;
        padding: 0.5rem 1rem;
        font-weight: 500;
    }
    
    /* Success/Error messages */
    .stAlert {
        border-radius: 8px;
    }
    
    /* Dataframe styling */
    .dataframe {
        border-radius: 8px;
    }
</style>
""", unsafe_allow_html=True)

print("✅ Jadwal.py initialized and running successfully")
