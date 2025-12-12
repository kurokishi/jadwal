###############################################################
#  JADWAL.PY — MAIN STREAMLIT APP (FINAL VERSION)
#  Path Fix Stabil untuk Streamlit Cloud & Lokal
###############################################################

import os
import sys

# ============================================================
# 1. FIX PYTHON PATH
# ============================================================

# Folder lokasi file jadwal.py
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Tambahkan BASE_DIR ke sys.path agar "app" bisa ditemukan
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

print("=== PATH DEBUG ===")
print("FILE:", __file__)
print("CWD :", os.getcwd())
print("BASE:", BASE_DIR)
print("sys.path[0:3] =", sys.path[:3])
print("Folder isi BASE:", os.listdir(BASE_DIR))

# ============================================================
# STREAMLIT IMPORT
# ============================================================

import streamlit as st

# ============================================================
# APP MODULE IMPORT (Dijamin tidak error)
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
    
    print("✅ Semua modul berhasil diimport")
    
except ImportError as e:
    print(f"❌ Import error: {e}")
    st.error(f"Error import module: {e}")
    
    # Fallback: buat minimal UI untuk debugging
    st.set_page_config(page_title="Jadwal Dokter", layout="wide")
    st.title("⚠️ Error Setup Aplikasi")
    st.error(f"Module import error: {e}")
    st.code(f"Base dir: {BASE_DIR}")
    st.code(f"Files: {os.listdir(BASE_DIR)}")
    if os.path.exists(os.path.join(BASE_DIR, "app")):
        st.code(f"App files: {os.listdir(os.path.join(BASE_DIR, 'app'))}")
    st.stop()

# ============================================================
# SESSION CONFIG INIT
# ============================================================

if "config" not in st.session_state:
    st.session_state["config"] = Config()

config = st.session_state["config"]

# ============================================================
# CORE OBJECT INITIALIZATION
# ============================================================

try:
    # Initialize TimeParser dengan config
    time_parser = TimeParser(
        start_hour=config.start_hour,
        start_minute=config.start_minute,
        interval_minutes=config.interval_minutes
    )
    print(f"✅ TimeParser initialized: start={config.start_hour}:{config.start_minute}, interval={config.interval_minutes}")
    
    # Initialize Cleaner
    cleaner = DataCleaner()
    print("✅ DataCleaner initialized")
    
    # Initialize Analyzer
    analyzer = ErrorAnalyzer()
    print("✅ ErrorAnalyzer initialized")
    
    # Initialize Scheduler
    scheduler = Scheduler(
        parser=time_parser,
        cleaner=cleaner,
        config=config
    )
    print("✅ Scheduler initialized")
    
    # Initialize ExcelWriter
    writer = ExcelWriter(config=config)
    print("✅ ExcelWriter initialized")
    
    # Initialize Validator
    validator = Validator()
    print("✅ Validator initialized")
    
except Exception as e:
    print(f"❌ Error initializing core objects: {e}")
    st.error(f"Error initializing application: {e}")
    st.code(f"Error details: {str(e)}")
    
    # Fallback minimal
    st.set_page_config(page_title="Jadwal Dokter", layout="wide")
    st.title("⚠️ Initialization Error")
    st.error(f"Failed to initialize: {e}")
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
    render_upload_tab(scheduler, writer, analyzer, config)

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
        st.write(f"- Start Time: {config.start_hour}:{config.start_minute}")
        st.write(f"- Interval: {config.interval_minutes} menit")
        st.write(f"- Max Poleks per Slot: {config.max_poleks_per_slot}")
        st.write(f"- Auto Fix Errors: {config.auto_fix_errors}")
        st.write(f"- Enable Sabtu: {config.enable_sabtu}")
        st.write(f"- Hari Order: {config.hari_order}")
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

print("✅ Jadwal.py initialized successfully")
