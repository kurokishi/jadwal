###############################################################
#  JADWAL.PY — MAIN STREAMLIT APP
#  Final Version (Path Fix + Tab System + Full Module Loader)
###############################################################

import os
import sys

# ============================================================
# PATH FIX → memastikan import "app.*" selalu berhasil.
# Ini 100% aman untuk Streamlit Cloud & environment lokal.
# ============================================================

ROOT = os.path.dirname(os.path.abspath(__file__))       # folder berisi jadwal.py
PARENT = os.path.dirname(ROOT)                          # folder di atasnya

if ROOT not in sys.path:
    sys.path.insert(0, ROOT)
if PARENT not in sys.path:
    sys.path.insert(0, PARENT)

print("=== PATH FIX ACTIVE ===")
print("sys.path[0:3] =", sys.path[0:3])
print("CWD =", os.getcwd())

# ============================================================
# IMPORT STREAMLIT + MODULES
# ============================================================

import streamlit as st

# app.config
from app.config import Config

# core modules
from app.core.scheduler import Scheduler
from app.core.cleaner import DataCleaner
from app.core.excel_writer import ExcelWriter
from app.core.time_parser import TimeParser
from app.core.validator import Validator
from app.core.analyzer import ErrorAnalyzer

# UI modules
from app.ui.sidebar import render_sidebar
from app.ui.tab_upload import render_upload_tab
from app.ui.tab_analyzer import render_analyzer_tab
from app.ui.tab_visualization import render_visualization_tab
from app.ui.tab_settings import render_settings_tab
from app.ui.tab_kanban_drag import render_drag_kanban


# ============================================================
# INITIALIZE SESSION CONFIG
# ============================================================

if "config" not in st.session_state:
    st.session_state["config"] = Config()

config = st.session_state["config"]

# ============================================================
# INITIALIZE CORE OBJECTS
# ============================================================

time_parser = TimeParser(
    start_hour=config.start_hour,
    start_minute=config.start_minute,
    interval_minutes=config.interval_minutes
)

cleaner = DataCleaner()
analyzer = ErrorAnalyzer()

scheduler = Scheduler(
    parser=time_parser,
    cleaner=cleaner,
    config=config
)

writer = ExcelWriter(config=config)

# ============================================================
# PAGE SETUP
# ============================================================

st.set_page_config(
    page_title="Jadwal Dokter",
    layout="wide"
)

render_sidebar(config)

st.title("🗓️ Sistem Jadwal Dokter")

# ============================================================
# TABS
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

