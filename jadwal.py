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

# ============================================================
# SESSION CONFIG INIT
# ============================================================

if "config" not in st.session_state:
    st.session_state["config"] = Config()

config = st.session_state["config"]

# ============================================================
# CORE OBJECT INITIALIZATION
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
