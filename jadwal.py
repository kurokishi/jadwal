import streamlit as st
import pandas as pd

from app.config import Config
from app.core.scheduler import Scheduler
from app.core.excel_writer import ExcelWriter
from app.core.analyzer import ErrorAnalyzer

from app.ui.sidebar import render_sidebar
from app.ui.tab_upload import render_upload_tab
from app.ui.tab_analyzer import render_analyzer_tab
from app.ui.tab_visualization import render_visualization_tab
from app.ui.tab_settings import render_settings_tab


def main():
    st.set_page_config(
        page_title="🏥 Pengisi Jadwal Poli Modular",
        layout="wide"
    )

    st.title("🏥 Pengisi Jadwal Poli — Modular Version")

    # init config in session
    if "config" not in st.session_state:
        st.session_state.config = Config()

    config = st.session_state.config

    # Sidebar
    render_sidebar(config)

    # Instantiate core modules
    scheduler = Scheduler(config)
    writer = ExcelWriter(config)
    analyzer = ErrorAnalyzer()

    # Tabs
    tab1, tab2, tab3, tab4 = st.tabs([
        "📤 Upload & Proses",
        "🔍 Error Analyzer",
        "📊 Visualisasi",
        "⚙️ Pengaturan"
    ])

    with tab1:
        render_upload_tab(scheduler, writer, analyzer, config)

    with tab2:
        render_analyzer_tab(analyzer, config)

    with tab3:
        render_visualization_tab(config)

    with tab4:
        render_settings_tab(config)


if __name__ == "__main__":
    main()
