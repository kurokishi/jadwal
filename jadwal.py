import streamlit as st

from app.core.scheduler import Scheduler
from app.core.excel_writer import ExcelWriter
from app.core.analyzer import ErrorAnalyzer
from app.core.cleaner import DataCleaner
from app.config import Config

from app.ui.sidebar import render_sidebar
from app.ui.tab_upload import render_upload_tab
from app.ui.tab_analyzer import render_analyzer_tab
from app.ui.tab_visualization import render_visualization_tab
from app.ui.tab_settings import render_settings_tab
from app.ui.tab_kanban_drag import render_drag_kanban


# =========================================
# KONFIGURASI APLIKASI
# =========================================
class Config:
    def __init__(self):
        self.interval_minutes = 30          # interval slot default
        self.max_poleks_per_slot = 5        # batas overload
        self.slot_times = []                # diisi otomatis dari scheduler


# =========================================
# MAIN ENTRY STREAMLIT
# =========================================
def main():
    st.set_page_config(
        page_title="Aplikasi Jadwal RS",
        layout="wide",
        initial_sidebar_state="expanded",
    )

    st.sidebar.title("Menu Utama")
    st.sidebar.info("Aplikasi Manajemen Jadwal Dokter & Poli")

    # ======================================
    # Inisialisasi objek core
    # ======================================
    config = Config()
    scheduler = Scheduler(config)
    analyzer = Analyzer(config)
    writer = ExcelWriter(config)

    # ======================================
    # Generate slot jam → simpan ke config
    # ======================================
    try:
        slots_dt = scheduler.generate_slots()
        config.slot_times = [s.strftime("%H:%M") for s in slots_dt]
    except Exception as e:
        st.sidebar.error(f"Error slot generator: {e}")
        config.slot_times = []

    st.sidebar.success("Slot waktu berhasil digenerate.")

    # ======================================
    # TABS APLIKASI
    # ======================================
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "📤 Upload Jadwal",
        "📊 Analisis",
        "📁 Hasil",
        "📌 Kanban",
        "📝 Kanban Drag & Drop"
    ])

    # ---------------- TAB 1 ----------------
    with tab1:
        render_upload_tab(scheduler, writer, analyzer, config)

    # ---------------- TAB 2 ----------------
    with tab2:
        st.header("📊 Analisis Jadwal")
        st.info("Fitur analisis akan otomatis aktif ketika jadwal sudah diproses.")

        if "processed_data" in st.session_state:
            df = st.session_state["processed_data"]
            st.dataframe(df, use_container_width=True)
        else:
            st.warning("Belum ada data yang diproses.")

    # ---------------- TAB 3 ----------------
    with tab3:
        st.header("📁 Hasil Jadwal")
        st.info("Download hasil akan muncul setelah jadwal diproses.")

    # ---------------- TAB 4 ----------------
    with tab4:
        st.header("📌 Kanban Developer (Statis)")
        st.write("Gunakan tab 'Kanban Drag & Drop' untuk versi interaktif.")

        st.markdown("""
        **Backlog**  
        - Heatmap aktivitas  
        - Heatmap poli  
        - Dropdown otomatis template  
        """)

    # ---------------- TAB 5 ----------------
    with tab5:
        render_drag_kanban()


# =========================================
# ENTRY POINT
# =========================================
if __name__ == "__main__":
    main()
