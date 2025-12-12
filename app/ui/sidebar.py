import streamlit as st


def render_sidebar(config):
    with st.sidebar:
        st.title("⚙️ Pengaturan")

        config.enable_sabtu = st.checkbox(
            "Aktifkan Hari Sabtu",
            value=config.enable_sabtu
        )

        config.auto_fix_errors = st.checkbox(
            "Auto Perbaiki Format Waktu",
            value=config.auto_fix_errors
        )

        config.start_hour = st.slider("Jam Mulai", 5, 12, config.start_hour)
        config.start_minute = st.select_slider("Menit", [0, 15, 30, 45], config.start_minute)

        config.interval_minutes = st.selectbox(
            "Interval",
            [15, 20, 30, 60],
            index=[15, 20, 30, 60].index(config.interval_minutes)
        )

        config.max_poleks_per_slot = st.number_input(
            "Maks Poleks per slot", 1, 50, config.max_poleks_per_slot
        )

        st.caption("Pengaturan tersimpan selama session berjalan.")
