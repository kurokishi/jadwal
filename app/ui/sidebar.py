import streamlit as st

def render_sidebar(config):

    with st.sidebar:
        st.title("⚙️ Pengaturan")

        # ======================
        # Hari & Format
        # ======================
        config.enable_sabtu = st.checkbox(
            "Aktifkan Hari Sabtu",
            value=bool(config.enable_sabtu)
        )

        config.auto_fix_errors = st.checkbox(
            "Auto Perbaiki Format Waktu (Cleaner)",
            value=bool(config.auto_fix_errors)
        )

        # ======================
        # Jam & Interval Slot
        # ======================
        col1, col2 = st.columns(2)

        with col1:
            config.start_hour = st.number_input(
                "Jam Mulai",
                min_value=0,
                max_value=23,
                value=int(config.start_hour),
                step=1
            )

        with col2:
            config.start_minute = st.select_slider(
                "Menit Mulai",
                options=[0, 15, 30, 45],
                value=int(config.start_minute)
            )

        config.interval_minutes = st.selectbox(
            "Interval per Slot (menit)",
            options=[10, 15, 20, 30, 60],
            index=[10, 15, 20, 30, 60].index(config.interval_minutes)
        )

        # ======================
        # Slot Poleks
        # ======================
        config.max_poleks_per_slot = st.number_input(
            "Maks Poleks di Slot yang Sama",
            min_value=1,
            max_value=100,
            value=int(config.max_poleks_per_slot)
        )

        # ======================
        # Info
        # ======================
        st.markdown("---")
        st.caption("Pengaturan akan digunakan langsung untuk proses Scheduler & Export Excel.")
