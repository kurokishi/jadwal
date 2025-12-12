import streamlit as st

def render_sidebar(config):
    with st.sidebar:
        st.title("⚙️ Pengaturan Jadwal")
        
        # Info status
        st.caption("Pengaturan mempengaruhi proses scheduling dan export")
        
        # ======================
        # Hari & Format
        # ======================
        st.subheader("📅 Hari Kerja")
        
        config.enable_sabtu = st.checkbox(
            "Aktifkan Hari Sabtu",
            value=bool(config.enable_sabtu),
            help="Tambahkan hari Sabtu dalam jadwal"
        )

        config.auto_fix_errors = st.checkbox(
            "Auto Perbaiki Format Waktu",
            value=bool(config.auto_fix_errors),
            help="Otomatis perbaiki format waktu yang tidak standar"
        )

        # ======================
        # Jam & Interval Slot
        # ======================
        st.subheader("⏰ Waktu Jadwal")
        
        col1, col2 = st.columns(2)

        with col1:
            config.start_hour = st.number_input(
                "Jam Mulai",
                min_value=0,
                max_value=23,
                value=int(config.start_hour),
                step=1,
                help="Jam mulai praktek pertama"
            )

        with col2:
            config.start_minute = st.select_slider(
                "Menit Mulai",
                options=[0, 15, 30, 45],
                value=int(config.start_minute),
                help="Menit mulai praktek pertama"
            )

        config.interval_minutes = st.selectbox(
            "Durasi per Slot (menit)",
            options=[10, 15, 20, 30, 60],
            index=[10, 15, 20, 30, 60].index(config.interval_minutes) 
            if config.interval_minutes in [10, 15, 20, 30, 60] else 3,
            help="Durasi setiap slot waktu dalam jadwal"
        )

        # ======================
        # Batasan Poleks
        # ======================
        st.subheader("📊 Batasan Poleks")
        
        config.max_poleks_per_slot = st.number_input(
            "Maks Poleks per Slot",
            min_value=1,
            max_value=20,
            value=int(config.max_poleks_per_slot),
            help="Jumlah maksimal dokter poleks di slot waktu yang sama"
        )

        # ======================
        # Info & Actions
        # ======================
        st.divider()
        
        with st.expander("ℹ️ Info Aplikasi"):
            st.write(f"**Versi:** 1.0.0")
            st.write(f"**Hari aktif:** {len(config.hari_list)} hari")
            st.write(f"**Slot waktu:** {config.start_hour:02d}:{config.start_minute:02d} - 14:30")
            
            # Hitung jumlah slot
            from datetime import time
            start_total = config.start_hour * 60 + config.start_minute
            end_total = 14 * 60 + 30  # 14:30
            total_slots = (end_total - start_total) // config.interval_minutes
            st.write(f"**Total slot/hari:** {total_slots}")
        
        # Reset button
        if st.button("🔄 Reset Aplikasi", use_container_width=True):
            for key in list(st.session_state.keys()):
                if key != "config":
                    del st.session_state[key]
            st.success("Aplikasi direset!")
            st.rerun()
        
        st.caption("© 2024 Sistem Jadwal Dokter")
