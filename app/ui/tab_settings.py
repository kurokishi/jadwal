# app/ui/tab_settings.py
import streamlit as st

def render_settings_tab(config):

    st.subheader("⚙️ Pengaturan Lanjutan")
    st.markdown("Semua perubahan hanya tersimpan pada session, tidak permanen.")

    # =====================================================
    # RESET CONFIG
    # =====================================================
    if st.button("🔄 Reset ke Default", type="primary"):
        from app.config import Config

        # Buat ulang config baru
        new_cfg = Config()
        st.session_state["config"] = new_cfg

        st.success("Konfigurasi berhasil di-reset.")

        # Paksa update assign
        st.session_state["config_updated"] = True
        st.experimental_rerun()

    st.markdown("---")

    # =====================================================
    # TAMPILKAN SEMUA VALUE CONFIG
    # =====================================================
    st.markdown("### 📘 Informasi Konfigurasi Saat Ini")

    col1, col2 = st.columns(2)

    with col1:
        st.write("**Hari aktif**:", ", ".join(config.hari_list))
        st.write("**Enable Sabtu**:", config.enable_sabtu)
        st.write("**Auto Fix Errors**:", config.auto_fix_errors)

    with col2:
        st.write("**Jam Mulai**:", f"{config.start_hour:02d}:{config.start_minute:02d}")
        st.write("**Interval (menit)**:", config.interval_minutes)
        st.write("**Maks Poleks per Slot**:", config.max_poleks_per_slot)

    st.markdown("---")

    st.caption("Tab ini hanya menampilkan & mengelola konfigurasi global aplikasi.")
