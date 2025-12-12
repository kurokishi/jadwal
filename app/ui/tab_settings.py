# app/ui/tab_settings.py
import streamlit as st

def render_settings_tab(config):
    st.subheader("⚙️ Pengaturan Lanjutan")
    st.markdown("Atur preferensi aplikasi. Perubahan disimpan di session saja.")

    if st.button("🔄 Reset ke default"):
        # reset: buat ulang dataclass
        from app.config import Config
        st.session_state['config'] = Config()
        st.success("Reset ke default. Silakan refresh (F5) halaman jika diperlukan.")

    st.markdown("### Hari yang aktif:")
    st.write(", ".join(config.hari_list))
