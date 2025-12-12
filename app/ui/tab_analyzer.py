# app/ui/tab_analyzer.py
import streamlit as st
import pandas as pd

def render_analyzer_tab(analyzer, config):
    st.subheader("🔍 Error Analyzer")

    uploaded = st.file_uploader("Upload file untuk analisis (jika belum diupload di tab Upload)", type=['xlsx'], key="analyzer_uploader")
    if uploaded is not None:
        try:
            df_reg = pd.read_excel(uploaded, sheet_name='Reguler')
        except Exception:
            df_reg = pd.DataFrame()
        try:
            df_pol = pd.read_excel(uploaded, sheet_name='Poleks')
        except Exception:
            df_pol = pd.DataFrame()

        if not df_reg.empty:
            rep = analyzer.analyze_sheet(df_reg, config.hari_list)
            st.markdown("**Reguler**")
            st.text(analyzer.format_report(rep))
        else:
            st.info("Sheet 'Reguler' tidak ditemukan atau kosong.")

        if not df_pol.empty:
            rep2 = analyzer.analyze_sheet(df_pol, config.hari_list)
            st.markdown("**Poleks**")
            st.text(analyzer.format_report(rep2))
        else:
            st.info("Sheet 'Poleks' tidak ditemukan atau kosong.")
    else:
        # gunakan data yang sudah diproses jika ada
        if 'processed_data' in st.session_state and st.session_state['processed_data'] is not None:
            df = st.session_state['processed_data']
            st.write("Menggunakan data yang sudah diproses.")
            st.dataframe(df.head(20), use_container_width=True)
        else:
            st.info("Tidak ada data. Upload file pada tab Upload & Proses atau gunakan uploader di atas.")
