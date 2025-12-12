# app/ui/tab_analyzer.py

import streamlit as st
import pandas as pd

def render_analyzer_tab(analyzer, config):

    st.subheader("🔍 Error Analyzer")

    # Pastikan hari_list sudah benar (Senin–Jumat + opsional Sabtu)
    hari_list = config.hari_list

    st.markdown("Analisis dilakukan untuk setiap sheet berdasarkan format waktu & struktur kolom.")

    uploaded = st.file_uploader(
        "Upload file untuk analisis (opsional, jika berbeda dari file di tab Upload)",
        type=['xlsx'],
        key="analyzer_uploader"
    )

    # =====================================================================
    # Jika user mengupload file baru
    # =====================================================================
    if uploaded is not None:

        # ---- Load sheet aman ----
        def load_sheet(sheet_name):
            try:
                return pd.read_excel(uploaded, sheet_name=sheet_name)
            except Exception:
                return pd.DataFrame()

        df_reg = load_sheet("Reguler")
        df_pol = load_sheet("Poleks")

        # ===== ANALISIS REGULER =====
        if not df_reg.empty:
            rep = analyzer.analyze_sheet(df_reg, hari_list)
            st.markdown("### 📘 Reguler")
            st.text(analyzer.format_report(rep))
        else:
            st.info("Sheet **Reguler** tidak ditemukan atau kosong.")

        # ===== ANALISIS POLEKS =====
        if not df_pol.empty:
            rep2 = analyzer.analyze_sheet(df_pol, hari_list)
            st.markdown("### 📙 Poleks")
            st.text(analyzer.format_report(rep2))
        else:
            st.info("Sheet **Poleks** tidak ditemukan atau kosong.")

        return

    # =====================================================================
    # Jika tidak upload file → gunakan hasil proses (processed_data)
    # =====================================================================
    if 'processed_data' in st.session_state and st.session_state['processed_data'] is not None:

        df = st.session_state['processed_data']

        st.info("Menggunakan *processed_data* hasil tab Upload & Proses.")
        st.dataframe(df.head(20), use_container_width=True)

        rep = analyzer.analyze_sheet(df, hari_list)
        st.text(analyzer.format_report(rep))

    else:
        st.info("Belum ada data. Upload file pada tab **Upload & Proses** atau unggah file baru di atas.")
