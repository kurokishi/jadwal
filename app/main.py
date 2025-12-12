# app/main.py

import streamlit as st
import pandas as pd

from app.config import Config
from app.core.scheduler import Scheduler
from app.core.excel_writer import ExcelWriter
from app.ui.sidebar import render_sidebar

def main():
    config = Config()
    render_sidebar(config)

    st.title("🚀 Pengolah Jadwal Poli Modular")

    uploaded = st.file_uploader("Upload Excel", type=['xlsx'])

    if uploaded:
        df_reg = pd.read_excel(uploaded, sheet_name='Reguler')
        df_pol = pd.read_excel(uploaded, sheet_name='Poleks')

        scheduler = Scheduler(config)

        df_r = scheduler.process(df_reg, "Reguler")
        df_e = scheduler.process(df_pol, "Poleks")

        df_final = pd.concat([df_r, df_e]).reset_index(drop=True)

        st.dataframe(df_final)

        slots = [c for c in df_final.columns if ':' in c]

        writer = ExcelWriter(config)
        buf = writer.write(uploaded, df_final, slots)

        st.download_button(
            "Download Excel Hasil",
            data=buf,
            file_name="jadwal_modular.xlsx"
        )

if __name__ == "__main__":
    main()
