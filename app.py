import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Jadwal Poli", layout="wide")

st.title("📅 Sistem Jadwal Poli – Final Version")

# ============================
# 1. UPLOAD EXCEL
# ============================
uploaded = st.file_uploader("Upload file Excel Template", type=["xlsx"])

if uploaded is not None:
    df = pd.read_excel(uploaded)

    # Normalisasi jenis poli
    df["Jenis"] = df["Jenis"].str.strip().str.lower()
    df["Jenis"] = df["Jenis"].replace({
        "reguler": "Reguler",
        "regular": "Reguler",
        "eksekutif": "Eksekutif",
        "executive": "Eksekutif",
        "poleks": "Eksekutif"
    })

    # ============================
    # 2. PARSE RANGE WAKTU
    # ============================
    def expand_range(row):
        start, end = row["Range"].split("-")
        start = pd.to_datetime(start.strip())
        end = pd.to_datetime(end.strip())

        times = []
        t = start
        while t < end:
            times.append(t.strftime("%H.%M"))
            t += pd.Timedelta(minutes=30)

        return times

    df["Jam"] = df.apply(expand_range, axis=1)
    df = df.explode("Jam")

    # ============================
    # 3. PENANDAAN R / E
    # ============================
    df["Kode"] = df["Jenis"].apply(lambda x: "R" if x == "Reguler" else "E")

    # ============================
    # 4. CEK KUOTA (Over 7 dokter untuk Eksekutif)
    # ============================
    df["Over_Kuota"] = False

    for hari in df["Hari"].unique():
        df_hari = df[df["Hari"] == hari]

        for jam in df_hari["Jam"].unique():
            df_slot = df_hari[df_hari["Jam"] == jam]

            # HANYA poleks/eksekutif yang dicek kuota
            eksek_count = len(df_slot[df_slot["Jenis"] == "Eksekutif"])

            if eksek_count > 7:
                df.loc[(df["Hari"] == hari) & (df["Jam"] == jam) & (df["Jenis"] == "Eksekutif"),
                       "Over_Kuota"] = True

    # ============================
    # 5. TABEL WARNA
    # ============================
    def color_row(row):
        if row["Over_Kuota"]:
            return "background-color: red; color: white"

        if row["Kode"] == "R":
            return "background-color: lightgreen"
        else:
            return "background-color: lightblue"

    st.subheader("📋 Jadwal Lengkap (30 Menit Interval)")

    st.dataframe(
        df.style.apply(lambda row: [color_row(row)] * len(row), axis=1),
        use_container_width=True
    )

    # ============================
    # 6. DASHBOARD
    # ============================
    st.subheader("📊 Dashboard – Jumlah Dokter per Hari/Jam")

    pivot = df.pivot_table(
        index="Jam",
        columns="Hari",
        values="Dokter",
        aggfunc="count",
        fill_value=0
    )

    st.dataframe(pivot, use_container_width=True)

    # ============================
    # 7. DOWNLOAD HASIL
    # ============================
    @st.cache_data
    def convert_to_excel(data):
        return data.to_excel(index=False, engine="openpyxl")

    st.download_button(
        label="📥 Download Jadwal Final (Excel)",
        data=df.to_csv(index=False).encode("utf-8"),
        file_name="jadwal_final.csv",
        mime="text/csv"
    )

else:
    st.info("Silakan upload file Excel untuk melanjutkan.")

