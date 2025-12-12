# app/ui/tab_visualization.py
import streamlit as st
import plotly.express as px
import pandas as pd

def render_visualization_tab(config):
    st.subheader("📊 Visualisasi")

    if 'processed_data' not in st.session_state or st.session_state['processed_data'] is None:
        st.info("Belum ada data hasil proses. Jalankan proses di tab Upload & Proses.")
        return

    df = st.session_state['processed_data']
    time_slots = [c for c in df.columns if ':' in c]
    if df.empty or not time_slots:
        st.warning("Data tidak lengkap untuk divisualisasikan.")
        return

    viz = st.selectbox("Pilih visualisasi", ["Heatmap", "Tabel", "Statistik"])

    if viz == "Heatmap":
        # Pivot ke bentuk hari x waktu -> max value (0 kosong, 1 R, 2 E)
        records = []
        for _, r in df.iterrows():
            for ts in time_slots:
                val = 0
                if r[ts] == 'R':
                    val = 1
                elif r[ts] == 'E':
                    val = 2
                records.append({'Hari': r['HARI'], 'Waktu': ts, 'Val': val})

        pivot = pd.DataFrame(records).pivot_table(index='Hari', columns='Waktu', values='Val', aggfunc='max', fill_value=0)
        # urutkan hari menurut config
        pivot = pivot.reindex(config.hari_list)
        fig = px.imshow(pivot, labels=dict(x="Waktu", y="Hari", color="Status"), aspect="auto",
                        color_continuous_scale=['white','lightgreen','lightblue','red'])
        fig.update_layout(height=450, title="Heatmap Jadwal")
        st.plotly_chart(fig, use_container_width=True)

    elif viz == "Tabel":
        st.dataframe(df, use_container_width=True)

    elif viz == "Statistik":
        total_slots = len(df) * len(time_slots)
        total_r = (df[time_slots] == 'R').sum().sum()
        total_e = (df[time_slots] == 'E').sum().sum()
        st.metric("Total Slot", total_slots)
        st.metric("Total Reguler", f"{total_r} ({(total_r/total_slots*100) if total_slots else 0:.1f}%)")
        st.metric("Total Poleks", f"{total_e} ({(total_e/total_slots*100) if total_slots else 0:.1f}%)")
