# app/ui/tab_visualization.py
import streamlit as st
import plotly.express as px
import pandas as pd
import re

def render_visualization_tab(config):

    st.subheader("📊 Visualisasi Jadwal")

    # ======================================================
    # CEK DATA
    # ======================================================
    if "processed_data" not in st.session_state or st.session_state["processed_data"] is None:
        st.info("Belum ada data hasil proses. Jalankan proses di tab Upload & Proses.")
        return

    df = st.session_state["processed_data"]

    # Ambil kolom slot waktu valid "HH:MM"
    time_slots = [c for c in df.columns if re.match(r"^\d{2}:\d{2}$", c)]

    if df.empty or len(time_slots) == 0:
        st.warning("Data tidak lengkap untuk divisualisasikan.")
        return

    # ======================================================
    # MENU
    # ======================================================
    viz = st.selectbox("Pilih visualisasi", ["Heatmap", "Tabel", "Statistik"])

    # ======================================================
    # HEATMAP
    # ======================================================
    if viz == "Heatmap":

        records = []
        for _, r in df.iterrows():
            for ts in time_slots:
                val = 0
                if r[ts] == "R":
                    val = 1
                elif r[ts] == "E":
                    val = 2
                elif r[ts] not in ["", None]:
                    # warna lain (overload)
                    val = 3

                records.append({
                    "Hari": r["HARI"],
                    "Waktu": ts,
                    "Val": val
                })

        pivot = (
            pd.DataFrame(records)
            .pivot_table(index="Hari", columns="Waktu", values="Val", aggfunc="max", fill_value=0)
        )

        # urutkan hari sesuai config, tapi hanya yang ada
        valid_order = [h for h in config.hari_list if h in pivot.index]
        pivot = pivot.reindex(valid_order)

        # BARU: warna heatmap sesuai sistem (Putih, Hijau, Biru, Merah)
        color_scale = [
            "white",        # 0 = kosong
            "lightgreen",   # 1 = Reguler
            "lightblue",    # 2 = Poleks
            "red"           # 3 = overload
        ]

        fig = px.imshow(
            pivot,
            labels=dict(x="Waktu", y="Hari", color="Status"),
            aspect="auto",
            color_continuous_scale=color_scale,
            zmin=0,
            zmax=3
        )
        fig.update_layout(height=500, title="Heatmap Jadwal (R=Hijau, E=Biru, Overload=Merah)")

        st.plotly_chart(fig, use_container_width=True)

    # ======================================================
    # TABEL
    # ======================================================
    elif viz == "Tabel":
        st.caption("Menampilkan tabel jadwal lengkap setelah diproses.")
        st.dataframe(df, use_container_width=True)

    # ======================================================
    # STATISTIK
    # ======================================================
    elif viz == "Statistik":

        total_slots = len(df) * len(time_slots)
        total_r = (df[time_slots] == "R").sum().sum()
        total_e = (df[time_slots] == "E").sum().sum()

        colA, colB, colC = st.columns(3)

        with colA:
            st.metric(
                "Total Slot",
                f"{total_slots}",
            )

        with colB:
            st.metric(
                "Total Reguler (R)",
                f"{total_r} slot",
                f"{(total_r / total_slots * 100):.1f}%" if total_slots else "0%"
            )

        with colC:
            st.metric(
                "Total Poleks (E)",
                f"{total_e} slot",
                f"{(total_e / total_slots * 100):.1f}%" if total_slots else "0%"
            )

        st.write("---")

        st.subheader("📈 Rangkuman")
        st.write(f"- **Slot Reguler**: {total_r}")
        st.write(f"- **Slot Poleks**: {total_e}")
        st.write(f"- **Total slot (jadwal × waktu)**: {total_slots}")
