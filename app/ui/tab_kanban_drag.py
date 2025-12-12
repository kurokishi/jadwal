import streamlit as st
from streamlit_sortable import sortable
import json
import os

KANBAN_FILE = "kanban.json"

DEFAULT_KANBAN = {
    "BACKLOG": [
        "Heatmap aktivitas dokter",
        "Heatmap beban poli",
        "Dropdown otomatis di template",
        "Validasi real-time upload",
        "Integrasi master data dokter",
        "Timeline view dokter"
    ],
    "READY": [
        "Optimasi performa Excel Writer",
        "Toggle merge shift dokter",
        "Watermark Excel",
        "Analisis trend jam poli per hari"
    ],
    "IN PROGRESS": [
        "Stabilitas merge shift dokter",
        "Penyempurnaan Peta Konflik Visual",
        "Sinkronisasi slot generator → writer → analyzer"
    ],
    "TESTING": [
        "Akurasi rekap layanan",
        "Logika overload Poleks",
        "Grafik beban poli kompatibilitas Excel"
    ],
    "DONE": [
        "Download Template Excel",
        "Peta Konflik Visual (Matrix)",
        "Peak Hour Analysis",
        "Penggabungan shift dokter otomatis",
        "Rekap Layanan",
        "Rekap Poli",
        "Rekap Dokter",
        "Tanpa border antar hari"
    ]
}


def load_kanban():
    if os.path.exists(KANBAN_FILE):
        try:
            with open(KANBAN_FILE, "r") as f:
                return json.load(f)
        except:
            return DEFAULT_KANBAN
    return DEFAULT_KANBAN


def save_kanban(data):
    with open(KANBAN_FILE, "w") as f:
        json.dump(data, f, indent=4)


def render_drag_kanban():

    st.title("📌 Kanban Developer — Drag & Drop")
    st.caption("Pindahkan kartu antar kolom seperti Trello")

    # Load only once
    if "kanban_data" not in st.session_state:
        st.session_state.kanban_data = load_kanban()

    kanban = st.session_state.kanban_data

    st.markdown("---")

    # Form tambah card baru
    with st.expander("➕ Tambah Item ke Kanban"):
        colA, colB = st.columns([3, 1])
        with colA:
            new_item = st.text_input("Nama item baru:")
        with colB:
            target_column = st.selectbox("Tambahkan ke kolom:", list(kanban.keys()))

        if st.button("Tambah"):
            if new_item.strip():
                kanban[target_column].append(new_item.strip())
                save_kanban(kanban)
                st.success("Item berhasil ditambahkan!")
                st.experimental_rerun()

    st.markdown("---")

    # RENDER 5 KOLOM KANBAN
    col1, col2, col3, col4, col5 = st.columns(5)
    cols = [col1, col2, col3, col4, col5]
    titles = list(kanban.keys())

    updated = {}

    for col, title in zip(cols, titles):
        with col:
            st.subheader(title)

            cards = kanban[title]

            new_cards = sortable(
                cards,
                key=title,
                style={
                    "backgroundColor": "#F0F2F6",
                    "padding": "10px",
                    "borderRadius": "10px",
                    "minHeight": "260px"
                },
                itemStyle={
                    "padding": "12px",
                    "margin": "8px",
                    "backgroundColor": "#FFFFFF",
                    "borderRadius": "8px",
                    "boxShadow": "1px 1px 5px rgba(0,0,0,0.3)"
                }
            )

            updated[title] = new_cards

    # UPDATE STATE
    st.session_state.kanban_data = updated

    # SAVE BUTTON
    if st.button("💾 Simpan Kanban"):
        save_kanban(updated)
        st.success("Kanban berhasil disimpan ke kanban.json!")
