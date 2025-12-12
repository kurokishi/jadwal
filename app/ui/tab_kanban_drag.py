import streamlit as st
from streamlit_sortable import sortable
import json
import os

KANBAN_FILE = "kanban.json"

DEFAULT_KANBAN = {
    "BACKLOG": [
        {"text": "Heatmap aktivitas dokter", "label": "Feature"},
        {"text": "Heatmap beban poli", "label": "Feature"},
        {"text": "Dropdown otomatis di template", "label": "Improvement"},
        {"text": "Validasi real-time upload", "label": "Improvement"},
        {"text": "Integrasi master data dokter", "label": "Improvement"},
        {"text": "Timeline view dokter", "label": "Feature"},
    ],
    "READY": [
        {"text": "Optimasi performa Excel Writer", "label": "Improvement"},
        {"text": "Toggle merge shift dokter", "label": "Feature"},
        {"text": "Watermark Excel", "label": "Feature"},
        {"text": "Analisis trend jam poli per hari", "label": "Feature"},
    ],
    "IN PROGRESS": [
        {"text": "Stabilitas merge shift dokter", "label": "Improvement"},
        {"text": "Penyempurnaan Peta Konflik Visual", "label": "Improvement"},
        {"text": "Sinkronisasi slot generator → writer → analyzer", "label": "Bug"},
    ],
    "TESTING": [
        {"text": "Akurasi rekap layanan", "label": "Bug"},
        {"text": "Logika overload Poleks", "label": "Bug"},
        {"text": "Grafik beban poli kompatibilitas Excel", "label": "Improvement"},
    ],
    "DONE": [
        {"text": "Download Template Excel", "label": "Feature"},
        {"text": "Peta Konflik Visual (Matrix)", "label": "Feature"},
        {"text": "Peak Hour Analysis", "label": "Feature"},
        {"text": "Penggabungan shift dokter otomatis", "label": "Improvement"},
        {"text": "Rekap Layanan", "label": "Feature"},
        {"text": "Rekap Poli", "label": "Feature"},
        {"text": "Rekap Dokter", "label": "Feature"},
        {"text": "Tanpa border antar hari", "label": "Improvement"}
    ]
}


# ---------------- UTILS -----------------

def normalize_kanban(data):
    """Convert old string list to object list automatically."""
    new_data = {}
    for col, items in data.items():
        new_col = []
        for item in items:
            if isinstance(item, str):
                new_col.append({"text": item, "label": "Feature"})
            else:
                new_col.append(item)
        new_data[col] = new_col
    return new_data


def load_kanban():
    if os.path.exists(KANBAN_FILE):
        try:
            with open(KANBAN_FILE, "r") as f:
                data = json.load(f)
                return normalize_kanban(data)
        except:
            return DEFAULT_KANBAN
    return DEFAULT_KANBAN


def save_kanban(data):
    with open(KANBAN_FILE, "w") as f:
        json.dump(data, f, indent=4)


def get_label_color(label):
    return {
        "Bug": "#ff4d4f",          # red
        "Feature": "#2f80ed",      # blue
        "Improvement": "#27ae60"   # green
    }.get(label, "#555")


# ---------------- MAIN UI -----------------

def render_drag_kanban():

    st.title("📌 Kanban Developer — Drag & Drop + Label Warna")
    st.caption("Kategori: 🔴 Bug | 🟦 Feature | 🟩 Improvement")

    # Load Kanban
    if "kanban_data" not in st.session_state:
        st.session_state.kanban_data = load_kanban()

    kanban = st.session_state.kanban_data

    st.markdown("---")

    # ========== FORM TAMBAH ITEM ==========
    with st.expander("➕ Tambah Item Baru"):
        colA, colB, colC = st.columns([3, 2, 2])

        with colA:
            new_item = st.text_input("Nama item:")

        with colB:
            new_label = st.selectbox("Label:", ["Bug", "Feature", "Improvement"])

        with colC:
            target_column = st.selectbox("Tambahkan ke kolom:", list(kanban.keys()))

        if st.button("Tambah"):
            if new_item.strip():
                kanban[target_column].append({
                    "text": new_item.strip(),
                    "label": new_label
                })
                save_kanban(kanban)
                st.success("Item berhasil ditambahkan!")
                st.experimental_rerun()

    st.markdown("---")

    # ========== RENDER 5 KOLOM ==========
    col1, col2, col3, col4, col5 = st.columns(5)
    cols = [col1, col2, col3, col4, col5]
    titles = list(kanban.keys())

    updated_state = {}

    for col, title in zip(cols, titles):
        with col:
            st.subheader(title)

            # Render item text with color
            display_items = [
                f"""<div style="
                    padding:8px;
                    border-radius:6px;
                    background-color:{get_label_color(item['label'])};
                    color:white;
                    margin-bottom:6px;
                    font-size:12px;
                    ">
                    <b>{item['label']}</b><br>{item['text']}
                </div>"""
                for item in kanban[title]
            ]

            sorted_items = sortable(
                display_items,
                key=title,
                style={
                    "backgroundColor": "#f0f2f6",
                    "padding": "10px",
                    "borderRadius": "10px",
                    "minHeight": "260px"
                },
                itemStyle={
                    "padding": "0",
                    "margin": "0",
                    "backgroundColor": "transparent"
                }
            )

            # Map sorted HTML back to original objects
            new_objects = []
            for html in sorted_items:
                for original in kanban[title]:
                    if original["text"] in html:
                        new_objects.append(original)
                        break

            updated_state[title] = new_objects

    # Update Kanban
    st.session_state.kanban_data = updated_state

    # SAVE BUTTON
    if st.button("💾 Simpan Kanban"):
        save_kanban(updated_state)
        st.success("Kanban berhasil disimpan!")
