# app/ui/tab_kanban_drag.py

import streamlit as st
import json

DEFAULT_KANBAN = {
    "BACKLOG": [
        {"text": "Heatmap aktivitas dokter", "label": "Feature"},
        {"text": "Heatmap beban poli", "label": "Feature"},
        {"text": "Dropdown otomatis di template", "label": "Improvement"},
    ],
    "READY": [
        {"text": "Optimasi performa Excel Writer", "label": "Improvement"},
        {"text": "Toggle merge shift dokter", "label": "Feature"},
    ],
    "IN PROGRESS": [
        {"text": "Stabilitas merge shift dokter", "label": "Improvement"},
    ],
    "TESTING": [
        {"text": "Akurasi rekap layanan", "label": "Bug"},
    ],
    "DONE": [
        {"text": "Download Template Excel", "label": "Feature"},
    ],
}


# ============================================================
# SESSION STORAGE — bukan penyimpanan file
# ============================================================
def get_server_kanban():
    if "kanban_data" not in st.session_state:
        st.session_state["kanban_data"] = DEFAULT_KANBAN.copy()
    return st.session_state["kanban_data"]


def save_server_kanban(data):
    st.session_state["kanban_data"] = data


# ============================================================
# RENDER TAB KANBAN
# ============================================================
def render_drag_kanban():
    st.title("📌 Kanban — Drag & Drop (Tanpa File System)")

    server_data = get_server_kanban()

    left, right = st.columns([1, 3])

    # ============================================================
    # LEFT PANEL — server controls
    # ============================================================
    with left:
        st.subheader("Server Controls")

        # ============= Upload JSON ============
        uploaded = st.file_uploader("Upload Kanban JSON", type=["json"])
        if uploaded:
            try:
                loaded = json.load(uploaded)

                # Normalisasi
                for col, items in loaded.items():
                    fixed = []
                    for item in items:
                        if isinstance(item, dict) and "text" in item:
                            if "label" not in item:
                                item["label"] = "Feature"
                            fixed.append(item)
                        elif isinstance(item, str):
                            fixed.append({"text": item, "label": "Feature"})
                    loaded[col] = fixed

                save_server_kanban(loaded)
                st.success("Kanban JSON berhasil di-load ke server (session).")
                st.experimental_rerun()

            except Exception as e:
                st.error(f"JSON tidak valid: {e}")

        # ============= Tambah card baru ============
        st.markdown("---")
        st.markdown("### ➕ Tambah Card")
        with st.form("add_card_form"):
            new_text = st.text_input("Judul")
            new_label = st.selectbox("Label", ["Feature", "Improvement", "Bug"])
            target = st.selectbox("Kolom", list(server_data.keys()))
            add_btn = st.form_submit_button("Tambah")

            if add_btn:
                if new_text.strip():
                    server_data[target].append({"text": new_text, "label": new_label})
                    save_server_kanban(server_data)
                    st.success("Ditambahkan!")
                    st.experimental_rerun()

        # ============= Download server JSON ============
        st.markdown("---")
        st.markdown("### 📥 Download Server JSON")
        st.download_button(
            "Download kanban.json",
            data=json.dumps(server_data, indent=2, ensure_ascii=False),
            file_name="kanban.json",
            mime="application/json"
        )

    # ============================================================
    # RIGHT PANEL — Kanban HTML Drag & Drop
    # ============================================================
    with right:
        st.subheader("Kanban Board (Client-side Drag & Drop)")

        initial_state = json.dumps(server_data, ensure_ascii=False)

        html = f"""
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <style>
    body {{ font-family: Arial; }}
    .board {{ display:flex; gap:14px; }}
    .column {{ width:260px; padding:10px; background:#f4f6f8; border-radius:8px; }}
    .col-title {{ font-weight:bold; margin-bottom:6px; }}
    .card {{ background:#fff; padding:8px; border-radius:6px; margin-bottom:8px;
             box-shadow:0 1px 2px rgba(0,0,0,0.1); }}
    .label {{ font-size:11px; padding:2px 6px; color:white; border-radius:5px; }}
    .Feature {{ background:#2f80ed; }}
    .Improvement {{ background:#27ae60; }}
    .Bug {{ background:#ff4d4f; }}
  </style>
</head>
<body>

<div>
  <button id="exportBtn">Export JSON</button>
</div>

<div id="board" class="board"></div>

<script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/Sortable.min.js"></script>
<script>
  const initial = {initial_state};

  function buildCard(item) {{
    let div = document.createElement("div");
    div.className = "card";

    let label = document.createElement("div");
    label.className = "label " + item.label;
    label.innerText = item.label;

    let title = document.createElement("div");
    title.innerText = item.text;

    div.appendChild(label);
    div.appendChild(title);
    return div;
  }}

  function render(state) {{
    const board = document.getElementById("board");
    board.innerHTML = "";

    for (const col of Object.keys(state)) {{
      let colDiv = document.createElement("div");
      colDiv.className = "column";

      let title = document.createElement("div");
      title.className = "col-title";
      title.innerText = col;

      let list = document.createElement("div");
      list.className = "list";
      list.dataset.col = col;

      state[col].forEach(item => list.appendChild(buildCard(item)));

      colDiv.appendChild(title);
      colDiv.appendChild(list);
      board.appendChild(colDiv);

      new Sortable(list, {{
        group: "kanban",
        animation: 150
      }});
    }}
  }}

  function exportData() {{
    const lists = document.querySelectorAll(".list");
    let data = {{}};

    lists.forEach(list => {{
      let col = list.dataset.col;
      data[col] = [];
      list.childNodes.forEach(card => {{
        const label = card.children[0].innerText;
        const text = card.children[1].innerText;
        data[col].push({{text:text, label:label}});
      }});
    }});

    const blob = new Blob([JSON.stringify(data, null, 2)], {{type:"application/json"}});
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "kanban_export.json";
    a.click();
  }}

  document.getElementById("exportBtn").onclick = exportData;

  render(initial);
</script>

</body>
</html>
"""

        st.components.v1.html(html, height=650, scrolling=True)
