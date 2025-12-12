# app/ui/tab_kanban_drag.py
import streamlit as st
import json
import os
from pathlib import Path

KANBAN_FILE = "kanban.json"

DEFAULT_KANBAN = {
    "BACKLOG": [
        {"text": "Heatmap aktivitas dokter", "label": "Feature"},
        {"text": "Heatmap beban poli", "label": "Feature"},
        {"text": "Dropdown otomatis di template", "label": "Improvement"}
    ],
    "READY": [
        {"text": "Optimasi performa Excel Writer", "label": "Improvement"},
        {"text": "Toggle merge shift dokter", "label": "Feature"}
    ],
    "IN PROGRESS": [
        {"text": "Stabilitas merge shift dokter", "label": "Improvement"},
    ],
    "TESTING": [
        {"text": "Akurasi rekap layanan", "label": "Bug"},
    ],
    "DONE": [
        {"text": "Download Template Excel", "label": "Feature"},
    ]
}


def _ensure_kanban_file_exists():
    if not Path(KANBAN_FILE).exists():
        with open(KANBAN_FILE, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_KANBAN, f, indent=2, ensure_ascii=False)


def load_server_kanban():
    _ensure_kanban_file_exists()
    try:
        with open(KANBAN_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
            # normalize: if old items are strings, convert
            for k, v in list(data.items()):
                new_list = []
                for item in v:
                    if isinstance(item, str):
                        new_list.append({"text": item, "label": "Feature"})
                    elif isinstance(item, dict) and "text" in item:
                        if "label" not in item:
                            item["label"] = "Feature"
                        new_list.append(item)
                    else:
                        # ignore invalid
                        pass
                data[k] = new_list
            return data
    except Exception as e:
        st.error(f"Gagal load kanban.json: {e}")
        return DEFAULT_KANBAN.copy()


def save_server_kanban(data):
    try:
        with open(KANBAN_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True, None
    except Exception as e:
        return False, str(e)


def render_drag_kanban():
    st.title("📌 Kanban — Drag & Drop (HTML + JS)")
    st.markdown(
        "Drag & drop menggunakan SortableJS (CDN). Untuk menyimpan perubahan ke server: "
        "1) klik `Export JSON` di area Kanban → 2) Upload file JSON di panel 'Server Controls' lalu klik 'Save to server'."
    )

    # load server kanban (server-side state)
    server_data = load_server_kanban()

    # --- Left column: Server controls (Upload / add / download saved) ---
    left, right = st.columns([1, 3])

    with left:
        st.subheader("Server Controls")

        # Upload JSON to update server kanban (user exported JSON from client)
        uploaded = st.file_uploader("Upload Kanban JSON (to save server-side)", type=["json"])
        if uploaded is not None:
            try:
                uploaded_data = json.load(uploaded)
                # normalize uploaded_data
                for k, v in list(uploaded_data.items()):
                    new_list = []
                    for item in v:
                        if isinstance(item, str):
                            new_list.append({"text": item, "label": "Feature"})
                        elif isinstance(item, dict) and "text" in item:
                            if "label" not in item:
                                item["label"] = "Feature"
                            new_list.append(item)
                    uploaded_data[k] = new_list
                ok, err = save_server_kanban(uploaded_data)
                if ok:
                    st.success("Kanban JSON berhasil disimpan ke server (kanban.json).")
                    st.experimental_rerun()
                else:
                    st.error(f"Gagal simpan: {err}")
            except Exception as e:
                st.error(f"File JSON tidak valid: {e}")

        # Add new item via Streamlit form (server side)
        with st.form("add_card_form"):
            st.markdown("**Tambah item server-side**")
            col1, col2 = st.columns([3, 2])
            with col1:
                new_text = st.text_input("Judul item")
            with col2:
                new_label = st.selectbox("Label", ["Feature", "Improvement", "Bug"])
            target_col = st.selectbox("Kolom tujuan", list(server_data.keys()))
            submitted = st.form_submit_button("Tambah item ke server")
            if submitted:
                if new_text.strip():
                    server_data[target_col].append({"text": new_text.strip(), "label": new_label})
                    ok, err = save_server_kanban(server_data)
                    if ok:
                        st.success("Item ditambahkan & disimpan ke server.")
                        st.experimental_rerun()
                    else:
                        st.error(f"Gagal simpan: {err}")
                else:
                    st.warning("Masukkan judul item terlebih dahulu.")

        # Download server-side kanban.json
        st.markdown("---")
        st.markdown("**Server kanban.json**")
        try:
            with open(KANBAN_FILE, "r", encoding="utf-8") as f:
                data_text = f.read()
            st.download_button("📥 Download server kanban.json", data=data_text, file_name="kanban.json", mime="application/json")
        except Exception:
            st.info("File server tidak tersedia.")

        st.markdown("---")
        st.caption("Catatan: drag/drop di area utama adalah client-side. Gunakan Export JSON lalu Upload untuk menyimpan state hasil drag/drop ke server.")

    # --- Right column: embed HTML kanban (client-side drag & drop) ---
    with right:
        st.subheader("Kanban Board (Client-side)")

        # Pass server_data into the HTML as JSON string
        initial_state = json.dumps(server_data, ensure_ascii=False)

        # HTML + CSS + JS (SortableJS CDN). Provides:
        # - drag & drop between columns
        # - edit card text inline
        # - export JSON (client download)
        # - copy to clipboard
        # - color labels
        html = f"""
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>Kanban</title>
  <style>
    body {{ font-family: Arial, Helvetica, sans-serif; }}
    .board {{ display:flex; gap:12px; align-items:flex-start; }}
    .column {{ background:#f4f6f8; padding:12px; border-radius:8px; width:260px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }}
    .col-title {{ font-weight:700; margin-bottom:8px; }}
    .card-list {{ min-height:120px; }}
    .card {{ background:#fff; padding:8px; margin-bottom:8px; border-radius:6px; box-shadow: 0 1px 2px rgba(0,0,0,0.06); }}
    .label {{ display:inline-block; padding:2px 6px; border-radius:6px; font-size:11px; color:white; margin-bottom:6px; }}
    .label.Bug {{ background:#ff4d4f; }}
    .label.Feature {{ background:#2f80ed; }}
    .label.Improvement {{ background:#27ae60; }}
    .card .title {{ margin:0; font-size:13px; }}
    .toolbar {{ margin-bottom:8px; }}
    .btn {{ padding:6px 10px; border-radius:6px; background:#2f80ed; color:#fff; border:none; cursor:pointer; }}
    .btn.secondary {{ background:#777; }}
    .small {{ font-size:12px; color:#666; }}
    textarea.card-edit {{ width:100%; box-sizing:border-box; }}
  </style>
</head>
<body>
  <div class="toolbar">
    <button class="btn" id="exportBtn">Export JSON</button>
    <button class="btn secondary" id="copyBtn">Copy JSON</button>
  </div>

  <div id="board" class="board"></div>

  <script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/Sortable.min.js"></script>
  <script>
    const initial = {initial_state};

    function createCardElement(item) {{
      const div = document.createElement("div");
      div.className = "card";

      const labelSpan = document.createElement("div");
      labelSpan.className = "label " + (item.label || "Feature");
      labelSpan.innerText = item.label || "Feature";

      const title = document.createElement("div");
      title.className = "title";
      title.innerText = item.text || "";

      const editBtn = document.createElement("button");
      editBtn.textContent = "Edit";
      editBtn.style.marginTop = "6px";
      editBtn.style.fontSize = "11px";

      editBtn.addEventListener("click", function() {{
        // replace title with textarea
        const ta = document.createElement("textarea");
        ta.className = "card-edit";
        ta.value = title.innerText;
        div.replaceChild(ta, title);
        editBtn.style.display = "none";

        ta.addEventListener("blur", function() {{
          title.innerText = ta.value;
          div.replaceChild(title, ta);
          editBtn.style.display = "inline-block";
        }});
      }});

      // container
      div.appendChild(labelSpan);
      div.appendChild(title);
      div.appendChild(editBtn);

      return div;
    }}

    function renderBoard(state) {{
      const board = document.getElementById("board");
      board.innerHTML = "";
      for (const colName of Object.keys(state)) {{
        const col = document.createElement("div");
        col.className = "column";
        const h = document.createElement("div");
        h.className = "col-title";
        h.innerText = colName;
        col.appendChild(h);

        const list = document.createElement("div");
        list.className = "card-list";
        list.setAttribute("data-col", colName);

        for (const item of state[colName]) {{
          const el = createCardElement(item);
          // store raw data on element
          el.dataset.item = JSON.stringify(item);
          list.appendChild(el);
        }}

        col.appendChild(list);
        board.appendChild(col);

        // make sortable
        new Sortable(list, {{
          group: 'kanban',
          animation: 150,
          onAdd: function (evt) {{
            // nothing special now
          }}
        }});
      }}
    }}

    function getBoardState() {{
      const state = {{}};
      const columns = document.querySelectorAll(".card-list");
      columns.forEach(col => {{
        const name = col.getAttribute("data-col");
        state[name] = [];
        const children = Array.from(col.children);
        children.forEach(ch => {{
          // find label and title
          const labelEl = ch.querySelector(".label");
          const titleEl = ch.querySelector(".title");
          const label = labelEl ? labelEl.innerText : "Feature";
          const text = titleEl ? titleEl.innerText : "";
          state[name].push({{text: text, label: label}});
        }});
      }});
      return state;
    }}

    document.getElementById("exportBtn").addEventListener("click", function() {{
      const data = getBoardState();
      const jsonStr = JSON.stringify(data, null, 2);
      const blob = new Blob([jsonStr], {{type: "application/json"}});
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = "kanban_export.json";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    }});

    document.getElementById("copyBtn").addEventListener("click", function() {{
      const data = getBoardState();
      navigator.clipboard.writeText(JSON.stringify(data, null, 2)).then(function() {{
        alert("JSON copied to clipboard. You can paste it into Upload control on left to save server-side.");
      }});
    }});

    // initial render
    renderBoard(initial);

  </script>
</body>
</html>
"""

        # embed the HTML
        st.components.v1.html(html, height=600, scrolling=True)
