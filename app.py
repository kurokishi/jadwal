# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import json

st.set_page_config(page_title="Jadwal Poli (Streamlit Full)", layout="wide")

# ---------------------------
# Constants for Excel-like view (UPDATED FOR FULL RANGE)
# ---------------------------
TIME_SLOTS = [
    "07:30", "08:00", "08:30", "09:00", "09:30", "10:00", "10:30", 
    "11:00", "11:30", "12:00", "12:30", "13:00", "13:30", "14:00", "14:30"
]

DAYS_ORDER = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]

# ---------------------------
# Navigation State
# ---------------------------
if "current_page" not in st.session_state:
    st.session_state.current_page = "Dashboard"

# ---------------------------
# Helper Functions
# ---------------------------
def _normalize_time_token(token: str) -> str:
    if token is None:
        return ""
    t = str(token).strip()
    if t == "" or t.lower() in ["nan", "none"]:
        return ""
    t = t.replace(".", ":").replace("–", "-").replace("—", "-")
    t = t.lower().replace("am", "").replace("pm", "").strip()
    if ":" not in t:
        if t.isdigit():
            return t.zfill(2) + ":00"
        else:
            return ""
    parts = t.split(":")
    if len(parts) == 2:
        hh = parts[0].zfill(2)
        mm = parts[1].zfill(2)
        try:
            mm_i = int(mm)
            if mm_i < 0 or mm_i > 59:
                return ""
        except:
            return ""
        return f"{hh}:{mm}"
    return ""

def expand_range_safe(range_str: str, interval_minutes: int = 30):
    if not isinstance(range_str, str) or range_str.strip() == "":
        return []
    text = range_str.replace(" ", "")
    parts = None
    for sep in ["-", "–", "—", "to"]:
        if sep in text:
            parts = text.split(sep)
            break
    if parts is None:
        tok = _normalize_time_token(text)
        return [tok] if tok else []
    if len(parts) < 2:
        return []
    start_tok = _normalize_time_token(parts[0])
    end_tok = _normalize_time_token(parts[1])
    if not start_tok or not end_tok:
        return []
    fmt = "%H:%M"
    try:
        sdt = datetime.strptime(start_tok, fmt)
        edt = datetime.strptime(end_tok, fmt)
    except:
        return []
    if edt < sdt:
        return []
    slots = []
    cur = sdt
    while cur <= edt:
        slots.append(cur.strftime("%H:%M"))
        cur += timedelta(minutes=interval_minutes)
    return slots

def create_excel_like_view(df_input: pd.DataFrame) -> pd.DataFrame:
    """Create Excel-like view with POLI, JENIS, HARI, DOKTER as rows and TIME_SLOTS as columns"""
    
    # Create a copy and ensure time slots are standardized
    df = df_input.copy()
    
    # Standardize time slots to match TIME_SLOTS format
    def standardize_time(time_str):
        try:
            t = datetime.strptime(str(time_str).strip(), "%H:%M")
            # Round to nearest 30 minutes
            minutes = t.minute
            if minutes < 15:
                minutes = 0
            elif minutes < 45:
                minutes = 30
            else:
                t += timedelta(hours=1)
                minutes = 0
            return t.replace(minute=minutes).strftime("%H:%M")
        except:
            return str(time_str).strip()
    
    df["Jam"] = df["Jam"].apply(standardize_time)
    
    # Filter only valid time slots
    df = df[df["Jam"].isin(TIME_SLOTS)]
    
    # Create a pivot-like structure
    records = []
    
    # Group by POLI, JENIS POLI, HARI, DOKTER
    grouped = df.groupby(["Poli", "Jenis", "Hari", "Dokter"])
    
    for (poli, jenis, hari, dokter), group in grouped:
        # Create a base record
        record = {
            "POLI ASAL": poli,
            "JENIS POLI": jenis,
            "HARI": hari,
            "DOKTER": dokter
        }
        
        # Initialize all time slots as empty
        for slot in TIME_SLOTS:
            record[slot] = ""
        
        # Fill in the time slots
        for _, row in group.iterrows():
            time_slot = row["Jam"]
            kode = row["Kode"]
            record[time_slot] = kode
        
        records.append(record)
    
    # Convert to DataFrame
    result_df = pd.DataFrame(records)
    
    # Order columns: POLI, JENIS, HARI, DOKTER, then time slots
    ordered_cols = ["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + TIME_SLOTS
    result_df = result_df[ordered_cols]
    
    # Sort by POLI, JENIS, HARI, DOKTER
    result_df = result_df.sort_values(["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"])
    
    # Reset index
    result_df = result_df.reset_index(drop=True)
    
    return result_df

def compute_status(df_in):
    d = df_in.copy()
    d["Over_Kuota"] = False
    d["Bentrok"] = False
    
    # over: count Eksekutif entries per (Hari,Jam)
    eksek = d[d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)]
    poleks_counts = eksek.groupby(["Hari","Jam"]).size()
    
    # Threshold for Over_Kuota (Example: > 7)
    OVER_KUOTA_THRESHOLD = 7
    
    over_slots = poleks_counts[poleks_counts > OVER_KUOTA_THRESHOLD].index if not poleks_counts.empty else []
    
    for (hari,jam) in over_slots:
        d.loc[(d["Hari"]==hari)&(d["Jam"]==jam)&(d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)), "Over_Kuota"] = True
    
    return d

def push_history(df_snapshot):
    if "history" not in st.session_state:
        st.session_state.history = []
    if "future" not in st.session_state:
        st.session_state.future = []
    
    st.session_state.history.append(df_snapshot.copy())
    st.session_state.future.clear()

def undo():
    if "history" in st.session_state and st.session_state.history:
        last = st.session_state.history.pop()
        if "future" not in st.session_state:
            st.session_state.future = []
        st.session_state.future.append(last)
        return last
    return None

def redo():
    if "future" in st.session_state and st.session_state.future:
        f = st.session_state.future.pop()
        if "history" not in st.session_state:
            st.session_state.history = []
        st.session_state.history.append(f)
        return f
    return None

def handle_drag_drop(drag_data_str):
    """Handle drag and drop events"""
    try:
        drag_data = json.loads(drag_data_str)
        st.session_state.last_drag_event = drag_data
        
        source_day = drag_data.get("source_day")
        source_slot = drag_data.get("source_slot")
        target_day = drag_data.get("target_day")
        target_slot = drag_data.get("target_slot")
        card_data = drag_data.get("card_data")
        
        if source_day and source_slot and target_day and target_slot and card_data:
            # 1. Update kanban state
            if "kanban_state" not in st.session_state:
                st.session_state.kanban_state = {}
            
            # Remove from source
            if source_day in st.session_state.kanban_state and source_slot in st.session_state.kanban_state[source_day]:
                st.session_state.kanban_state[source_day][source_slot] = [
                    c for c in st.session_state.kanban_state[source_day][source_slot]
                    if c.get("id") != card_data.get("id")
                ]
            
            # Add to target
            if target_day not in st.session_state.kanban_state:
                st.session_state.kanban_state[target_day] = {}
            if target_slot not in st.session_state.kanban_state[target_day]:
                st.session_state.kanban_state[target_day][target_slot] = []
            
            # Update card ID with new slot
            card_data["id"] = f"{target_day}|{target_slot}|{np.random.randint(1e9)}"
            card_data["Over"] = False
            card_data["Bentrok"] = False
            
            # Append the card
            st.session_state.kanban_state[target_day][target_slot].append(card_data)
            
            # 2. Rebuild main DataFrame and re-compute status
            new_rows_current_day = []
            for s in st.session_state.kanban_state[target_day].keys():
                for c in st.session_state.kanban_state[target_day][s]:
                    new_rows_current_day.append({
                        "Hari": target_day,
                        "Jam": s,
                        "Poli": c.get("Poli", ""),
                        "Jenis": c.get("Jenis", ""),
                        "Dokter": c.get("Dokter", "")
                    })
            
            df_kanban = pd.DataFrame(new_rows_current_day)
            df_kanban["Kode"] = df_kanban["Jenis"].apply(lambda x: "R" if str(x).lower()=="reguler" else "E")
            
            df_other_days_list = []
            if "history" in st.session_state and st.session_state.history:
                latest_df = st.session_state.history[-1]
                df_other_days_list.append(latest_df[latest_df["Hari"] != target_day].copy())
            
            df_new = pd.concat(df_other_days_list + [df_kanban], ignore_index=True)
            df_new = compute_status(df_new)
            
            lanes_new = {}
            for s in TIME_SLOTS:
                rows_s = df_new[(df_new["Hari"]==target_day)&(df_new["Jam"]==s)].reset_index(drop=True)
                cards = []
                rows_s = rows_s.sort_values(by=["Jenis"], key=lambda x: x.str.lower().str.contains("eksekutif|poleks", na=False))
                for i,r in rows_s.iterrows():
                    cards.append({
                        "id": f"{target_day}|{s}|{i}|{np.random.randint(1e9)}",
                        "Dokter": r["Dokter"],
                        "Poli": r["Poli"],
                        "Jenis": r["Jenis"],
                        "Kode": r["Kode"],
                        "Over": bool(r["Over_Kuota"]),
                        "Bentrok": False,
                        "Hari": target_day,
                        "Jam": s
                    })
                if cards:
                    lanes_new[s]=cards
            
            st.session_state.kanban_state[target_day] = lanes_new
            
            # 3. Update Excel view and history
            excel_view_df = create_excel_like_view(df_new)
            st.session_state.excel_view_df = excel_view_df
            push_history(df_new.copy())
            
            return True
    except Exception as e:
        st.error(f"Error handling drag drop: {e}")
    return False

# ---------------------------
# Page Functions
# ---------------------------
def show_dashboard():
    st.title("📊 Dashboard Utama")
    
    if "df" not in st.session_state or st.session_state.df.empty:
        st.info("Silakan upload file data terlebih dahulu di halaman Upload Data")
        return
    
    df = st.session_state.df
    
    # Summary metrics
    st.subheader("📈 Ringkasan Statistik")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Slot", len(df))
    c2.metric("Total Dokter Unik", df["Dokter"].nunique())
    c3.metric("Total Poli", df["Poli"].nunique())
    c4.metric("Slot Waktu Unik", df[["Hari","Jam"]].drop_duplicates().shape[0])
    
    # Poli summary
    st.subheader("📋 Distribusi per Poli")
    try:
        poli_summary = df.groupby(["Poli", "Jenis"]).size().reset_index(name="Jumlah Slot")
        if not poli_summary.empty:
            poli_pivot = poli_summary.pivot(index="Poli", columns="Jenis", values="Jumlah Slot").fillna(0)
            
            cols = st.columns(min(4, len(poli_pivot.index.unique())))
            for idx, (poli, row) in enumerate(poli_pivot.iterrows()):
                with cols[idx % len(cols)]:
                    total = row.sum()
                    reguler = row.get("Reguler", 0)
                    eksekutif = row.get("Eksekutif", 0)
                    
                    st.metric(
                        label=f"🩺 {poli}",
                        value=int(total),
                        delta=f"R:{int(reguler)} E:{int(eksekutif)}"
                    )
    except Exception as e:
        st.warning(f"Tidak dapat menampilkan statistik per Poli: {e}")
    
    # Heatmap
    st.subheader("🌡️ Heatmap (Jam × Hari)")
    try:
        summary = df.groupby(["Hari","Jam"]).size().reset_index(name="Jumlah")
        
        if not summary.empty:
            if "Hari" in summary.columns:
                summary["Hari"] = pd.Categorical(summary["Hari"], categories=DAYS_ORDER, ordered=True)
            if "Jam" in summary.columns:
                summary["Jam"] = pd.Categorical(summary["Jam"], categories=TIME_SLOTS, ordered=True)
            
            summary = summary.sort_values(["Hari", "Jam"])
            
            pivot = summary.pivot_table(index="Jam", columns="Hari", values="Jumlah", aggfunc='first').fillna(0)
            pivot = pivot.reindex(index=TIME_SLOTS, columns=DAYS_ORDER, fill_value=0)
            
            import plotly.express as px
            fig = px.imshow(pivot, 
                            labels=dict(x="Hari", y="Jam", color="Jumlah Dokter"), 
                            color_continuous_scale="Blues",
                            aspect="auto")
            fig.update_xaxes(side="top")
            st.plotly_chart(fig, use_container_width=True)
    except Exception as e:
        st.warning(f"Tidak dapat menampilkan heatmap: {e}")

def show_excel_view():
    st.title("📊 Tampilan Excel-like")
    
    if "excel_view_df" not in st.session_state or st.session_state.excel_view_df.empty:
        st.info("Silakan upload file data terlebih dahulu di halaman Upload Data")
        return
    
    # Filter options for Excel view
    excel_filter_cols = st.columns([2, 2, 2, 1])
    with excel_filter_cols[0]:
        excel_poli_filter = st.multiselect(
            "Filter Poli", 
            sorted(st.session_state.excel_view_df["POLI ASAL"].unique()),
            default=sorted(st.session_state.excel_view_df["POLI ASAL"].unique())
        )
    with excel_filter_cols[1]:
        excel_jenis_filter = st.multiselect(
            "Filter Jenis Poli", 
            sorted(st.session_state.excel_view_df["JENIS POLI"].unique()),
            default=sorted(st.session_state.excel_view_df["JENIS POLI"].unique())
        )
    with excel_filter_cols[2]:
        excel_hari_filter = st.multiselect(
            "Filter Hari", 
            sorted(st.session_state.excel_view_df["HARI"].unique()),
            default=sorted(st.session_state.excel_view_df["HARI"].unique())
        )
    
    # Apply filters
    filtered_excel_df = st.session_state.excel_view_df.copy()
    if excel_poli_filter:
        filtered_excel_df = filtered_excel_df[filtered_excel_df["POLI ASAL"].isin(excel_poli_filter)]
    if excel_jenis_filter:
        filtered_excel_df = filtered_excel_df[filtered_excel_df["JENIS POLI"].isin(excel_jenis_filter)]
    if excel_hari_filter:
        filtered_excel_df = filtered_excel_df[filtered_excel_df["HARI"].isin(excel_hari_filter)]
    
    # Style function for Excel-like view
    def style_excel_view(df_to_style: pd.DataFrame) -> pd.DataFrame:
        def highlight_cell(val):
            if val == "R":
                return 'background-color: #90EE90; color: black; font-weight: bold; text-align: center;'
            elif val == "E":
                return 'background-color: #87CEEB; color: black; font-weight: bold; text-align: center;'
            elif val == "":
                return 'background-color: #F5F5F5; color: #999; text-align: center;'
            else:
                return 'text-align: center;'
        
        styled_df = df_to_style.style.apply(lambda x: x.map(highlight_cell) if x.name in TIME_SLOTS else [''] * len(x))
        
        styled_df = styled_df.set_properties(**{
            'border': '1px solid #ddd',
            'font-size': '12px',
            'font-family': 'Arial, sans-serif'
        })
        
        styled_df = styled_df.set_table_styles([
            {'selector': 'thead th',
             'props': [('background-color', '#4CAF50'), 
                       ('color', 'white'),
                       ('font-weight', 'bold'),
                       ('text-align', 'center')]},
            {'selector': 'th',
             'props': [('background-color', '#f2f2f2'),
                       ('font-weight', 'bold'),
                       ('border', '1px solid #ddd')]}
        ])
        
        return styled_df
    
    # Display the Excel-like view
    st.write(f"**Total baris: {len(filtered_excel_df)}**")
    
    with st.container():
        excel_html = style_excel_view(filtered_excel_df).to_html()
        
        scrollable_html = f"""
        <div style="width: 100%; overflow-x: auto; border: 1px solid #ddd; border-radius: 5px; max-height: 700px; overflow-y: auto;">
            {excel_html}
        </div>
        """
        
        st.markdown(scrollable_html, unsafe_allow_html=True)
    
    # Summary statistics
    st.subheader("📈 Ringkasan Tampilan Excel")
    summary_cols = st.columns(4)
    with summary_cols[0]:
        st.metric("Total Poli", filtered_excel_df["POLI ASAL"].nunique())
    with summary_cols[1]:
        st.metric("Total Dokter", filtered_excel_df["DOKTER"].nunique())
    with summary_cols[2]:
        reguler_count = (filtered_excel_df[TIME_SLOTS] == "R").sum().sum()
        st.metric("Slot Reguler (R)", reguler_count)
    with summary_cols[3]:
        eksekutif_count = (filtered_excel_df[TIME_SLOTS] == "E").sum().sum()
        st.metric("Slot Eksekutif (E)", eksekutif_count)

def show_kanban_editor():
    st.title("🎯 Kanban Editor")
    
    if "df" not in st.session_state or st.session_state.df.empty:
        st.info("Silakan upload file data terlebih dahulu di halaman Upload Data")
        return
    
    df = st.session_state.df
    
    # Sidebar controls for Kanban
    st.sidebar.header("⚙️ Kontrol Kanban")
    
    # Populate selected day with all available days
    all_available_days = sorted(df["Hari"].unique())
    selected_day = st.sidebar.selectbox("Pilih Hari", all_available_days)
    
    # undo/redo buttons
    st.sidebar.subheader("🔄 History")
    colu1, colu2 = st.sidebar.columns(2)
    with colu1:
        if st.button("Undo", use_container_width=True):
            prev = undo()
            if prev is not None:
                st.session_state.df = prev.copy()
                st.session_state.df = compute_status(st.session_state.df)
                st.session_state.excel_view_df = create_excel_like_view(st.session_state.df)
                st.session_state.kanban_state = {}
                st.success("Undo berhasil")
                st.rerun()
    with colu2:
        if st.button("Redo", use_container_width=True):
            nxt = redo()
            if nxt is not None:
                st.session_state.df = nxt.copy()
                st.session_state.df = compute_status(st.session_state.df)
                st.session_state.excel_view_df = create_excel_like_view(st.session_state.df)
                st.session_state.kanban_state = {}
                st.success("Redo berhasil")
                st.rerun()
    
    # Filter the main DataFrame for the selected day
    df_day = df[df["Hari"]==selected_day]
    
    # Initialize kanban state for the selected day if not exists or if data changes
    if "kanban_state" not in st.session_state:
        st.session_state.kanban_state = {}
    
    if selected_day not in st.session_state.kanban_state or st.session_state.kanban_state == {}:
        lanes = {}
        for s in TIME_SLOTS:
            rows_s = df_day[(df_day["Jam"]==s)].reset_index(drop=True)
            cards = []
            
            rows_s["Sort_Order"] = rows_s["Jenis"].apply(lambda x: 0 if str(x).lower()=="reguler" else 1)
            rows_s = rows_s.sort_values(by="Sort_Order")
            
            for i,r in rows_s.iterrows():
                cards.append({
                    "id": f"{selected_day}|{s}|{i}|{np.random.randint(1e9)}",
                    "Dokter": r["Dokter"],
                    "Poli": r["Poli"],
                    "Jenis": r["Jenis"],
                    "Kode": r["Kode"],
                    "Over": bool(r["Over_Kuota"]),
                    "Bentrok": bool(r["Bentrok"]),
                    "Hari": selected_day,
                    "Jam": s
                })
            if cards:
                 lanes[s]=cards
        st.session_state.kanban_state[selected_day] = lanes
    
    # Fill in empty slots explicitly in the session state for rendering all columns
    lanes = st.session_state.kanban_state.get(selected_day, {})
    
    # JavaScript for drag and drop
    drag_drop_js = """
    <script>
    function setupDragAndDrop() {
        const cards = document.querySelectorAll('.kanban-card');
        const columns = document.querySelectorAll('.kanban-column');
        
        cards.forEach(card => {
            card.setAttribute('draggable', 'true');
            
            card.addEventListener('dragstart', (e) => {
                const cardData = JSON.parse(card.getAttribute('data-card'));
                e.dataTransfer.setData('text/plain', JSON.stringify({
                    source_day: cardData.hari,
                    source_slot: cardData.jam,
                    card_data: cardData
                }));
                card.classList.add('dragging');
            });
            
            card.addEventListener('dragend', () => {
                card.classList.remove('dragging');
            });
        });
        
        columns.forEach(column => {
            column.addEventListener('dragover', (e) => {
                e.preventDefault();
                column.classList.add('drag-over');
            });
            
            column.addEventListener('dragleave', () => {
                column.classList.remove('drag-over');
            });
            
            column.addEventListener('drop', (e) => {
                e.preventDefault();
                column.classList.remove('drag-over');
                
                try {
                    const dragData = JSON.parse(e.dataTransfer.getData('text/plain'));
                    const columnData = JSON.parse(column.getAttribute('data-column'));
                    
                    dragData.target_day = columnData.hari;
                    dragData.target_slot = columnData.slot;
                    
                    if (window.parent.document.getElementById('streamlit-kanban-bridge')) {
                        window.parent.document.getElementById('streamlit-kanban-bridge').value = JSON.stringify(dragData);
                        window.parent.document.getElementById('streamlit-kanban-bridge').dispatchEvent(new Event('change'));
                    }
                } catch (error) {
                    console.error('Drop error:', error);
                }
            });
        });
    }
    
    setupDragAndDrop();
    </script>
    
    <style>
    .kanban-container {
        display: flex;
        gap: 5px;
        overflow-x: auto;
        padding: 5px;
        margin-bottom: 20px;
        min-height: 450px;
        max-height: 450px;
    }
    
    .kanban-column {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 5px;
        min-width: 120px;
        max-width: 120px;
        flex-shrink: 0;
        border: 1px solid #dee2e6;
        transition: all 0.2s;
        min-height: 440px;
        max-height: 440px;
        overflow-y: auto;
    }
    
    .kanban-column.drag-over {
        background: #e3f2fd;
        border-color: #2196f3;
    }
    
    .kanban-card {
        background: white;
        border-radius: 4px;
        padding: 6px;
        margin-bottom: 4px;
        box-shadow: 0 1px 2px rgba(0,0,0,0.08);
        cursor: grab;
        font-size: 10px;
        transition: all 0.2s;
        border-left: 3px solid;
        position: relative;
    }
    
    .kanban-card:hover {
        transform: translateY(-1px);
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    
    .kanban-card.dragging {
        opacity: 0.5;
        transform: rotate(2deg);
    }
    
    .kanban-card-reguler {
        border-left-color: #28a745;
    }
    
    .kanban-card-eksekutif {
        border-left-color: #007bff;
    }
    
    .kanban-card-over {
        border-left-color: #dc3545;
    }
    
    .kanban-card-bentrok {
        border-left-color: #ffc107; 
    }
    
    .card-header {
        font-weight: bold;
        font-size: 9px;
        margin-bottom: 3px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
    }
    
    .card-details {
        font-size: 8px;
        color: #666;
    }
    
    .card-icon {
        position: absolute;
        top: 3px;
        right: 3px;
        font-size: 8px;
    }
    
    .slot-header {
        background: #6c757d;
        color: white;
        padding: 4px;
        border-radius: 4px;
        text-align: center;
        margin-bottom: 5px;
        font-size: 11px;
        font-weight: bold;
        position: sticky;
        top: 0;
        z-index: 10;
    }
    
    .empty-slot {
        color: #999;
        font-style: italic;
        text-align: center;
        padding: 10px;
        font-size: 10px;
    }
    </style>
    """
    
    # Create drag and drop interface
    st.markdown("### 🎯 Drag & Drop untuk memindahkan jadwal")
    st.markdown(f"**Hari yang sedang diedit:** {selected_day}")
    st.markdown("**Instruksi:** Seret kartu dokter dari satu slot waktu ke slot lainnya")
    
    # Hidden input to receive drag data from custom JS
    drag_data_receiver = st.empty()
    drag_data_str = drag_data_receiver.text_input("Drag Drop Receiver (Hidden)", key="streamlit-kanban-bridge", label_visibility="collapsed")
    
    if drag_data_str and drag_data_str != st.session_state.get('last_processed_drag_data'):
        st.session_state['last_processed_drag_data'] = drag_data_str
        if handle_drag_drop(drag_data_str):
            st.success("Perpindahan jadwal berhasil! Status Over Kuota diperbarui.")
            st.session_state["streamlit-kanban-bridge"] = ""
            st.rerun()
    
    # Create the kanban board HTML
    kanban_html = "<div class='kanban-container'>"
    
    for slot in TIME_SLOTS:
        cards = lanes.get(slot, [])
        
        column_data = json.dumps({
            "hari": selected_day,
            "slot": slot
        })
        
        kanban_html += f"""
        <div class='kanban-column' data-column='{column_data}'>
            <div class='slot-header'>{slot}</div>
        """
        
        if cards:
            for card in cards:
                card_class = "kanban-card"
                if card.get("Over"):
                    card_class += " kanban-card-over"
                elif card.get("Bentrok"):
                    card_class += " kanban-card-bentrok"
                elif card.get("Kode") == "R":
                    card_class += " kanban-card-reguler"
                else:
                    card_class += " kanban-card-eksekutif"
                
                status_icon = ""
                if card.get("Over"):
                    status_icon = "🔴"
                elif card.get("Bentrok"):
                    status_icon = "🟡"
                elif card.get("Kode") == "R":
                    status_icon = "🟢"
                else:
                    status_icon = "🔵"
                
                card_data = json.dumps(card)
                
                kanban_html += f"""
                <div class='{card_class}' draggable='true' data-card='{card_data}'>
                    <div class='card-header' title='{card['Dokter']}'>{card['Dokter']}</div>
                    <div class='card-details'>
                        <div><strong>Poli:</strong> {card['Poli']}</div>
                        <div><strong>Tipe:</strong> {card['Jenis']}</div>
                    </div>
                    <div class='card-icon'>{status_icon}</div>
                </div>
                """
        else:
            kanban_html += "<div class='empty-slot'>Kosong</div>"
        
        kanban_html += "</div>"
    
    kanban_html += "</div>"
    
    st.components.v1.html(drag_drop_js + kanban_html, height=500, scrolling=False)
    
    # Manual move section
    with st.expander("📝 Pindah Manual (Fallback)", expanded=False):
        all_cards = []
        for s in TIME_SLOTS:
            for card in lanes.get(s, []):
                card_copy = card.copy()
                card_copy["Jam"] = s
                all_cards.append(card_copy)
        
        if all_cards:
            col_select, col_move = st.columns([3, 2])
            
            with col_select:
                options = []
                display_texts = []
                for i, card in enumerate(all_cards):
                    display_text = f"{card['Dokter']} - {card['Poli']} ({card['Jenis']}) @ {card['Jam']}"
                    if card.get("Over"):
                        display_text += " 🔴 OVER"
                    options.append(i)
                    display_texts.append(display_text)
                
                selected_index = st.selectbox(
                    "Pilih jadwal dokter:",
                    options=options,
                    format_func=lambda x: display_texts[x],
                    key="manual_move_select"
                )
                
                selected_card = all_cards[selected_index]
                st.info(f"**Terpilih:** {selected_card['Dokter']} - {selected_card['Poli']} @ {selected_card['Jam']}")
            
            with col_move:
                target_slot = st.selectbox(
                    "Pindahkan ke jam:",
                    TIME_SLOTS,
                    index=TIME_SLOTS.index(selected_card["Jam"]) if selected_card["Jam"] in TIME_SLOTS else 0,
                    key="manual_move_target"
                )
                
                if st.button("🚀 Pindahkan Jadwal", type="primary", use_container_width=True):
                    card_to_move = all_cards[selected_index]
                    original_slot = card_to_move["Jam"]
                    
                    st.session_state.kanban_state[selected_day][original_slot] = [
                        c for c in st.session_state.kanban_state[selected_day][original_slot]
                        if c.get("id") != card_to_move.get("id")
                    ]
                    
                    new_card = {
                        "id": f"{selected_day}|{target_slot}|{np.random.randint(1e9)}",
                        "Dokter": card_to_move["Dokter"],
                        "Poli": card_to_move["Poli"],
                        "Jenis": card_to_move["Jenis"],
                        "Kode": card_to_move.get("Kode", "E"),
                        "Over": False,
                        "Bentrok": False,
                        "Hari": selected_day,
                        "Jam": target_slot
                    }
                    
                    if target_slot not in st.session_state.kanban_state[selected_day]:
                        st.session_state.kanban_state[selected_day][target_slot] = []
                    
                    st.session_state.kanban_state[selected_day][target_slot].append(new_card)
                    
                    new_rows_current_day = []
                    for s in TIME_SLOTS:
                        for c in st.session_state.kanban_state[selected_day].get(s, []):
                            new_rows_current_day.append({
                                "Hari": selected_day,
                                "Jam": s,
                                "Poli": c.get("Poli", ""),
                                "Jenis": c.get("Jenis", ""),
                                "Dokter": c.get("Dokter", "")
                            })

                    df_kanban = pd.DataFrame(new_rows_current_day)
                    df_kanban["Kode"] = df_kanban["Jenis"].apply(lambda x: "R" if str(x).lower()=="reguler" else "E")
                    
                    df_other = st.session_state.df[st.session_state.df["Hari"] != selected_day].copy()
                    df_new = pd.concat([df_other, df_kanban], ignore_index=True)
                    df_new = compute_status(df_new)
                    
                    st.session_state.df = df_new
                    st.session_state.excel_view_df = create_excel_like_view(df_new)
                    
                    lanes_new = {}
                    for s in TIME_SLOTS:
                        rows_s = df_new[(df_new["Hari"]==selected_day)&(df_new["Jam"]==s)].reset_index(drop=True)
                        cards_new = []
                        rows_s["Sort_Order"] = rows_s["Jenis"].apply(lambda x: 0 if str(x).lower()=="reguler" else 1)
                        rows_s = rows_s.sort_values(by="Sort_Order")

                        for i,r in rows_s.iterrows():
                            cards_new.append({
                                "id": f"{selected_day}|{s}|{i}|{np.random.randint(1e9)}",
                                "Dokter": r["Dokter"],
                                "Poli": r["Poli"],
                                "Jenis": r["Jenis"],
                                "Kode": r["Kode"],
                                "Over": bool(r["Over_Kuota"]),
                                "Bentrok": False, 
                                "Hari": selected_day,
                                "Jam": s
                            })
                        if cards_new:
                             lanes_new[s]=cards_new

                    st.session_state.kanban_state[selected_day] = lanes_new
                    push_history(df_new.copy())
                    
                    st.success(f"✅ {card_to_move['Dokter']} dipindahkan dari {original_slot} ke {target_slot}")
                    st.rerun()

def show_upload_data():
    st.title("📤 Upload Data")
    
    st.markdown("### Upload file Excel/CSV dengan format:")
    st.code("""
    Kolom yang dibutuhkan:
    - Hari (Senin, Selasa, Rabu, Kamis, Jum'at)
    - Range (contoh: 07.30-09.00, 08:00-10:30)
    - Poli (nama poli)
    - Jenis (Reguler/Eksekutif)
    - Dokter (nama dokter)
    """)
    
    # Upload section
    uploaded = st.file_uploader(
        "Upload file Excel atau CSV", 
        type=["xlsx", "csv"],
        key="file_uploader"
    )
    
    if st.button("Download Template Example"):
        sample = pd.DataFrame({
            "Hari": ["Senin", "Senin", "Selasa"],
            "Range": ["07.30-09.00", "09.00-11.00", "07.00-08.30"],
            "Poli": ["Anak", "Anak", "Gigi"],
            "Jenis": ["Reguler", "Eksekutif", "Reguler"],
            "Dokter": ["dr. Budi", "dr. Sari", "drg. Putri"]
        })
        
        example_range = f"{TIME_SLOTS[0]}-{TIME_SLOTS[-1]}"
        sample_full_range = pd.DataFrame({
            "Hari": ["Rabu", "Rabu"],
            "Range": [example_range, example_range],
            "Poli": ["Penyakit Dalam", "Jantung"],
            "Jenis": ["Reguler", "Eksekutif"],
            "Dokter": ["dr. Dedi", "dr. Siti"]
        })
        
        sample = pd.concat([sample, sample_full_range], ignore_index=True)
        
        st.download_button(
            "Download template.xlsx", 
            data=sample.to_excel(index=False, engine="openpyxl"), 
            file_name="template_jadwal.xlsx", 
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    if uploaded:
        @st.cache_data
        def load_raw(bytes_io, fname):
            try:
                if fname.lower().endswith(".csv"):
                    return pd.read_csv(io.BytesIO(bytes_io))
                else:
                    xls = pd.ExcelFile(io.BytesIO(bytes_io))
                    sheet = "Jadwal" if "Jadwal" in xls.sheet_names else xls.sheet_names[0]
                    df = xls.parse(sheet)
                    df.columns = [str(c).strip() for c in df.columns]
                    return df
            except Exception as e:
                st.error(f"Gagal membaca file: {e}")
                return pd.DataFrame()

        raw = load_raw(uploaded.getvalue(), uploaded.name)
        
        # Column mapping
        col_map = {
            "Hari": ["Hari","Day","hari","day"],
            "Range": ["Range","Jam","Waktu","Time","range","jam","waktu","time"],
            "Poli": ["Poli","Poliklinik","Unit","poli","poliklinik","unit"],
            "Jenis": ["Jenis","Type","Kategori","jenis","type","kategori"],
            "Dokter": ["Dokter","Nama Dokter","Doctor","dokter","nama dokter","doctor"]
        }
        
        def find_col(cols, candidates):
            for c in candidates:
                if c in cols:
                    return c
            return None

        cols = raw.columns.tolist()
        Hari_col = find_col(cols, col_map["Hari"])
        Range_col = find_col(cols, col_map["Range"])
        Poli_col = find_col(cols, col_map["Poli"])
        Jenis_col = find_col(cols, col_map["Jenis"])
        Dokter_col = find_col(cols, col_map["Dokter"])
        
        if not (Hari_col and Range_col and Poli_col and Jenis_col and Dokter_col):
            st.error("Kolom input tidak lengkap. Pastikan file memiliki kolom: Hari, Range, Poli, Jenis, Dokter.")
            st.write("Kolom yang terbaca:", cols)
            return
        
        # Expand ranges -> slots
        expanded = []
        for _, r in raw.iterrows():
            hari = str(r.get(Hari_col)).strip()
            rng = r.get(Range_col)
            poli = str(r.get(Poli_col)).strip()
            jenis = str(r.get(Jenis_col)).strip()
            dokter = str(r.get(Dokter_col)).strip()
            
            if not hari or pd.isna(rng) or not poli or not jenis or not dokter:
                continue
                
            slots = expand_range_safe(str(rng), interval_minutes=30)
            if not slots:
                tok = _normalize_time_token(str(rng))
                if tok:
                    slots = [tok]
                    
            for s in slots:
                expanded.append({
                    "Hari": hari, 
                    "Jam": s, 
                    "Poli": poli, 
                    "Jenis": jenis, 
                    "Dokter": dokter
                })
        
        if len(expanded) == 0:
            st.error("Tidak ada slot terbentuk. Periksa format Range.")
            return
        
        df = pd.DataFrame(expanded)
        
        # Normalize Jenis
        df["Jenis"] = df["Jenis"].astype(str).str.strip().replace({
            "reguler": "Reguler", 
            "regular": "Reguler", 
            "eksekutif": "Eksekutif", 
            "executive": "Eksekutif", 
            "poleks": "Eksekutif", 
            "POLEKS": "Eksekutif"
        })
        
        df["Kode"] = df["Jenis"].apply(lambda x: "R" if str(x).lower()=="reguler" else "E")
        
        # Compute status
        df = compute_status(df)
        
        # Create Excel-like view
        excel_view_df = create_excel_like_view(df)
        
        # Store in session state
        st.session_state.df = df
        st.session_state.excel_view_df = excel_view_df
        st.session_state.history = [df.copy()]
        st.session_state.future = []
        st.session_state.kanban_state = {}
        
        st.success(f"✅ Data berhasil diupload! {len(df)} slot jadwal ditemukan.")
        st.balloons()
        
        # Show preview
        st.subheader("📋 Preview Data")
        st.dataframe(df.head(10), use_container_width=True)
        
        # Auto navigate to dashboard
        st.session_state.current_page = "Dashboard"
        st.rerun()
    else:
        st.info("Silakan upload file Excel/CSV untuk memulai.")

def show_export():
    st.title("💾 Export Data")
    
    if "df" not in st.session_state or st.session_state.df.empty:
        st.info("Silakan upload file data terlebih dahulu di halaman Upload Data")
        return
    
    st.markdown("### Pilih format export:")
    
    export_cols = st.columns(3)
    
    with export_cols[0]:
        csv_bytes = st.session_state.df.to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Download CSV (Detail)",
            csv_bytes,
            file_name="jadwal_detail.csv",
            mime="text/csv",
            use_container_width=True
        )
    
    with export_cols[1]:
        def to_xlsx_bytes(df_in):
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df_in.to_excel(writer, index=False, sheet_name="Jadwal_Detail")
            return out.getvalue()
        
        xlsx_bytes = to_xlsx_bytes(st.session_state.df)
        st.download_button(
            "📥 Download XLSX (Detail)",
            xlsx_bytes,
            file_name="jadwal_detail.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    with export_cols[2]:
        def to_excel_view_xlsx(df_in):
            out = io.BytesIO()
            with pd.ExcelWriter(out, engine="openpyxl") as writer:
                df_in.to_excel(writer, index=False, sheet_name="Jadwal_Excel_View")
            return out.getvalue()
        
        if "excel_view_df" in st.session_state and not st.session_state.excel_view_df.empty:
            excel_view_xlsx = to_excel_view_xlsx(st.session_state.excel_view_df)
            st.download_button(
                "📊 Download Excel View",
                excel_view_xlsx,
                file_name="jadwal_excel_view.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("Data Excel View tidak tersedia")

# ---------------------------
# Main App
# ---------------------------
def main():
    # Sidebar Navigation
    with st.sidebar:
        st.title("📅 Jadwal Poli")
        st.markdown("---")
        
        # Navigation menu
        st.subheader("📋 Menu Navigasi")
        
        pages = {
            "Dashboard": "📊",
            "Tampilan Excel": "📈",
            "Kanban Editor": "🎯",
            "Upload Data": "📤",
            "Export": "💾"
        }
        
        for page_name, icon in pages.items():
            if st.button(f"{icon} {page_name}", 
                        use_container_width=True,
                        type="primary" if st.session_state.current_page == page_name else "secondary"):
                st.session_state.current_page = page_name
                st.rerun()
        
        st.markdown("---")
        
        # App info
        st.caption("""
        **Jadwal Poli Management System**
        
        Versi: 2.0
        Fitur:
        - Upload data Excel/CSV
        - Tampilan Excel-like
        - Kanban editor drag & drop
        - Dashboard statistik
        - Export berbagai format
        """)
    
    # Main content area
    if st.session_state.current_page == "Dashboard":
        show_dashboard()
    elif st.session_state.current_page == "Tampilan Excel":
        show_excel_view()
    elif st.session_state.current_page == "Kanban Editor":
        show_kanban_editor()
    elif st.session_state.current_page == "Upload Data":
        show_upload_data()
    elif st.session_state.current_page == "Export":
        show_export()

if __name__ == "__main__":
    main()
