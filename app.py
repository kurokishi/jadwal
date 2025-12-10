# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import json
import os

st.set_page_config(page_title="Jadwal Poli (Streamlit Full)", layout="wide")
st.title("📅 Jadwal Poli — Streamlit Full")

# ---------------------------
# Constants for Excel-like view
# ---------------------------
TIME_SLOTS = [
    "07:00", "07:30", "08:00", "08:30", "09:00", "09:30", "10:00", "10:30", 
    "11:00", "11:30", "12:00", "12:30", "13:00", "13:30", "14:00", "14:30"
]

DAYS_ORDER = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]

# ---------------------------
# Session state: history (undo/redo), kanban_state
# ---------------------------
if "history" not in st.session_state:
    st.session_state.history = []
if "future" not in st.session_state:
    st.session_state.future = []
if "kanban_state" not in st.session_state:
    st.session_state.kanban_state = {}  # day -> lanes dict
if "excel_view_df" not in st.session_state:
    st.session_state.excel_view_df = pd.DataFrame()

def push_history(df_snapshot):
    st.session_state.history.append(df_snapshot.copy())
    st.session_state.future.clear()  # clear redo stack

def undo():
    if st.session_state.history:
        last = st.session_state.history.pop()
        st.session_state.future.append(last)
        return last
    return None

def redo():
    if st.session_state.future:
        f = st.session_state.future.pop()
        st.session_state.history.append(f)
        return f
    return None

# ---------------------------
# Upload input
# ---------------------------
st.sidebar.header("Upload / Template")
uploaded = st.sidebar.file_uploader("Upload Excel Jadwal", type=["xlsx","xls"])

if uploaded is None:
    st.info("Upload file Excel jadwal untuk mulai.")
    st.stop()

# Function to parse Excel format from your file
def parse_excel_format(uploaded_file):
    """Parse Excel format from the uploaded file"""
    try:
        # Read the Excel file
        xls = pd.ExcelFile(uploaded_file)
        
        # Try to read sheet "Jadwal" or first sheet
        sheet_name = "Jadwal" if "Jadwal" in xls.sheet_names else xls.sheet_names[0]
        df = xls.parse(sheet_name, header=0)
        
        # Display raw data for debugging
        st.sidebar.write(f"📊 Sheet: {sheet_name}")
        st.sidebar.write(f"📋 Kolom: {list(df.columns)}")
        
        # Check the structure of the data
        if len(df) > 0:
            st.sidebar.write("🔍 Data preview (5 rows):")
            st.sidebar.dataframe(df.head(), use_container_width=True)
        
        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]
        
        # Find the correct column names
        # Look for common column patterns in your Excel file
        poli_col = None
        jenis_col = None
        hari_col = None
        dokter_col = None
        
        for col in df.columns:
            col_lower = col.lower()
            if "poli" in col_lower:
                poli_col = col
            elif "jenis" in col_lower or "type" in col_lower:
                jenis_col = col
            elif "hari" in col_lower or "day" in col_lower:
                hari_col = col
            elif "dokter" in col_lower or "doctor" in col_lower:
                dokter_col = col
        
        # If we couldn't find the columns, use the first few columns
        if not hari_col and len(df.columns) > 0:
            hari_col = df.columns[0] if "Hari" in df.columns[0] else "Hari"
        
        if not poli_col and len(df.columns) > 1:
            poli_col = df.columns[1] if "Poli" in df.columns[1] else "Poli"
        
        if not jenis_col and len(df.columns) > 2:
            jenis_col = df.columns[2] if "Jenis" in df.columns[2] else "Jenis"
        
        if not dokter_col and len(df.columns) > 3:
            dokter_col = df.columns[3] if "Dokter" in df.columns[3] else "Dokter"
        
        # Find time slot columns (columns with time patterns)
        time_columns = []
        for col in df.columns:
            if any(time_str in str(col) for time_str in ["07:", "08:", "09:", "10:", "11:", "12:", "13:", "14:", "15:"]):
                time_columns.append(col)
            elif ":" in str(col) and len(str(col)) <= 5:  # Time format like "07:30"
                time_columns.append(col)
        
        # If no time columns found, check for numeric time representations
        if not time_columns:
            # Check if there are columns that might be times
            for col in df.columns:
                try:
                    # Try to parse as time
                    pd.to_datetime(col, format='%H:%M')
                    time_columns.append(col)
                except:
                    continue
        
        # If still no time columns, check columns after the first 4
        if not time_columns and len(df.columns) > 4:
            time_columns = df.columns[4:].tolist()
        
        st.sidebar.write(f"⏰ Kolom waktu ditemukan: {len(time_columns)}")
        
        if not time_columns:
            st.error("Tidak ditemukan kolom waktu dalam file Excel!")
            return None
        
        # Now parse the data
        expanded_data = []
        
        for _, row in df.iterrows():
            hari = str(row.get(hari_col, "")).strip()
            poli = str(row.get(poli_col, "")).strip()
            jenis = str(row.get(jenis_col, "")).strip()
            dokter = str(row.get(dokter_col, "")).strip()
            
            # Skip empty rows
            if not hari or not dokter or pd.isna(hari) or pd.isna(dokter):
                continue
            
            # Check each time column
            for time_col in time_columns:
                value = str(row.get(time_col, "")).strip()
                
                # Check if this cell has a schedule (R or E)
                if value and value.upper() in ["R", "E"]:
                    expanded_data.append({
                        "Hari": hari,
                        "Jam": str(time_col).strip(),
                        "Poli": poli,
                        "Jenis": jenis if jenis else ("Reguler" if value.upper() == "R" else "Eksekutif"),
                        "Dokter": dokter,
                        "Kode": value.upper()
                    })
        
        if not expanded_data:
            # Try alternative parsing - maybe the data is in a different format
            st.warning("Mencoba metode parsing alternatif...")
            
            # Try to melt the dataframe
            if len(df.columns) >= 5:
                # Assume first 4 columns are metadata, rest are time slots
                id_vars = df.columns[:4].tolist()
                value_vars = df.columns[4:].tolist()
                
                melted_df = pd.melt(df, 
                                   id_vars=id_vars,
                                   value_vars=value_vars,
                                   var_name="Jam",
                                   value_name="Kode")
                
                # Filter only rows with R or E
                melted_df = melted_df[melted_df["Kode"].astype(str).str.upper().isin(["R", "E"])]
                
                if not melted_df.empty:
                    # Rename columns
                    column_map = {}
                    for i, col in enumerate(id_vars):
                        if i == 0: column_map[col] = "Hari"
                        elif i == 1: column_map[col] = "Poli"
                        elif i == 2: column_map[col] = "Jenis"
                        elif i == 3: column_map[col] = "Dokter"
                    
                    melted_df = melted_df.rename(columns=column_map)
                    
                    # Convert to list of dicts
                    expanded_data = melted_df.to_dict('records')
        
        if not expanded_data:
            st.error("Tidak ada data jadwal yang ditemukan dalam file!")
            st.write("Struktur file yang diharapkan:")
            st.write("- Kolom 1: Hari (Senin, Selasa, dll)")
            st.write("- Kolom 2: Poli (Poli Anak, Poli Bedah, dll)")
            st.write("- Kolom 3: Jenis (Reguler, Eksekutif/Poleks)")
            st.write("- Kolom 4: Dokter (Nama dokter)")
            st.write("- Kolom 5+: Waktu (07:30, 08:00, dll) dengan nilai R atau E")
            return None
        
        return pd.DataFrame(expanded_data)
        
    except Exception as e:
        st.error(f"Error parsing Excel file: {str(e)}")
        return None

# Parse the uploaded file
df = parse_excel_format(uploaded)

if df is None or df.empty:
    st.error("Tidak dapat memproses file Excel. Pastikan format file sesuai.")
    st.stop()

# Ensure we have the required columns
required_cols = ["Hari", "Jam", "Poli", "Jenis", "Dokter", "Kode"]
for col in required_cols:
    if col not in df.columns:
        st.error(f"Kolom '{col}' tidak ditemukan dalam data!")
        st.stop()

# Clean and standardize the data
# Normalize Jenis based on Kode
df["Jenis"] = df["Kode"].apply(lambda x: "Reguler" if x == "R" else "Eksekutif")

# Clean time format
def clean_time(time_str):
    try:
        # Remove any whitespace
        time_str = str(time_str).strip()
        
        # Handle different time formats
        if ":" in time_str:
            parts = time_str.split(":")
            if len(parts) == 2:
                hours = parts[0].zfill(2)
                minutes = parts[1].zfill(2)
                return f"{hours}:{minutes}"
        
        # Try to parse as datetime
        dt = pd.to_datetime(time_str, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime("%H:%M")
        
        return time_str
    except:
        return time_str

df["Jam"] = df["Jam"].apply(clean_time)

# Filter only valid time slots
df = df[df["Jam"].isin(TIME_SLOTS)]

if df.empty:
    st.error("Tidak ada data dengan slot waktu yang valid (07:00-14:30).")
    st.stop()

# ---------------------------
# NEW: Convert to Excel-like format (Pivot table)
# ---------------------------
def create_excel_like_view(df_input: pd.DataFrame) -> pd.DataFrame:
    """Create Excel-like view with POLI, JENIS, HARI, DOKTER as rows and TIME_SLOTS as columns"""
    
    # Create a copy and ensure time slots are standardized
    df = df_input.copy()
    
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
    # Only include columns that exist
    ordered_cols = [col for col in ordered_cols if col in result_df.columns or col in TIME_SLOTS]
    
    # Reorder and add missing columns
    for col in TIME_SLOTS:
        if col not in result_df.columns:
            result_df[col] = ""
    
    result_df = result_df[["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + TIME_SLOTS]
    
    # Sort by POLI, JENIS, HARI, DOKTER
    result_df = result_df.sort_values(["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"])
    
    # Reset index
    result_df = result_df.reset_index(drop=True)
    
    return result_df

# Create Excel-like view
excel_view_df = create_excel_like_view(df)
st.session_state.excel_view_df = excel_view_df

# push initial snapshot to history
push_history(df.copy())

# ---------------------------
# REMOVED: Compute Over-kuota only for Eksekutif/Poleks
# ---------------------------
def compute_status(df_in):
    d = df_in.copy()
    d["Over_Kuota"] = False
    
    # ONLY check over kuota for Eksekutif/Poleks (maximum 7 per slot)
    eksek = d[d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)]
    poleks_counts = eksek.groupby(["Hari","Jam"]).size()
    over_slots = poleks_counts[poleks_counts > 7].index if not poleks_counts.empty else []
    
    for (hari,jam) in over_slots:
        d.loc[(d["Hari"]==hari)&(d["Jam"]==jam)&(d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)), "Over_Kuota"] = True
    
    # REMOVED: Bentrok checking - dokter boleh ada di multiple poli di slot yang sama
    # This rule has been removed as requested
    
    return d

df = compute_status(df)

# ---------------------------
# Display success message and data summary
# ---------------------------
st.success(f"✅ File berhasil diproses! Data loaded: {len(df)} jadwal")

col1, col2, col3, col4 = st.columns(4)
with col1:
    st.metric("Total Hari", df["Hari"].nunique())
with col2:
    st.metric("Total Poli", df["Poli"].nunique())
with col3:
    st.metric("Total Dokter", df["Dokter"].nunique())
with col4:
    st.metric("Jadwal Reguler", len(df[df["Kode"] == "R"]))

# ---------------------------
# UI: Filters & summary
# ---------------------------
st.sidebar.header("Filter & Actions")
hari_list = sorted(df["Hari"].unique())
selected_day = st.sidebar.selectbox("Pilih Hari (kanban)", ["--Semua--"] + hari_list)
poli_filter = st.sidebar.multiselect("Filter Poli (opsional)", sorted(df["Poli"].unique()), default=list(df["Poli"].unique()))
jenis_filter = st.sidebar.multiselect("Filter Jenis", sorted(df["Jenis"].unique()), default=list(df["Jenis"].unique()))

# undo/redo buttons
colu1, colu2, colu3 = st.sidebar.columns([1,1,2])
with colu1:
    if st.button("Undo"):
        prev = undo()
        if prev is not None:
            df = prev.copy()
            # Recreate Excel view
            excel_view_df = create_excel_like_view(df)
            st.session_state.excel_view_df = excel_view_df
            st.success("Undo berhasil")
with colu2:
    if st.button("Redo"):
        nxt = redo()
        if nxt is not None:
            df = nxt.copy()
            # Recreate Excel view
            excel_view_df = create_excel_like_view(df)
            st.session_state.excel_view_df = excel_view_df
            st.success("Redo berhasil")

# ---------------------------
# NEW: Excel-like View
# ---------------------------
st.header("📊 Tampilan Jadwal (Format Excel)")

# Filter options for Excel view
excel_filter_cols = st.columns([2, 2, 2, 1])
with excel_filter_cols[0]:
    excel_poli_filter = st.multiselect(
        "Filter Poli (Excel View)", 
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

# Apply filters to Excel view
filtered_excel_df = st.session_state.excel_view_df.copy()
if excel_poli_filter:
    filtered_excel_df = filtered_excel_df[filtered_excel_df["POLI ASAL"].isin(excel_poli_filter)]
if excel_jenis_filter:
    filtered_excel_df = filtered_excel_df[filtered_excel_df["JENIS POLI"].isin(excel_jenis_filter)]
if excel_hari_filter:
    filtered_excel_df = filtered_excel_df[filtered_excel_df["HARI"].isin(excel_hari_filter)]

# Style function for Excel-like view
def style_excel_view(df_to_style: pd.DataFrame) -> pd.DataFrame:
    """Apply styling to Excel-like view dataframe"""
    
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
    
    # Add borders and formatting
    styled_df = styled_df.set_properties(**{
        'border': '1px solid #ddd',
        'font-size': '12px',
        'font-family': 'Arial, sans-serif'
    })
    
    # Header styling
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

# Add horizontal scrolling container
with st.container():
    excel_html = style_excel_view(filtered_excel_df).to_html()
    
    # Wrap in div with horizontal scroll
    scrollable_html = f"""
    <div style="width: 100%; overflow-x: auto; border: 1px solid #ddd; border-radius: 5px; max-height: 700px; overflow-y: auto;">
        {excel_html}
    </div>
    """
    
    st.markdown(scrollable_html, unsafe_allow_html=True)

# Summary statistics for Excel view
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

# ---------------------------
# Dashboard & summary
# ---------------------------
st.header("📊 Dashboard Ringkasan")
c1, c2, c3, c4 = st.columns(4)
c1.metric("Total slot", len(df))
c2.metric("Total dokter unik", df["Dokter"].nunique())
c3.metric("Total Poli", df["Poli"].nunique())
c4.metric("Slot unik", df[["Hari","Jam"]].drop_duplicates().shape[0])

# FIXED: Heatmap dengan pivot_table untuk handle duplicates
st.subheader("Heatmap (Jam × Hari)")
try:
    # Gunakan pivot_table dengan aggregation function 'size'
    summary = df.groupby(["Hari","Jam"]).size().reset_index(name="Jumlah")
    
    # Ensure proper ordering
    if not summary.empty:
        # Convert to categorical for proper ordering
        if "Hari" in summary.columns:
            summary["Hari"] = pd.Categorical(summary["Hari"], categories=DAYS_ORDER, ordered=True)
        if "Jam" in summary.columns:
            summary["Jam"] = pd.Categorical(summary["Jam"], categories=TIME_SLOTS, ordered=True)
        
        summary = summary.sort_values(["Hari", "Jam"])
        
        # Gunakan pivot_table dengan aggfunc='first' karena data sudah di-aggregate
        pivot = summary.pivot_table(index="Jam", columns="Hari", values="Jumlah", aggfunc='first').fillna(0)
        
        # Sort index and columns
        pivot = pivot.reindex(index=TIME_SLOTS, columns=DAYS_ORDER, fill_value=0)
        
        import plotly.express as px
        fig = px.imshow(pivot, 
                        labels=dict(x="Hari", y="Jam", color="Jumlah Dokter"), 
                        color_continuous_scale="Blues",
                        aspect="auto")
        fig.update_xaxes(side="top")
        st.plotly_chart(fig, width='stretch', height=400)
    else:
        st.info("Tidak ada data untuk ditampilkan dalam heatmap.")
except Exception as e:
    st.warning(f"Tidak dapat menampilkan heatmap: {e}")

# ---------------------------
# NEW: IMPROVED KANBAN EDITOR - FULL VIEW WITH SEPARATE REGULER & POLEKS
# ---------------------------
st.header("🎯 Kanban Editor - Semua Jam (07:00-14:30)")

if selected_day == "--Semua--":
    st.info("Pilih satu hari di sidebar untuk membuka Kanban editor.")
else:
    # Use all time slots from 07:00 to 14:30
    SLOTS = TIME_SLOTS
    
    # Initialize kanban state if not exists
    if selected_day not in st.session_state.kanban_state:
        lanes = {}
        for s in SLOTS:
            # Get all doctors for this time slot (both Reguler and Eksekutif)
            rows_s = df[(df["Hari"]==selected_day)&(df["Jam"]==s)].reset_index(drop=True)
            cards = []
            for i,r in rows_s.iterrows():
                cards.append({
                    "id": f"{selected_day}|{s}|{i}|{np.random.randint(1e9)}",
                    "Dokter": r["Dokter"],
                    "Poli": r["Poli"],
                    "Jenis": r["Jenis"],
                    "Kode": r["Kode"],
                    "Over": bool(r["Over_Kuota"]),
                    "Hari": selected_day,
                    "Jam": s
                })
            lanes[s]=cards
        st.session_state.kanban_state[selected_day] = lanes
    
    lanes = st.session_state.kanban_state[selected_day]
    
    # Create two separate views: Reguler and Eksekutif
    st.markdown(f"### 📅 Hari: {selected_day}")
    
    # Create tabs for better organization
    tab1, tab2 = st.tabs(["🎯 Tampilan Kanban (Drag & Drop)", "📝 Edit Manual"])
    
    with tab1:
        # JavaScript for drag and drop
        drag_drop_js = """
        <script>
        // Drag and Drop functionality
        function setupDragAndDrop() {
            const cards = document.querySelectorAll('.kanban-card');
            const columns = document.querySelectorAll('.kanban-column');
            
            // Setup draggable cards
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
                    // Refresh page to show changes
                    setTimeout(() => window.location.reload(), 300);
                });
            });
            
            // Setup droppable columns
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
                        
                        // Update drag data with target
                        dragData.target_day = columnData.hari;
                        dragData.target_slot = columnData.slot;
                        
                        // Send to Streamlit via parent window
                        window.parent.postMessage({
                            type: 'KANBAN_DRAG_DROP',
                            data: dragData
                        }, '*');
                        
                    } catch (error) {
                        console.error('Drop error:', error);
                    }
                });
            });
        }
        
        // Initialize when page loads
        if (document.readyState === 'loading') {
            document.addEventListener('DOMContentLoaded', setupDragAndDrop);
        } else {
            setupDragAndDrop();
        }
        </script>
        
        <style>
        .kanban-board {
            display: flex;
            gap: 5px;
            overflow-x: auto;
            padding: 10px;
            margin-bottom: 20px;
            background: #f5f5f5;
            border-radius: 8px;
        }
        
        .kanban-column {
            background: white;
            border-radius: 6px;
            padding: 8px;
            min-width: 130px;
            max-width: 150px;
            border: 1px solid #ddd;
            transition: all 0.2s;
            min-height: 500px;
            display: flex;
            flex-direction: column;
        }
        
        .kanban-column.drag-over {
            background: #e8f4fd;
            border-color: #4dabf7;
            border-width: 2px;
        }
        
        .slot-header {
            background: #495057;
            color: white;
            padding: 6px;
            border-radius: 4px;
            text-align: center;
            margin-bottom: 10px;
            font-size: 11px;
            font-weight: bold;
            position: sticky;
            top: 0;
            z-index: 10;
        }
        
        .card-container {
            flex-grow: 1;
            overflow-y: auto;
            padding-right: 2px;
        }
        
        .card-container::-webkit-scrollbar {
            width: 4px;
        }
        
        .card-container::-webkit-scrollbar-thumb {
            background: #ccc;
            border-radius: 2px;
        }
        
        .kanban-card {
            background: white;
            border-radius: 4px;
            padding: 6px;
            margin-bottom: 5px;
            box-shadow: 0 1px 2px rgba(0,0,0,0.1);
            cursor: grab;
            font-size: 10px;
            transition: all 0.2s;
            border-left: 3px solid;
            position: relative;
            break-inside: avoid;
        }
        
        .kanban-card:hover {
            transform: translateY(-1px);
            box-shadow: 0 2px 4px rgba(0,0,0,0.15);
        }
        
        .kanban-card.dragging {
            opacity: 0.5;
            transform: rotate(3deg);
        }
        
        .kanban-card-reguler {
            border-left-color: #28a745;
            background: linear-gradient(to right, #f0fff4, white);
        }
        
        .kanban-card-eksekutif {
            border-left-color: #007bff;
            background: linear-gradient(to right, #f0f8ff, white);
        }
        
        .kanban-card-over {
            border-left-color: #dc3545;
            background: linear-gradient(to right, #fff5f5, white);
        }
        
        .card-header {
            font-weight: bold;
            font-size: 9px;
            margin-bottom: 3px;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
            color: #333;
        }
        
        .card-details {
            font-size: 8px;
            color: #666;
            line-height: 1.2;
        }
        
        .card-icon {
            position: absolute;
            top: 4px;
            right: 4px;
            font-size: 8px;
        }
        
        .empty-slot {
            color: #999;
            font-style: italic;
            text-align: center;
            padding: 10px;
            font-size: 9px;
            border: 1px dashed #ddd;
            border-radius: 4px;
            margin: 5px 0;
        }
        
        .section-divider {
            height: 2px;
            background: linear-gradient(to right, transparent, #ddd, transparent);
            margin: 8px 0;
        }
        
        .type-label {
            font-size: 8px;
            font-weight: bold;
            color: white;
            padding: 1px 4px;
            border-radius: 2px;
            margin-bottom: 3px;
            display: inline-block;
        }
        
        .type-reguler {
            background: #28a745;
        }
        
        .type-eksekutif {
            background: #007bff;
        }
        </style>
        """
        
        # Create the kanban board for all time slots
        st.markdown("**🎯 Seret dan lepas kartu untuk memindahkan jadwal**")
        
        # Create a container for the full kanban board
        kanban_html = "<div class='kanban-board'>"
        
        for slot in SLOTS:
            cards = lanes.get(slot, [])
            
            # Separate cards by type
            reguler_cards = [c for c in cards if c.get("Kode") == "R"]
            eksekutif_cards = [c for c in cards if c.get("Kode") == "E"]
            
            # Column data for JavaScript
            column_data = json.dumps({
                "hari": selected_day,
                "slot": slot
            })
            
            kanban_html += f"""
            <div class='kanban-column' data-column='{column_data}'>
                <div class='slot-header'>{slot}</div>
                <div class='card-container'>
            """
            
            # Add Reguler cards first (on top)
            if reguler_cards:
                kanban_html += f"""<div class='type-label type-reguler'>REGULER ({len(reguler_cards)})</div>"""
                for card in reguler_cards:
                    card_class = "kanban-card kanban-card-reguler"
                    if card.get("Over"):
                        card_class += " kanban-card-over"
                    
                    status_icon = "🟢" if card.get("Kode") == "R" else "🔵"
                    if card.get("Over"):
                        status_icon = "🔴"
                    
                    card_data = json.dumps(card)
                    
                    kanban_html += f"""
                    <div class='{card_class}' data-card='{card_data}'>
                        <div class='card-header' title='{card['Dokter']}'>{card['Dokter']}</div>
                        <div class='card-details'>
                            <div><strong>{card['Poli']}</strong></div>
                        </div>
                        <div class='card-icon'>{status_icon}</div>
                    </div>
                    """
            
            # Add divider between Reguler and Eksekutif
            if reguler_cards and eksekutif_cards:
                kanban_html += "<div class='section-divider'></div>"
            
            # Add Eksekutif cards (below Reguler)
            if eksekutif_cards:
                kanban_html += f"""<div class='type-label type-eksekutif'>EKSEKUTIF ({len(eksekutif_cards)})</div>"""
                for card in eksekutif_cards:
                    card_class = "kanban-card kanban-card-eksekutif"
                    if card.get("Over"):
                        card_class += " kanban-card-over"
                    
                    status_icon = "🔵"
                    if card.get("Over"):
                        status_icon = "🔴"
                    
                    card_data = json.dumps(card)
                    
                    kanban_html += f"""
                    <div class='{card_class}' data-card='{card_data}'>
                        <div class='card-header' title='{card['Dokter']}'>{card['Dokter']}</div>
                        <div class='card-details'>
                            <div><strong>{card['Poli']}</strong></div>
                        </div>
                        <div class='card-icon'>{status_icon}</div>
                    </div>
                    """
            
            if not reguler_cards and not eksekutif_cards:
                kanban_html += "<div class='empty-slot'>Kosong</div>"
            
            kanban_html += "</div></div>"  # Close card-container and kanban-column
        
        kanban_html += "</div>"  # Close kanban-board
        
        # Add JavaScript and HTML
        st.components.v1.html(drag_drop_js + kanban_html, height=550, scrolling=False)
        
        # Legend
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.markdown("🟢 **Reguler**")
        with col2:
            st.markdown("🔵 **Eksekutif**")
        with col3:
            st.markdown("🔴 **Over Kuota** (Poleks > 7)")
        with col4:
            st.markdown("**Aturan:** Poleks maksimal 7 dokter per slot")
    
    with tab2:
        # Manual move section
        st.markdown("### 📝 Edit Manual")
        
        # Collect all cards for manual move
        all_cards = []
        for s in SLOTS:
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
                    time_display = card['Jam']
                    jenis_display = "REG" if card.get("Kode") == "R" else "EKS"
                    display_text = f"[{time_display}] {card['Dokter'][:20]} - {card['Poli'][:15]} ({jenis_display})"
                    if card.get("Over"):
                        display_text += " 🔴"
                    options.append(i)
                    display_texts.append(display_text)
                
                selected_index = st.selectbox(
                    "Pilih jadwal dokter:",
                    options=options,
                    format_func=lambda x: display_texts[x],
                    key="manual_select"
                )
                
                if selected_index is not None:
                    selected_card = all_cards[selected_index]
                    card_type = "Reguler" if selected_card.get("Kode") == "R" else "Eksekutif"
                    st.info(f"**Terpilih:** {selected_card['Dokter']}\n- Poli: {selected_card['Poli']}\n- Tipe: {card_type}\n- Jam: {selected_card['Jam']}")
            
            with col_move:
                # Show current slot
                if selected_index is not None:
                    selected_card = all_cards[selected_index]
                    st.write(f"**Slot saat ini:** {selected_card['Jam']}")
                
                target_slot = st.selectbox(
                    "Pindahkan ke jam:",
                    SLOTS,
                    index=0 if SLOTS else None,
                    key="target_slot"
                )
                
                # Check for over kuota if moving Eksekutif to target slot
                if selected_index is not None:
                    selected_card = all_cards[selected_index]
                    if selected_card.get("Kode") == "E":  # Eksekutif
                        eksekutif_in_target = len([c for c in lanes.get(target_slot, []) if c.get("Kode") == "E"])
                        if eksekutif_in_target >= 7:
                            st.warning(f"⚠️ Slot {target_slot} sudah ada {eksekutif_in_target} dokter Poleks (maksimal 7)")
                
                col_move1, col_move2 = st.columns(2)
                with col_move1:
                    if st.button("🚀 Pindahkan", type="primary", use_container_width=True):
                        card_to_move = all_cards[selected_index]
                        original_slot = card_to_move["Jam"]
                        
                        # Remove from original slot
                        if original_slot in st.session_state.kanban_state[selected_day]:
                            st.session_state.kanban_state[selected_day][original_slot] = [
                                c for c in st.session_state.kanban_state[selected_day][original_slot]
                                if c.get("id") != card_to_move.get("id")
                            ]
                        
                        # Add to target slot
                        new_card = {
                            "id": f"{selected_day}|{target_slot}|{np.random.randint(1e9)}",
                            "Dokter": card_to_move["Dokter"],
                            "Poli": card_to_move["Poli"],
                            "Jenis": card_to_move["Jenis"],
                            "Kode": card_to_move.get("Kode", "E"),
                            "Over": False,
                            "Hari": selected_day,
                            "Jam": target_slot
                        }
                        
                        if target_slot not in st.session_state.kanban_state[selected_day]:
                            st.session_state.kanban_state[selected_day][target_slot] = []
                        
                        st.session_state.kanban_state[selected_day][target_slot].append(new_card)
                        
                        # Rebuild df and compute status
                        df_other = df[df["Hari"] != selected_day].copy()
                        new_rows = []
                        
                        for s in SLOTS:
                            for c in st.session_state.kanban_state[selected_day][s]:
                                new_rows.append({
                                    "Hari": selected_day,
                                    "Jam": s,
                                    "Poli": c.get("Poli", ""),
                                    "Jenis": c.get("Jenis", ""),
                                    "Dokter": c.get("Dokter", "")
                                })
                        
                        df_new = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
                        df_new = compute_status(df_new)
                        
                        # Update Excel view
                        excel_view_df = create_excel_like_view(df_new)
                        st.session_state.excel_view_df = excel_view_df
                        
                        # Update main df
                        df = df_new.copy()
                        
                        # Add to history
                        push_history(df.copy())
                        
                        st.success(f"✅ {card_to_move['Dokter']} dipindahkan dari {original_slot} ke {target_slot}")
                        st.rerun()
                
                with col_move2:
                    if st.button("🗑️ Hapus", type="secondary", use_container_width=True):
                        card_to_delete = all_cards[selected_index]
                        original_slot = card_to_delete["Jam"]
                        
                        # Remove from slot
                        if original_slot in st.session_state.kanban_state[selected_day]:
                            st.session_state.kanban_state[selected_day][original_slot] = [
                                c for c in st.session_state.kanban_state[selected_day][original_slot]
                                if c.get("id") != card_to_delete.get("id")
                            ]
                        
                        # Rebuild df
                        df_other = df[df["Hari"] != selected_day].copy()
                        new_rows = []
                        
                        for s in SLOTS:
                            for c in st.session_state.kanban_state[selected_day][s]:
                                new_rows.append({
                                    "Hari": selected_day,
                                    "Jam": s,
                                    "Poli": c.get("Poli", ""),
                                    "Jenis": c.get("Jenis", ""),
                                    "Dokter": c.get("Dokter", "")
                                })
                        
                        df_new = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
                        df_new = compute_status(df_new)
                        
                        # Update Excel view
                        excel_view_df = create_excel_like_view(df_new)
                        st.session_state.excel_view_df = excel_view_df
                        
                        # Update main df
                        df = df_new.copy()
                        
                        # Add to history
                        push_history(df.copy())
                        
                        st.success(f"🗑️ {card_to_delete['Dokter']} dihapus dari {original_slot}")
                        st.rerun()
        else:
            st.info("Tidak ada jadwal untuk hari ini.")

# ---------------------------
# Export buttons
# ---------------------------
st.markdown("---")
st.header("💾 Export & Simpan")

export_cols = st.columns(3)

with export_cols[0]:
    csv_bytes = df.to_csv(index=False).encode("utf-8")
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
    
    xlsx_bytes = to_xlsx_bytes(df)
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
    
    excel_view_xlsx = to_excel_view_xlsx(st.session_state.excel_view_df)
    st.download_button(
        "📊 Download Excel View",
        excel_view_xlsx,
        file_name="jadwal_excel_view.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ---------------------------
# Quick Stats
# ---------------------------
st.markdown("---")
st.subheader("📋 Statistik Cepat")

# Create a summary by Poli
try:
    poli_summary = df.groupby(["Poli", "Jenis"]).size().reset_index(name="Jumlah Slot")
    if not poli_summary.empty:
        poli_pivot = poli_summary.pivot(index="Poli", columns="Jenis", values="Jumlah Slot").fillna(0)
        
        # Display as metrics
        st.write("**Jumlah Slot per Poli:**")
        cols = st.columns(min(4, len(poli_pivot)))
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
