# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io
import json

# Mengatur tampilan halaman ke mode lebar (wide)
st.set_page_config(page_title="Jadwal Poli (Streamlit Full)", layout="wide")
st.title("📅 Jadwal Poli — Streamlit Full (Offline)")

# ---------------------------
# Constants for Excel-like view (FULL RANGE)
# ---------------------------
TIME_SLOTS = [
    "07:30", "08:00", "08:30", "09:00", "09:30", "10:00", "10:30", 
    "11:00", "11:30", "12:00", "12:30", "13:00", "13:30", "14:00", "14:30"
]

DAYS_ORDER = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]

# ---------------------------
# Helpers: time normalization & expansion (TIDAK BERUBAH)
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

# ---------------------------
# Convert to Excel-like format (Pivot table) (TIDAK BERUBAH)
# ---------------------------
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
        # Sorting for Poleks below Reguler is handled later in Kanban, not here
        for _, row in group.iterrows():
            time_slot = row["Jam"]
            kode = row["Kode"]
            # Handle multiple entries in one cell by prioritizing Poleks if both exist (though ideally shouldn't happen)
            if record[time_slot] == "":
                record[time_slot] = kode
            elif record[time_slot] == "R" and kode == "E":
                record[time_slot] = "E" # If both, show E (Eksekutif/Poleks)
            
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

# ---------------------------
# Session state: history (undo/redo), kanban_state (CLEARED/SIMPLIFIED)
# ---------------------------
if "history" not in st.session_state:
    st.session_state.history = []
if "future" not in st.session_state:
    st.session_state.future = []
if "excel_view_df" not in st.session_state:
    st.session_state.excel_view_df = pd.DataFrame()
if "last_drag_event" not in st.session_state:
    st.session_state.last_drag_event = None
# kanban_state is no longer needed as an intermediate state, rely on df history

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
# Drag & Drop Functions (UPDATED FOR V2 GRID STRUCTURE)
# ---------------------------
def handle_drag_drop(drag_data_str, current_df):
    """Handle drag and drop events based on new Excel-like Kanban structure."""
    try:
        drag_data = json.loads(drag_data_str)
        st.session_state.last_drag_event = drag_data
        
        # Data needed for drag:
        source_day = drag_data.get("source_day")
        source_slot = drag_data.get("source_slot")
        
        target_day = drag_data.get("target_day")
        target_slot = drag_data.get("target_slot")
        
        card_data = drag_data.get("card_data")
        
        # Unique identifier for the doctor/poli combination (the fixed row)
        poli = card_data.get("Poli")
        jenis = card_data.get("Jenis")
        dokter = card_data.get("Dokter")
        
        # --- Check for target match (should be same day, same doctor/poli/jenis) ---
        if target_day != source_day:
             # This should be prevented by JS but is a good safety check
             st.warning("Perpindahan jadwal antar hari tidak diizinkan di tampilan ini.")
             return False
        
        # If the target is the same as the source, do nothing
        if target_slot == source_slot:
            return True

        # --- Update the main DataFrame (df) ---
        df_new = current_df.copy()

        # Identify the row(s) to be updated/removed:
        filter_mask = (df_new["Hari"] == source_day) & \
                      (df_new["Jam"] == source_slot) & \
                      (df_new["Poli"] == poli) & \
                      (df_new["Jenis"] == jenis) & \
                      (df_new["Dokter"] == dokter)

        # Drop the row(s) matching the source slot (remove the old entry)
        df_new = df_new[~filter_mask].reset_index(drop=True)
        
        # 2. Add the new entry at the target slot
        new_row = {
            "Hari": target_day,
            "Jam": target_slot,
            "Poli": poli,
            "Jenis": jenis,
            "Dokter": dokter,
            "Kode": card_data.get("Kode", "E")
        }
        
        df_new = pd.concat([df_new, pd.DataFrame([new_row])], ignore_index=True)
        
        # 3. Re-compute status
        df_new = compute_status(df_new)
        
        # 4. Update Excel view and history
        excel_view_df = create_excel_like_view(df_new)
        st.session_state.excel_view_df = excel_view_df
        push_history(df_new.copy())
        
        return True
    
    except Exception as e:
        st.error(f"Error handling drag drop: {e}")
        return False

# ---------------------------
# Upload input (TIDAK BERUBAH)
# ---------------------------
st.sidebar.header("Upload / Template")
uploaded = st.sidebar.file_uploader("Upload Excel (sheet) or CSV with columns: Hari, Range, Poli, Jenis, Dokter", type=["xlsx","csv"])
if st.sidebar.button("Download template example"):
    sample = pd.DataFrame({
        "Hari":["Senin","Senin","Selasa"],
        "Range":["07.30-09.00","09.00-11.00","07.00-08.30"],
        "Poli":["Anak","Anak","Gigi"],
        "Jenis":["Reguler","Eksekutif","Reguler"],
        "Dokter":["dr. Budi","dr. Sari","drg. Putri"]
    })
    example_range = f"{TIME_SLOTS[0]}-{TIME_SLOTS[-1]}" 
    sample_full_range = pd.DataFrame({
        "Hari":["Rabu","Rabu"],
        "Range":[example_range, example_range],
        "Poli":["Penyakit Dalam","Jantung"],
        "Jenis":["Reguler","Eksekutif"],
        "Dokter":["dr. Dedi","dr. Siti"]
    })
    sample = pd.concat([sample, sample_full_range], ignore_index=True)
    st.sidebar.download_button("Download template.xlsx", data=sample.to_excel(index=False, engine="openpyxl"), file_name="template_jadwal.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

if uploaded is None:
    st.info("Upload file Excel/CSV untuk mulai. Gunakan tombol 'Download template example' jika perlu contoh.")
    st.stop()

# read file tolerant
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

# tolerant column mapping
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
    st.error("Kolom input tidak lengkap. Pastikan file memiliki kolom: Hari, Range, Poli, Jenis, Dokter (varian nama didukung).")
    st.write("Terbaca kolom:", cols)
    st.stop()

# ---------------------------
# Expand ranges -> slots (TIDAK BERUBAH)
# ---------------------------
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
        expanded.append({"Hari":hari, "Jam":s, "Poli":poli, "Jenis":jenis, "Dokter":dokter})

if len(expanded) == 0:
    st.error("Tidak ada slot terbentuk. Periksa format Range.")
    st.stop()

df = pd.DataFrame(expanded)
# normalize Jenis
df["Jenis"] = df["Jenis"].astype(str).str.strip().replace({
    "reguler":"Reguler", "regular":"Reguler", 
    "eksekutif":"Eksekutif", "executive":"Eksekutif", 
    "poleks":"Eksekutif", "POLEKS":"Eksekutif"
})
# Kode
df["Kode"] = df["Jenis"].apply(lambda x: "R" if str(x).lower()=="reguler" else "E")

# ---------------------------
# Compute Over-kuota (Bentrok removed) (TIDAK BERUBAH)
# ---------------------------
def compute_status(df_in):
    d = df_in.copy()
    d["Over_Kuota"] = False
    d["Bentrok"] = False # Keep Bentrok column, but always False
    
    # over: count Eksekutif entries per (Hari,Jam)
    eksek = d[d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)]
    poleks_counts = eksek.groupby(["Hari","Jam"]).size()
    
    # Threshold for Over_Kuota (Example: > 7)
    OVER_KUOTA_THRESHOLD = 7
    
    over_slots = poleks_counts[poleks_counts > OVER_KUOTA_THRESHOLD].index if not poleks_counts.empty else []
    
    for (hari,jam) in over_slots:
        d.loc[(d["Hari"]==hari)&(d["Jam"]==jam)&(d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)), "Over_Kuota"] = True
    
    return d

df = compute_status(df)

# Create Excel-like view
excel_view_df = create_excel_like_view(df)
st.session_state.excel_view_df = excel_view_df

# push initial snapshot to history
push_history(df.copy())

# Ensure df is the latest version from history
if st.session_state.history:
    df = st.session_state.history[-1]

# ---------------------------
# UI: Filters & actions (DI PERTAHANKAN DI SIDEBAR KIRI)
# ---------------------------
st.sidebar.header("Filter & Actions")
# Populate selected day with all available days, not just days with data
all_available_days = sorted(df["Hari"].unique())
selected_day = st.sidebar.selectbox("Pilih Hari (kanban)", ["--Semua--"] + all_available_days)
poli_filter = st.sidebar.multiselect("Filter Poli (opsional)", sorted(df["Poli"].unique()), default=list(df["Poli"].unique()))
jenis_filter = st.sidebar.multiselect("Filter Jenis", sorted(df["Jenis"].unique()), default=list(df["Jenis"].unique()))

# undo/redo buttons
colu1, colu2, colu3 = st.sidebar.columns([1,1,2])
with colu1:
    if st.button("Undo"):
        prev = undo()
        if prev is not None:
            df = prev.copy()
            # Recreate Excel view and re-compute status
            df = compute_status(df)
            excel_view_df = create_excel_like_view(df)
            st.session_state.excel_view_df = excel_view_df
            st.success("Undo berhasil")
            st.rerun() # Rerun to update the view
with colu2:
    if st.button("Redo"):
        nxt = redo()
        if nxt is not None:
            df = nxt.copy()
            # Recreate Excel view and re-compute status
            df = compute_status(df)
            excel_view_df = create_excel_like_view(df)
            st.session_state.excel_view_df = excel_view_df
            st.success("Redo berhasil")
            st.rerun() # Rerun to update the view

# ---------------------------
# Excel-like View (TIDAK BERUBAH)
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

# Summary statistics for Excel view (TIDAK BERUBAH)
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
# Dashboard & summary (TIDAK BERUBAH)
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
# KANBAN EDITOR BARU (GRID/EXCEL-LIKE)
# ---------------------------
st.header("🎯 Kanban Editor (Drag & Drop)")

if selected_day == "--Semua--":
    st.info("Pilih satu hari di sidebar untuk membuka Kanban editor.")
else:
    SLOTS = TIME_SLOTS
    
    # Filter the main DataFrame for the selected day
    df_day = df[df["Hari"]==selected_day].reset_index(drop=True)
    
    # 1. Prepare row data (unique Poli/Jenis/Dokter for the day)
    # Sorting to ensure Poleks/Eksekutif is below Reguler in the row list
    unique_rows = df_day[["Poli", "Jenis", "Dokter"]].drop_duplicates().sort_values(["Poli", "Jenis", "Dokter"]).reset_index(drop=True)
    
    # 2. Get card data grouped by the row keys and slot
    card_map = {}
    for i, r in df_day.iterrows():
        # Key is the Poli, Jenis, Dokter combination
        key = (r["Poli"], r["Jenis"], r["Dokter"])
        slot = r["Jam"]
        
        # Create or update the card data for the cell (should only be one card per unique (key, slot))
        if key not in card_map:
            card_map[key] = {}
            
        # The new structure only shows ONE card per cell (Poli/Dokter/Jenis/Slot)
        card_map[key][slot] = {
            "id": f"{selected_day}|{slot}|{i}|{np.random.randint(1e9)}",
            "Dokter": r["Dokter"],
            "Poli": r["Poli"],
            "Jenis": r["Jenis"],
            "Kode": r["Kode"],
            "Over": bool(r["Over_Kuota"]),
            "Bentrok": bool(r["Bentrok"]),
            "Hari": selected_day,
            "Jam": slot
        }

    # Custom CSS for the new grid layout
    new_kanban_css = f"""
        <style>
        /* New Grid/Table Layout CSS */
        .kanban-grid-container {{
            overflow-x: auto;
            padding: 5px;
            font-size: 10px;
            min-height: 450px;
            max-height: 600px;
            overflow-y: auto;
            border: 1px solid #ddd;
            border-radius: 5px;
        }}

        .kanban-grid {{
            display: grid;
            /* Define columns: 3 fixed-width columns for info + N time slots (70px each) */
            grid-template-columns: 150px 100px 150px repeat({len(TIME_SLOTS)}, 70px); 
            gap: 0;
            min-width: calc(150px + 100px + 150px + {len(TIME_SLOTS) * 70}px); 
        }}
        
        .grid-header, .grid-cell {{
            border: 1px solid #eee;
            padding: 4px;
            display: flex;
            align-items: center;
            justify-content: center;
            text-align: center;
            font-weight: normal;
            overflow: hidden;
            white-space: nowrap;
            text-overflow: ellipsis;
        }}
        
        .grid-header {{
            background-color: #4CAF50; /* Green header */
            color: white;
            font-weight: bold;
            position: sticky;
            top: 0;
            z-index: 10;
            border-bottom: 2px solid #38761D;
            height: 30px;
        }}

        .grid-row-header {{
            background-color: #f2f2f2;
            font-weight: bold;
            text-align: left;
            justify-content: flex-start;
            padding: 4px;
            position: sticky;
            left: 0;
            z-index: 5;
            height: 30px;
        }}

        /* Column-specific headers */
        .poli-header {{ grid-column: 1; }}
        .jenis-header {{ grid-column: 2; }}
        .dokter-header {{ grid-column: 3; }}

        .grid-cell {{
            background-color: white;
            height: 30px; /* Fixed cell height */
            padding: 0; 
        }}
        
        .droppable-cell {{
            background-color: #f8f9fa;
            border: none;
            height: 100%;
            width: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            transition: background-color 0.1s;
        }}
        
        .droppable-cell.drag-over {{
            background-color: #e3f2fd;
            border: 2px dashed #2196f3;
            box-sizing: border-box;
        }}

        /* Small card styling for table cell */
        .kanban-card {{
            border-radius: 2px;
            padding: 0; 
            margin: 0;
            box-shadow: 0 0 1px rgba(0,0,0,0.1);
            cursor: grab;
            font-size: 10px;
            line-height: 1.1;
            transition: all 0.2s;
            border-left: 2px solid;
            width: 100%;
            height: 100%;
            display: flex;
            align-items: center;
            justify-content: center;
            font-weight: bold;
        }}
        
        .kanban-card-reguler {{ border-left-color: #28a745; background-color: #e6ffed; }} /* Light green */
        .kanban-card-eksekutif {{ border-left-color: #007bff; background-color: #e0f2ff; }} /* Light blue (Poleks) */
        .kanban-card-over {{ border-left-color: #dc3545; background-color: #ffe6e8; }} /* Light red */
        
        .kanban-card-text {{
            font-size: 10px;
            padding: 2px;
        }}
        
        .kanban-card.dragging {{
            opacity: 0.5;
        }}
        </style>
    """

    # Custom JavaScript for the new grid structure
    new_drag_drop_js = f"""
        <script>
        // Drag and Drop functionality
        function setupDragAndDrop() {{
            const cards = document.querySelectorAll('.kanban-card');
            const droppableCells = document.querySelectorAll('.droppable-cell');
            
            // Setup draggable cards
            cards.forEach(card => {{
                card.setAttribute('draggable', 'true');
                
                card.addEventListener('dragstart', (e) => {{
                    const cardData = JSON.parse(card.getAttribute('data-card'));
                    e.dataTransfer.setData('text/plain', JSON.stringify({{
                        source_day: cardData.Hari,
                        source_slot: cardData.Jam,
                        source_poli: cardData.Poli,
                        source_jenis: cardData.Jenis,
                        card_data: cardData
                    }}));
                    card.classList.add('dragging');
                }});
                
                card.addEventListener('dragend', () => {{
                    card.classList.remove('dragging');
                }});
            }});
            
            // Setup droppable columns/cells
            droppableCells.forEach(cell => {{
                cell.addEventListener('dragover', (e) => {{
                    e.preventDefault();
                    cell.classList.add('drag-over');
                }});
                
                cell.addEventListener('dragleave', () => {{
                    cell.classList.remove('drag-over');
                }});
                
                cell.addEventListener('drop', (e) => {{
                    e.preventDefault();
                    cell.classList.remove('drag-over');
                    
                    try {{
                        const dragData = JSON.parse(e.dataTransfer.getData('text/plain'));
                        const cellData = JSON.parse(cell.getAttribute('data-cell'));
                        
                        // Prevent dropping onto an already occupied cell
                        if (cell.querySelector('.kanban-card')) {{
                             console.log("Cannot drop: Target cell is already occupied.");
                             return; 
                        }}

                        // TARGET MUST MATCH SOURCE POLI/JENIS/DOKTER (Fixed Row)
                        if (dragData.card_data.Dokter !== cellData.dokter ||
                            dragData.card_data.Poli !== cellData.poli ||
                            dragData.card_data.Jenis !== cellData.jenis) {{
                            console.log("Cannot drop: Doctor/Poli/Jenis mismatch. Must drop into the same row.");
                            return; // Prevent dropping outside the doctor's row
                        }}
                        
                        // Update drag data with target
                        dragData.target_day = cellData.hari;
                        dragData.target_slot = cellData.slot;
                        
                        // Send to Streamlit
                        if (window.parent.document.getElementById('streamlit-kanban-bridge')) {{
                            window.parent.document.getElementById('streamlit-kanban-bridge').value = JSON.stringify(dragData);
                            window.parent.document.getElementById('streamlit-kanban-bridge').dispatchEvent(new Event('change'));
                        }} else {{
                            console.error('Streamlit bridge element not found.');
                        }}
                    }} catch (error) {{
                        console.error('Drop error:', error);
                    }}
                }});
            }});
        }}
        
        // Initialize when page loads and after Streamlit updates
        setupDragAndDrop();
        </script>
    """
    
    # 3. Create the kanban board HTML (Grid Structure)
    
    # Header row
    kanban_html = "<div class='kanban-grid-container'><div class='kanban-grid'>"
    kanban_html += "<div class='grid-header poli-header'>Poli</div>"
    kanban_html += "<div class='grid-header jenis-header'>Jenis</div>"
    kanban_html += "<div class='grid-header dokter-header'>Dokter</div>"
    for slot in TIME_SLOTS:
        kanban_html += f"<div class='grid-header'>{slot}</div>"
        
    # Data rows
    for i, r in unique_rows.iterrows():
        # Row Headers (Poli, Jenis, Dokter)
        kanban_html += f"<div class='grid-row-header' title='{r['Poli']}'>{r['Poli']}</div>"
        kanban_html += f"<div class='grid-row-header' title='{r['Jenis']}'>{r['Jenis']}</div>"
        kanban_html += f"<div class='grid-row-header' title='{r['Dokter']}'>{r['Dokter']}</div>"
        
        # Time Slots (Droppable Cells)
        row_key = (r['Poli'], r['Jenis'], r['Dokter'])
        for slot in TIME_SLOTS:
            cell_data = json.dumps({
                "hari": selected_day,
                "slot": slot,
                "poli": r['Poli'],
                "jenis": r['Jenis'],
                "dokter": r['Dokter']
            })
            
            card = card_map.get(row_key, {}).get(slot)
            
            card_content = ""
            if card:
                card_class = "kanban-card"
                if card.get("Over"):
                    card_class += " kanban-card-over"
                elif card.get("Kode") == "R":
                    card_class += " kanban-card-reguler"
                else:
                    # Poleks/Eksekutif
                    card_class += " kanban-card-eksekutif" 
                
                # Card data for JavaScript
                card_data = json.dumps(card)
                
                # Show Kode (R/E) in the cell
                card_content = f"""
                <div class='{card_class}' draggable='true' data-card='{card_data}' title='{card['Dokter']}'>
                    <span class='kanban-card-text'>{card['Kode']}</span>
                </div>
                """
            
            kanban_html += f"""
            <div class='grid-cell'>
                <div class='droppable-cell' data-cell='{cell_data}'>
                    {card_content}
                </div>
            </div>
            """
    
    kanban_html += "</div></div>"
    
    # Hidden input to receive drag data from custom JS
    drag_data_receiver = st.empty()
    drag_data_str = drag_data_receiver.text_input("Drag Drop Receiver (Hidden)", key="streamlit-kanban-bridge", label_visibility="collapsed")
    
    if drag_data_str and drag_data_str != st.session_state.get('last_processed_drag_data'):
        # Check if the event is new before processing
        st.session_state['last_processed_drag_data'] = drag_data_str
        # Pass the current DF to the handler
        if handle_drag_drop(drag_data_str, df): 
            st.success(f"Perpindahan jadwal berhasil dari slot {json.loads(drag_data_str)['source_slot']} ke {json.loads(drag_data_str)['target_slot']}! Status Over Kuota diperbarui.")
            st.session_state["streamlit-kanban-bridge"] = "" # Clear the input after processing
            st.rerun()
        else:
            st.warning("Perpindahan jadwal gagal.")
            
    # Render the kanban board
    st.components.v1.html(new_drag_drop_js + new_kanban_css + kanban_html, height=650, scrolling=False)
    
    # Manual move section is now obsolete/redundant with the Excel-like drag drop, so it's removed to simplify.

# ---------------------------
# Export buttons & Quick Stats (TIDAK BERUBAH)
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

try:
    poli_summary = df.groupby(["Poli", "Jenis"]).size().reset_index(name="Jumlah Slot")
    if not poli_summary.empty:
        poli_pivot = poli_summary.pivot(index="Poli", columns="Jenis", values="Jumlah Slot").fillna(0)
        
        st.write("**Jumlah Slot per Poli:**")
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
