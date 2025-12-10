# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

# Try to import sortables (drag & drop). If not available, fallback.
try:
    from sortables import sort_table
    DRAG_AVAILABLE = True
except Exception:
    DRAG_AVAILABLE = False

st.set_page_config(page_title="Jadwal Dokter (Range → Slots) — Kanban", layout="wide")

# ---------------------------
# Helpers: normalize & expand ranges
# ---------------------------
def normalize_time_token(token: str) -> str:
    """Normalize a single time token into HH:MM."""
    if token is None:
        return ""
    t = str(token).strip()
    if t == "":
        return ""
    t = t.replace(".", ":")
    t = t.lower().replace("am","").replace("pm","")
    # Remove spaces
    t = t.strip()
    # If just hour like '7' -> '07:00'
    if ":" not in t:
        if len(t) == 1:
            t = f"0{t}:00"
        elif len(t) == 2:
            t = f"{t}:00"
        else:
            # fallback try parse
            try:
                dt = datetime.strptime(t, "%H%M")
                t = dt.strftime("%H:%M")
            except Exception:
                pass
    else:
        parts = t.split(":")
        hh = parts[0].zfill(2)
        mm = parts[1].zfill(2)
        t = f"{hh}:{mm}"
    return t

def expand_time_range(range_str: str, interval_minutes:int=30):
    """
    Expand a range string like '07.30-09.00' or '7-9' into list of slots
    inclusive of both endpoints at step interval_minutes.
    """
    if not isinstance(range_str, str) or range_str.strip() == "":
        return []
    # Accept different dash types
    sep_candidates = ["-", "–", "—", " to "]
    for sep in sep_candidates:
        if sep in range_str:
            parts = range_str.split(sep)
            break
    else:
        # no separator found -> treat as single time token
        single = normalize_time_token(range_str)
        return [single] if single else []

    if len(parts) < 2:
        return []
    start_raw = parts[0].strip()
    end_raw = parts[1].strip()
    start_s = normalize_time_token(start_raw)
    end_s = normalize_time_token(end_raw)
    if start_s == "" or end_s == "":
        return []
    fmt = "%H:%M"
    try:
        start_dt = datetime.strptime(start_s, fmt)
        end_dt = datetime.strptime(end_s, fmt)
    except Exception:
        return []
    # If end < start, assume end same day but later (no overnight support)
    if end_dt < start_dt:
        # try swap? or assume invalid -> return empty
        return []

    slots = []
    cur = start_dt
    while cur <= end_dt:
        slots.append(cur.strftime(fmt))
        cur += timedelta(minutes=interval_minutes)
    return slots

# ---------------------------
# Read uploaded file & expand ranges
# ---------------------------
st.title("Jadwal Dokter (Range → 30-min Slots) — Kanban & Dashboard")

uploaded = st.file_uploader("Upload Excel (sheet with Range) atau CSV", type=["xlsx","csv"])

@st.cache_data
def load_raw_df(bytes_io, fname):
    try:
        if fname.lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(bytes_io))
        else:
            xls = pd.ExcelFile(io.BytesIO(bytes_io))
            # Prefer sheet named 'Jadwal' or first sheet
            sheet_name = "Jadwal" if "Jadwal" in xls.sheet_names else xls.sheet_names[0]
            df = xls.parse(sheet_name)
        # normalize column names
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return pd.DataFrame()

if uploaded:
    raw_df = load_raw_df(uploaded.getvalue(), uploaded.name)
else:
    st.info("Silakan unggah file Excel/CSV yang berisi kolom: Hari, Range, Poli, Dokter.")
    st.stop()

# Attempt to reconcile column names (support some variants)
col_map_candidates = {
    "Hari": ["Hari","Day","hari","day"],
    "Range": ["Range","Jam","Waktu","Time","range","jam","waktu","time"],
    "Poli": ["Poli","Poliklinik","Polyclinic","poli","poliklinik"],
    "Dokter": ["Dokter","Nama Dokter","Doctor","dokter","nama dokter","doctor"]
}

def find_col(df_cols, names):
    for n in names:
        if n in df_cols:
            return n
    return None

cols = raw_df.columns.tolist()
Hari_col = find_col(cols, col_map_candidates["Hari"])
Range_col = find_col(cols, col_map_candidates["Range"])
Poli_col = find_col(cols, col_map_candidates["Poli"])
Dokter_col = find_col(cols, col_map_candidates["Dokter"])

if not (Hari_col and Range_col and Poli_col and Dokter_col):
    st.error("Kolom input tidak sesuai. Pastikan file memiliki kolom: Hari, Range, Poli, Dokter (bisa menggunakan variasi nama).")
    st.write("Terbaca kolom:", cols)
    st.stop()

# Expand ranges into per-slot rows
expanded_rows = []
for idx, r in raw_df.iterrows():
    hari = str(r.get(Hari_col)).strip()
    range_raw = r.get(Range_col)
    poli = str(r.get(Poli_col)).strip()
    dokter = str(r.get(Dokter_col)).strip()
    if not hari or pd.isna(range_raw) or not poli or not dokter:
        continue
    slots = expand_time_range(str(range_raw), interval_minutes=30)
    if not slots:
        # if cannot parse as range, try parse as single time token
        tok = normalize_time_token(str(range_raw))
        if tok:
            slots = [tok]
    for s in slots:
        expanded_rows.append({
            "Hari": hari,
            "Jam": s,
            "Poli": poli,
            "Dokter": dokter
        })

if len(expanded_rows) == 0:
    st.error("Tidak ada slot yang berhasil dibentuk dari Range. Periksa format Range.")
    st.stop()

df = pd.DataFrame(expanded_rows)

# ---------------------------
# Compute Kode/Status
# ---------------------------
def compute_status_codes(df):
    d = df.copy()
    # Kode: Reguler -> R, else E
    d["Kode"] = d["Poli"].astype(str).str.contains("reguler", case=False).map({True:"R", False:"E"})
    # Count per Hari+Jam
    d["Count"] = d.groupby(["Hari","Jam"])["Dokter"].transform("count")
    d["Status"] = ""
    d.loc[d["Count"] > 7, "Status"] = "Over Kuota"
    # If more than 1 unique dokter in slot -> E entries become Bentrok
    grouped = d.groupby(["Hari","Jam"])["Dokter"].nunique().reset_index(name="nunique")
    clashes = grouped[grouped["nunique"] > 1]
    if not clashes.empty:
        for _, row in clashes.iterrows():
            h = row["Hari"]; j = row["Jam"]
            mask = (d["Hari"]==h) & (d["Jam"]==j) & (d["Kode"]=="E")
            d.loc[mask, "Status"] = "Bentrok"
    return d

df = compute_status_codes(df)

# ---------------------------
# UI: Left pane controls & Dashboard
# ---------------------------
st.sidebar.header("Kontrol")
selected_day = st.sidebar.selectbox("Pilih Hari (untuk Kanban editing)", sorted(df["Hari"].unique()))
download_all_name = st.sidebar.text_input("Nama file export (tanpa ekstensi)", "jadwal_updated")

st.markdown("## Ringkasan")
col1, col2, col3 = st.columns(3)
col1.metric("Total Baris (slot)", len(df))
col2.metric("Total Dokter (unik)", df["Dokter"].nunique())
col3.metric("Unique Slots", df[["Hari","Jam"]].drop_duplicates().shape[0])

# Dashboard: heatmap & chart
st.header("Dashboard")
summary = df.groupby(["Hari","Jam"]).size().reset_index(name="Jumlah")
pivot = summary.pivot(index="Jam", columns="Hari", values="Jumlah").fillna(0).sort_index()

st.subheader("Heatmap (Jam × Hari)")
import plotly.express as px
fig = px.imshow(
    pivot,
    labels=dict(x="Hari", y="Jam", color="Jumlah Dokter"),
    aspect="auto",
    color_continuous_scale="Blues"
)
st.plotly_chart(fig, use_container_width=True, height=450)

st.subheader("Jumlah Dokter per Jam (line per Hari)")
chart = px.line(summary, x="Jam", y="Jumlah", color="Hari", markers=True)
st.plotly_chart(chart, use_container_width=True, height=350)

# ---------------------------
# Kanban: build lanes for selected_day
# ---------------------------
st.header(f"Kanban Editor — {selected_day}")
SLOTS = []
# construct slots between min and max from data for that day, or full day default 07:00-14:30
day_slots = sorted(df[df["Hari"]==selected_day]["Jam"].unique())
if len(day_slots) >= 1:
    SLOTS = sorted(day_slots, key=lambda x: datetime.strptime(x,"%H:%M"))
else:
    # fallback default
    def make_standard_slots():
        s = datetime.strptime("07:00","%H:%M")
        e = datetime.strptime("14:30","%H:%M")
        slots=[]
        cur=s
        while cur<=e:
            slots.append(cur.strftime("%H:%M"))
            cur+=timedelta(minutes=30)
        return slots
    SLOTS = make_standard_slots()

# Prepare lanes dict: slot -> list of cards
def lanes_from_df_for_day(df, day):
    lanes = {}
    for slot in SLOTS:
        rows = df[(df["Hari"]==day) & (df["Jam"]==slot)].reset_index(drop=True)
        cards=[]
        for i,r in rows.iterrows():
            cards.append({
                "id": f"{day}|{slot}|{i}|{np.random.randint(1e9)}",
                "Dokter": r["Dokter"],
                "Poli": r["Poli"],
                "Kode": r["Kode"],
                "Status": r["Status"]
            })
        lanes[slot] = cards
    return lanes

if "kanban_state" not in st.session_state:
    st.session_state["kanban_state"] = {}

if selected_day not in st.session_state["kanban_state"]:
    st.session_state["kanban_state"][selected_day] = lanes_from_df_for_day(df, selected_day)

st.write("Seret kartu antar kolom untuk memindahkan dokter ke slot waktu lain. Jika drag tidak tersedia, gunakan mode 'Select & Move' di bawah.")

# Render Kanban columns
cols = st.columns(len(SLOTS))
new_state = {s:list(st.session_state["kanban_state"][selected_day].get(s,[])) for s in SLOTS}
moved_any = False

if DRAG_AVAILABLE:
    # Use sort_table per column to allow reordering within column (some libs may support cross-list drag)
    for i, s in enumerate(SLOTS):
        with cols[i]:
            st.markdown(f"**{s}**")
            cards = new_state[s]
            if len(cards)==0:
                st.info("—")
            else:
                tab = pd.DataFrame(cards)
                # attempt to call sort_table; library behaviors vary across versions
                try:
                    sorted_tab, moved = sort_table(tab, key="id", height="300px")
                    new_cards = sorted_tab.to_dict(orient="records")
                    new_state[s] = new_cards
                    if moved:
                        moved_any = True
                except Exception:
                    # fallback: show simple table (no drag)
                    st.dataframe(tab[["Dokter","Poli","Kode","Status"]], use_container_width=True)
else:
    st.warning("Drag & drop tidak tersedia. Gunakan mode Select & Move di bawah.")
    # show simple lists per column
    for i,s in enumerate(SLOTS):
        with cols[i]:
            st.markdown(f"**{s}**")
            cards = new_state[s]
            if len(cards)==0:
                st.info("—")
            else:
                for c in cards:
                    # render compact card
                    color_box = ""
                    # select color small badge
                    if c["Status"] in ["Bentrok","Over Kuota"]:
                        color_box = "🔴"
                    elif c["Kode"]=="R":
                        color_box = "🟢"
                    elif c["Kode"]=="E" and "poleks" in c["Poli"].lower():
                        color_box = "🔵"
                    else:
                        color_box = "⚪"
                    st.markdown(f"{color_box} **{c['Dokter']}** — {c['Poli']} ({c['Kode']})")

# If any moved within columns, reconstruct df for that day
if moved_any:
    df_other = df[df["Hari"]!=selected_day].copy()
    new_rows=[]
    for slot in SLOTS:
        for c in new_state[slot]:
            new_rows.append({
                "Hari": selected_day,
                "Jam": slot,
                "Poli": c.get("Poli",""),
                "Dokter": c.get("Dokter","")
            })
    df_updated = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
    df = compute_status_codes(df_updated)
    # refresh kanban state
    st.session_state["kanban_state"][selected_day] = lanes_from_df_for_day(df, selected_day)
    st.success("Perubahan tersimpan secara sementara (session). Tekan 'Export' untuk mengunduh hasil.")

# Fallback Select & Move mode
st.markdown("---")
st.subheader("Select & Move (Fallback atau alat bantu)")
all_cards=[]
for s in SLOTS:
    for c in new_state[s]:
        card = c.copy(); card["Jam"]=s
        all_cards.append(card)

if len(all_cards)>0:
    sel_idx = st.selectbox("Pilih kartu yang ingin dipindah", options=list(range(len(all_cards))),
                           format_func=lambda i: f"{all_cards[i]['Dokter']} — {all_cards[i]['Poli']} @ {all_cards[i]['Jam']}")
    target_slot = st.selectbox("Pindahkan ke slot", SLOTS, index=0)
    if st.button("Pindahkan kartu ke slot tujuan"):
        card = all_cards[sel_idx]
        orig = card["Jam"]
        # remove original
        st.session_state["kanban_state"][selected_day][orig] = [c for c in st.session_state["kanban_state"][selected_day][orig] if not (c["Dokter"]==card["Dokter"] and c["Poli"]==card["Poli"])]
        # append to target
        st.session_state["kanban_state"][selected_day][target_slot].append({
            "id": f"{selected_day}|{target_slot}|{np.random.randint(1e9)}",
            "Dokter": card["Dokter"],
            "Poli": card["Poli"],
            "Kode": card.get("Kode","E"),
            "Status": ""
        })
        # rebuild df
        df_other = df[df["Hari"]!=selected_day].copy()
        new_rows=[]
        for s in SLOTS:
            for c in st.session_state["kanban_state"][selected_day][s]:
                new_rows.append({
                    "Hari": selected_day,
                    "Jam": s,
                    "Poli": c.get("Poli",""),
                    "Dokter": c.get("Dokter","")
                })
        df = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
        df = compute_status_codes(df)
        st.success("Kartu berhasil dipindahkan.")

# Show counts per slot with color indicator
st.markdown("---")
st.subheader("Kondisi Slot (Jumlah Dokter)")
cols_counts = st.columns(len(SLOTS))
for i,s in enumerate(SLOTS):
    with cols_counts[i]:
        cnt = len(st.session_state["kanban_state"][selected_day].get(s,[]))
        if cnt > 7:
            st.markdown(f"**{s}**  \n🔴 {cnt}")
        else:
            st.markdown(f"**{s}**  \n🟢 {cnt}")

# Export updated schedule
st.markdown("---")
st.subheader("Export / Download")

# build final df from session_state for all days
def build_full_df_from_state(original_df):
    # Start from original df but rebuild days present in kanban_state
    out_rows=[]
    days_in_state = list(st.session_state["kanban_state"].keys())
    # include days not edited
    for day in original_df["Hari"].unique():
        if day not in days_in_state:
            rows = original_df[original_df["Hari"]==day]
            out_rows.extend(rows[["Hari","Jam","Poli","Dokter"]].to_dict(orient="records"))
    # include days from state
    for day in days_in_state:
        for slot in st.session_state["kanban_state"][day]:
            for c in st.session_state["kanban_state"][day][slot]:
                out_rows.append({
                    "Hari": day,
                    "Jam": slot,
                    "Poli": c.get("Poli",""),
                    "Dokter": c.get("Dokter","")
                })
    return pd.DataFrame(out_rows)

df_final = build_full_df_from_state(df)

# recompute codes & status
df_final = compute_status_codes(df_final)

# download CSV
csv_bytes = df_final.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV Jadwal (updated)", csv_bytes, file_name=f"{download_all_name}.csv", mime="text/csv")

# download Excel (xlsx in-memory)
def to_excel_bytes(df_in):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Jadwal")
    return output.getvalue()

xlsx_bytes = to_excel_bytes(df_final)
st.download_button("Download Excel (xlsx) Jadwal (updated)", xlsx_bytes, file_name=f"{download_all_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.info("Tips: Untuk input, pastikan kolom di file Anda bernama: Hari, Range, Poli, Dokter (varian nama di-toleransi). Range boleh: '07.30-09.00', '7-9', '07:00-08:30', dll.")
