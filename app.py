# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

# Optional drag & drop package; fallback available
try:
    from sortables import sort_table
    DRAG_AVAILABLE = True
except Exception:
    DRAG_AVAILABLE = False

st.set_page_config(page_title="Jadwal Dokter (Poli + Jenis)", layout="wide")

# -----------------------
# Helpers: time parsing
# -----------------------
def normalize_time_token(token: str) -> str:
    if token is None:
        return ""
    t = str(token).strip()
    if t == "" or t.lower() in ["nan", "none"]:
        return ""
    t = t.replace(".", ":")
    t = t.lower().replace("am", "").replace("pm", "")
    t = t.strip()
    if ":" not in t:
        if len(t) in (1,2):
            t = t.zfill(2) + ":00"
        else:
            try:
                dt = datetime.strptime(t, "%H%M")
                t = dt.strftime("%H:%M")
            except Exception:
                return ""
    else:
        parts = t.split(":")
        hh = parts[0].zfill(2)
        mm = parts[1].zfill(2)
        t = f"{hh}:{mm}"
    return t

def expand_time_range(range_str: str, interval_minutes:int=30):
    if not isinstance(range_str, str) or range_str.strip() == "":
        return []
    seps = ["-", "–", "—", " to "]
    parts = None
    for sep in seps:
        if sep in range_str:
            parts = range_str.split(sep)
            break
    if parts is None:
        # single token
        tok = normalize_time_token(range_str)
        return [tok] if tok else []
    if len(parts) < 2:
        return []
    start = normalize_time_token(parts[0])
    end = normalize_time_token(parts[1])
    if not start or not end:
        return []
    try:
        sdt = datetime.strptime(start, "%H:%M")
        edt = datetime.strptime(end, "%H:%M")
    except Exception:
        return []
    if edt < sdt:
        return []
    slots=[]
    cur = sdt
    while cur <= edt:
        slots.append(cur.strftime("%H:%M"))
        cur += timedelta(minutes=interval_minutes)
    return slots

# -----------------------
# Read uploaded file
# -----------------------
st.title("Jadwal Dokter — Poli + Jenis (Reguler / Eksekutif)")

uploaded = st.file_uploader("Upload Excel (sheet) atau CSV — kolom: Hari, Range, Poli, Jenis, Dokter", type=["xlsx","csv"])
@st.cache_data
def load_raw(bytes_io, fname):
    try:
        if fname.lower().endswith(".csv"):
            df = pd.read_csv(io.BytesIO(bytes_io))
        else:
            xls = pd.ExcelFile(io.BytesIO(bytes_io))
            sheet = "Jadwal" if "Jadwal" in xls.sheet_names else xls.sheet_names[0]
            df = xls.parse(sheet)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return pd.DataFrame()

if not uploaded:
    st.info("Silakan upload file Excel/CSV yang berisi kolom: Hari, Range, Poli, Jenis, Dokter.")
    st.markdown("Contoh baris: `Senin | 07.30-09.00 | Anak | Reguler | dr. Budi`")
    st.stop()

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
    st.error("Kolom tidak lengkap. Pastikan file memiliki kolom: Hari, Range, Poli, Jenis, Dokter.")
    st.write("Terbaca:", cols)
    st.stop()

# -----------------------
# Expand ranges -> per-slot rows
# -----------------------
rows = []
for _, r in raw.iterrows():
    hari = str(r.get(Hari_col)).strip()
    range_raw = r.get(Range_col)
    poli = str(r.get(Poli_col)).strip()
    jenis = str(r.get(Jenis_col)).strip()
    dokter = str(r.get(Dokter_col)).strip()
    if not hari or pd.isna(range_raw) or not poli or not jenis or not dokter:
        continue
    slots = expand_time_range(str(range_raw), interval_minutes=30)
    if not slots:
        tok = normalize_time_token(str(range_raw))
        if tok:
            slots = [tok]
    for s in slots:
        rows.append({"Hari": hari, "Jam": s, "Poli": poli, "Jenis": jenis, "Dokter": dokter})

if len(rows) == 0:
    st.error("Tidak ada slot yang berhasil dibentuk dari Range. Periksa format Range.")
    st.stop()

df = pd.DataFrame(rows)

# -----------------------
# compute kode & status
# -----------------------
def map_kode(jenis):
    if isinstance(jenis, str) and "reguler" in jenis.lower():
        return "R"
    # A: Treat Eksekutif and others as E (including Poleks)
    return "E"

def compute_status(df):
    d = df.copy()
    d["Kode"] = d["Jenis"].apply(map_kode)
    d["Count"] = d.groupby(["Hari","Jam"])["Dokter"].transform("count")
    d["Status"] = ""
    d.loc[d["Count"] > 7, "Status"] = "Over Kuota"
    # If more than 1 unique doctor in slot -> E entries become Bentrok
    grouped = d.groupby(["Hari","Jam"])["Dokter"].nunique().reset_index(name="nunique")
    clashes = grouped[grouped["nunique"] > 1]
    if not clashes.empty:
        for _, r in clashes.iterrows():
            h = r["Hari"]; j = r["Jam"]
            mask = (d["Hari"]==h) & (d["Jam"]==j) & (d["Kode"]=="E")
            d.loc[mask, "Status"] = "Bentrok"
    return d

df = compute_status(df)

# -----------------------
# Left controls & summary
# -----------------------
st.sidebar.header("Kontrol")
selected_day = st.sidebar.selectbox("Pilih Hari (untuk Kanban editing)", sorted(df["Hari"].unique()))
export_name = st.sidebar.text_input("Nama file export (tanpa ekstensi)", "jadwal_updated")

st.markdown("## Ringkasan")
c1, c2, c3 = st.columns(3)
c1.metric("Total baris (slot)", len(df))
c2.metric("Total dokter unik", df["Dokter"].nunique())
c3.metric("Total slot unik", df[["Hari","Jam"]].drop_duplicates().shape[0])

# -----------------------
# Dashboard
# -----------------------
st.header("Dashboard")
summary = df.groupby(["Hari","Jam"]).size().reset_index(name="Jumlah")
pivot = summary.pivot(index="Jam", columns="Hari", values="Jumlah").fillna(0).sort_index()

import plotly.express as px
st.subheader("Heatmap (Jam × Hari)")
fig = px.imshow(pivot, labels=dict(x="Hari", y="Jam", color="Jumlah Dokter"), color_continuous_scale="Blues")
st.plotly_chart(fig, use_container_width=True, height=450)

st.subheader("Jumlah Dokter per Jam (line per Hari)")
fig2 = px.line(summary, x="Jam", y="Jumlah", color="Hari", markers=True)
st.plotly_chart(fig2, use_container_width=True, height=350)

# -----------------------
# Kanban editor
# -----------------------
st.header(f"Kanban Editor — {selected_day}")

# determine slots for the day
day_slots = sorted(df[df["Hari"]==selected_day]["Jam"].unique(), key=lambda x: datetime.strptime(x,"%H:%M"))
if len(day_slots)==0:
    # fallback default slots
    s = datetime.strptime("07:00","%H:%M")
    e = datetime.strptime("14:30","%H:%M")
    slots=[]
    cur=s
    while cur<=e:
        slots.append(cur.strftime("%H:%M"))
        cur+=timedelta(minutes=30)
    day_slots = slots

SLOTS = day_slots

def make_lanes(day):
    lanes={}
    for slot in SLOTS:
        rows = df[(df["Hari"]==day) & (df["Jam"]==slot)].reset_index(drop=True)
        cards=[]
        for i,r in rows.iterrows():
            cards.append({
                "id": f"{day}|{slot}|{i}|{np.random.randint(1e9)}",
                "Dokter": r["Dokter"],
                "Poli": r["Poli"],
                "Jenis": r["Jenis"],
                "Kode": r["Kode"],
                "Status": r["Status"]
            })
        lanes[slot]=cards
    return lanes

if "kanban_state" not in st.session_state:
    st.session_state["kanban_state"] = {}

if selected_day not in st.session_state["kanban_state"]:
    st.session_state["kanban_state"][selected_day] = make_lanes(selected_day)

st.write("Seret kartu antar kolom untuk memindahkan; jika drag tidak tersedia, gunakan Select & Move.")

cols = st.columns(len(SLOTS))
new_state = {s:list(st.session_state["kanban_state"][selected_day].get(s,[])) for s in SLOTS}
moved_any = False

if DRAG_AVAILABLE:
    for i,s in enumerate(SLOTS):
        with cols[i]:
            st.markdown(f"**{s}**")
            cards = new_state[s]
            if len(cards)==0:
                st.info("—")
            else:
                tab = pd.DataFrame(cards)
                try:
                    sorted_tab, moved = sort_table(tab, key="id", height="300px")
                    new_cards = sorted_tab.to_dict(orient="records")
                    new_state[s] = new_cards
                    if moved:
                        moved_any = True
                except Exception:
                    st.dataframe(tab[["Dokter","Poli","Jenis","Kode","Status"]], use_container_width=True)
else:
    st.warning("Drag & drop tidak tersedia. Mode fallback aktif.")
    for i,s in enumerate(SLOTS):
        with cols[i]:
            st.markdown(f"**{s}**")
            cards = new_state[s]
            if len(cards)==0:
                st.info("—")
            else:
                for c in cards:
                    badge = ""
                    if c["Status"] in ["Bentrok","Over Kuota"]:
                        badge = "🔴"
                    elif c["Kode"]=="R":
                        badge = "🟢"
                    else:
                        badge = "🔵"
                    st.markdown(f"{badge} **{c['Dokter']}** — {c['Poli']} ({c['Jenis']})")

if moved_any:
    # rebuild df for day from new_state
    df_other = df[df["Hari"]!=selected_day].copy()
    new_rows=[]
    for s in SLOTS:
        for c in new_state[s]:
            new_rows.append({"Hari": selected_day, "Jam": s, "Poli": c.get("Poli",""), "Jenis": c.get("Jenis",""), "Dokter": c.get("Dokter","")})
    df = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
    df = compute_status(df)
    st.session_state["kanban_state"][selected_day] = make_lanes(selected_day)
    st.success("Perubahan disimpan temporer (session). Tekan Export untuk unduh hasil.")

# Select & Move fallback
st.markdown("---")
st.subheader("Select & Move")
all_cards=[]
for s in SLOTS:
    for c in new_state[s]:
        card=c.copy(); card["Jam"]=s
        all_cards.append(card)

if len(all_cards)>0:
    sel_idx = st.selectbox("Pilih kartu", options=list(range(len(all_cards))),
                           format_func=lambda i: f"{all_cards[i]['Dokter']} — {all_cards[i]['Poli']} @ {all_cards[i]['Jam']}")
    target_slot = st.selectbox("Pindah ke slot", SLOTS, index=0)
    if st.button("Pindahkan"):
        card = all_cards[sel_idx]
        orig = card["Jam"]
        st.session_state["kanban_state"][selected_day][orig] = [c for c in st.session_state["kanban_state"][selected_day][orig] if not (c["Dokter"]==card["Dokter"] and c["Poli"]==card["Poli"])]
        st.session_state["kanban_state"][selected_day][target_slot].append({
            "id": f"{selected_day}|{target_slot}|{np.random.randint(1e9)}",
            "Dokter": card["Dokter"],
            "Poli": card["Poli"],
            "Jenis": card["Jenis"],
            "Kode": card.get("Kode","E"),
            "Status": ""
        })
        # rebuild df
        df_other = df[df["Hari"]!=selected_day].copy()
        new_rows=[]
        for s in SLOTS:
            for c in st.session_state["kanban_state"][selected_day][s]:
                new_rows.append({"Hari": selected_day, "Jam": s, "Poli": c.get("Poli",""), "Jenis": c.get("Jenis",""), "Dokter": c.get("Dokter","")})
        df = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
        df = compute_status(df)
        st.success("Kartu berhasil dipindah.")

# slot counts
st.markdown("---")
st.subheader("Kondisi Slot")
counts_cols = st.columns(len(SLOTS))
for i,s in enumerate(SLOTS):
    with counts_cols[i]:
        cnt = len(st.session_state["kanban_state"][selected_day].get(s,[]))
        if cnt > 7:
            st.markdown(f"**{s}**  \n🔴 {cnt}")
        else:
            st.markdown(f"**{s}**  \n🟢 {cnt}")

# export
st.markdown("---")
st.subheader("Export / Download")

def build_full_df(original_df):
    out=[]
    days_state = list(st.session_state["kanban_state"].keys())
    # add non-edited days
    for day in original_df["Hari"].unique():
        if day not in days_state:
            rows = original_df[original_df["Hari"]==day]
            out.extend(rows[["Hari","Jam","Poli","Jenis","Dokter"]].to_dict(orient="records"))
    # add edited days
    for day in days_state:
        for slot in st.session_state["kanban_state"][day]:
            for c in st.session_state["kanban_state"][day][slot]:
                out.append({"Hari": day, "Jam": slot, "Poli": c.get("Poli",""), "Jenis": c.get("Jenis",""), "Dokter": c.get("Dokter","")})
    return pd.DataFrame(out)

df_final = build_full_df(df)
df_final = compute_status(df_final)

csv = df_final.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV Jadwal (updated)", csv, file_name=f"{export_name}.csv", mime="text/csv")

def to_xlsx_bytes(df_in):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Jadwal")
    return out.getvalue()

xlsx = to_xlsx_bytes(df_final)
st.download_button("Download Excel (xlsx) Jadwal (updated)", xlsx, file_name=f"{export_name}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.info("Template input: Hari | Range | Poli | Jenis | Dokter. Jenis: Reguler (untuk kode R) atau Eksekutif (untuk kode E).")
