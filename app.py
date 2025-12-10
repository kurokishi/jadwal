# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

# Optional drag & drop library
try:
    from sortables import sort_table
    DRAG_AVAILABLE = True
except Exception:
    DRAG_AVAILABLE = False

st.set_page_config(page_title="Jadwal Poli (Streamlit Full)", layout="wide")
st.title("📅 Jadwal Poli — Streamlit Full (Offline)")

# ---------------------------
# Helpers: time normalization & expansion
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
# Session state: history (undo/redo), kanban_state
# ---------------------------
if "history" not in st.session_state:
    st.session_state.history = []
if "future" not in st.session_state:
    st.session_state.future = []
if "kanban_state" not in st.session_state:
    st.session_state.kanban_state = {}  # day -> lanes dict

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
uploaded = st.sidebar.file_uploader("Upload Excel (sheet) or CSV with columns: Hari, Range, Poli, Jenis, Dokter", type=["xlsx","csv"])
if st.sidebar.button("Download template example"):
    sample = pd.DataFrame({
        "Hari":["Senin","Senin","Selasa"],
        "Range":["07.30-09.00","09.00-11.00","07.00-08.30"],
        "Poli":["Anak","Anak","Gigi"],
        "Jenis":["Reguler","Eksekutif","Reguler"],
        "Dokter":["dr. Budi","dr. Sari","drg. Putri"]
    })
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
# Expand ranges -> slots
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
df["Jenis"] = df["Jenis"].astype(str).str.strip().replace({"reguler":"Reguler","regular":"Reguler","eksekutif":"Eksekutif","executive":"Eksekutif","poleks":"Eksekutif","POLEKS":"Eksekutif"})
# Kode
df["Kode"] = df["Jenis"].apply(lambda x: "R" if str(x).lower()=="reguler" else "E")

# push initial snapshot to history
push_history(df.copy())

# ---------------------------
# Compute Over-kuota only for Eksekutif/Poleks
# ---------------------------
def compute_status(df_in):
    d = df_in.copy()
    d["Over_Kuota"] = False
    d["Bentrok"] = False
    # over: count Eksekutif entries per (Hari,Jam)
    eksek = d[d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)]
    poleks_counts = eksek.groupby(["Hari","Jam"]).size()
    over_slots = poleks_counts[poleks_counts > 7].index if not poleks_counts.empty else []
    for (hari,jam) in over_slots:
        d.loc[(d["Hari"]==hari)&(d["Jam"]==jam)&(d["Jenis"].str.lower().str.contains("eksekutif|poleks", na=False)), "Over_Kuota"] = True
    # bentrok: same dokter assigned to multiple poli in same slot
    grouped = d.groupby(["Hari","Jam","Dokter"]).size()
    for (hari,jam,dok), cnt in grouped.items():
        if cnt > 1:
            d.loc[(d["Hari"]==hari)&(d["Jam"]==jam)&(d["Dokter"]==dok), "Bentrok"] = True
    return d

df = compute_status(df)

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
            st.success("Undo berhasil")
with colu2:
    if st.button("Redo"):
        nxt = redo()
        if nxt is not None:
            df = nxt.copy()
            st.success("Redo berhasil")

# ---------------------------
# Dashboard & summary
# ---------------------------
st.header("Ringkasan")
c1,c2,c3 = st.columns(3)
c1.metric("Total slot", len(df))
c2.metric("Total dokter unik", df["Dokter"].nunique())
c3.metric("Slot unik", df[["Hari","Jam"]].drop_duplicates().shape[0])

# Heatmap
st.subheader("Heatmap (Jam × Hari)")
summary = df.groupby(["Hari","Jam"]).size().reset_index(name="Jumlah")
pivot = summary.pivot(index="Jam", columns="Hari", values="Jumlah").fillna(0).sort_index()
import plotly.express as px
fig = px.imshow(pivot, labels=dict(x="Hari", y="Jam", color="Jumlah Dokter"), color_continuous_scale="Blues")
st.plotly_chart(fig, use_container_width=True, height=420)

# ---------------------------
# Schedule View (improved table)
# ---------------------------
st.subheader("Tabel Jadwal (dapat disaring)")
filtered = df[(df["Poli"].isin(poli_filter)) & (df["Jenis"].isin(jenis_filter))]
if selected_day != "--Semua--":
    filtered = filtered[filtered["Hari"]==selected_day]

def style_rows(row):
    if row["Over_Kuota"]:
        return ["background-color:#ffb3b3;color:black"]*len(row)
    if row["Bentrok"]:
        return ["background-color:#ffd9b3;color:black"]*len(row)
    if row["Kode"]=="R":
        return ["background-color:#dff2d8;color:black"]*len(row)
    # E default blue tint; Poleks/Eksek specifically blue
    return ["background-color:#e8f0ff;color:black"]*len(row)

st.dataframe(filtered.sort_values(["Hari","Jam","Poli","Dokter"]).reset_index(drop=True).style.apply(style_rows, axis=1), use_container_width=True)

# ---------------------------
# Kanban editor (per selected_day)
# ---------------------------
st.subheader("Kanban Editor")
if selected_day == "--Semua--":
    st.info("Pilih satu hari di sidebar untuk membuka Kanban editor.")
else:
    SLOTS = sorted(df[df["Hari"]==selected_day]["Jam"].unique(), key=lambda x: datetime.strptime(x,"%H:%M"))
    if len(SLOTS)==0:
        st.info("Tidak ada slot untuk hari ini.")
    else:
        # initialize kanban state if not exists
        if selected_day not in st.session_state.kanban_state:
            lanes = {}
            for s in SLOTS:
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
                        "Bentrok": bool(r["Bentrok"])
                    })
                lanes[s]=cards
            st.session_state.kanban_state[selected_day] = lanes

        lanes = st.session_state.kanban_state[selected_day]
        cols = st.columns(min(len(SLOTS), 8))
        col_objs = []
        for i,s in enumerate(SLOTS):
            col_objs.append(cols[i % len(cols)])
        new_state = {s: list(lanes.get(s,[])) for s in SLOTS}
        moved_any = False

        if DRAG_AVAILABLE:
            for i,s in enumerate(SLOTS):
                with col_objs[i % len(cols)]:
                    st.markdown(f"**{s}**")
                    cards = new_state[s]
                    if not cards:
                        st.info("—")
                    else:
                        tab = pd.DataFrame(cards)
                        try:
                            sorted_tab, moved = sort_table(tab, key="id", height="280px")
                            new_cards = sorted_tab.to_dict(orient="records")
                            new_state[s] = new_cards
                            if moved:
                                moved_any = True
                        except Exception:
                            st.dataframe(tab[["Dokter","Poli","Jenis","Kode","Over","Bentrok"]], use_container_width=True)
        else:
            st.warning("Drag & drop tidak tersedia. Gunakan 'Select & Move' di bawah.")
            for i,s in enumerate(SLOTS):
                with col_objs[i % len(cols)]:
                    st.markdown(f"**{s}**")
                    for c in new_state[s]:
                        badge = "🟢" if c["Kode"]=="R" else ("🔵" if not c["Over"] else "🔴")
                        st.markdown(f"{badge} **{c['Dokter']}** — {c['Poli']} ({c['Jenis']})")

        if moved_any:
            # reconstruct df for this day from new_state
            df_other = df[df["Hari"]!=selected_day].copy()
            new_rows = []
            for s in SLOTS:
                for c in new_state[s]:
                    new_rows.append({"Hari":selected_day, "Jam":s, "Poli":c.get("Poli",""), "Jenis":c.get("Jenis",""), "Dokter":c.get("Dokter","")})
            df_new = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
            df_new = compute_status(df_new)
            push_history(df_new.copy())
            df = df_new.copy()
            # refresh kanban_state
            st.session_state.kanban_state[selected_day] = {}
            for s in SLOTS:
                rows_s = df[(df["Hari"]==selected_day)&(df["Jam"]==s)].reset_index(drop=True)
                cards = [{"id": f"{selected_day}|{s}|{i}|{np.random.randint(1e9)}","Dokter":r["Dokter"],"Poli":r["Poli"],"Jenis":r["Jenis"],"Kode":r["Kode"],"Over":bool(r["Over_Kuota"]),"Bentrok":bool(r["Bentrok"])} for i,r in rows_s.iterrows()]
                st.session_state.kanban_state[selected_day][s]=cards
            st.success("Perubahan disimpan (session). Tekan Export untuk unduh.")

        # Select & Move fallback
        st.markdown("---")
        st.write("Select & Move:")
        all_cards=[]
        for s in SLOTS:
            for c in new_state[s]:
                cc = c.copy(); cc["Jam"]=s; all_cards.append(cc)
        if len(all_cards)>0:
            sel_idx = st.selectbox("Pilih kartu", options=list(range(len(all_cards))), format_func=lambda i: f"{all_cards[i]['Dokter']} — {all_cards[i]['Poli']} @ {all_cards[i]['Jam']}")
            target_slot = st.selectbox("Target slot", SLOTS, index=0)
            if st.button("Pindahkan kartu"):
                card = all_cards[sel_idx]
                orig = card["Jam"]
                st.session_state.kanban_state[selected_day][orig] = [c for c in st.session_state.kanban_state[selected_day][orig] if not (c["Dokter"]==card["Dokter"] and c["Poli"]==card["Poli"])]
                st.session_state.kanban_state[selected_day][target_slot].append({"id":f"{selected_day}|{target_slot}|{np.random.randint(1e9)}","Dokter":card["Dokter"],"Poli":card["Poli"],"Jenis":card["Jenis"],"Kode":card.get("Kode","E"),"Over":False,"Bentrok":False})
                # rebuild df and compute status
                df_other = df[df["Hari"]!=selected_day].copy()
                new_rows=[]
                for s in SLOTS:
                    for c in st.session_state.kanban_state[selected_day][s]:
                        new_rows.append({"Hari":selected_day,"Jam":s,"Poli":c.get("Poli",""),"Jenis":c.get("Jenis",""),"Dokter":c.get("Dokter","")})
                df_new = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
                df_new = compute_status(df_new)
                push_history(df_new.copy())
                df = df_new.copy()
                st.success("Kartu dipindah dan status diperbarui.")

# ---------------------------
# Export buttons
# ---------------------------
st.markdown("---")
st.header("Export & Simpan")
csv_bytes = df.to_csv(index=False).encode("utf-8")
st.download_button("Download CSV", csv_bytes, file_name="jadwal_updated.csv", mime="text/csv")
# xlsx
def to_xlsx_bytes(df_in):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df_in.to_excel(writer, index=False, sheet_name="Jadwal")
    return out.getvalue()
xlsx_bytes = to_xlsx_bytes(df)
st.download_button("Download XLSX", xlsx_bytes, file_name="jadwal_updated.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
