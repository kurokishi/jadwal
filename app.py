# app.py
import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import io

# try import sortables (drag & drop)
drag_available = True
try:
    from sortables import sort_table  # used in examples; should support cross-list in newer versions
except Exception as e:
    drag_available = False
    # we'll fallback to select-and-move mode

st.set_page_config(page_title="Kanban Jadwal Dokter", layout="wide")

# -------------------------
# Helpers: time slots
# -------------------------
def make_timeslots(start="07:00", end="14:30", interval_minutes=30):
    fmt = "%H:%M"
    s = datetime.strptime(start, fmt)
    e = datetime.strptime(end, fmt)
    slots = []
    cur = s
    while cur <= e:
        slots.append(cur.strftime(fmt))
        cur += timedelta(minutes=interval_minutes)
    return slots

SLOTS = make_timeslots("07:00", "14:30", 30)

# -------------------------
# Load / create data
# -------------------------
st.title("🗂️ Kanban Jadwal Dokter (07:00–14:30, 30m interval)")

uploaded = st.file_uploader("Upload file Excel (sheet: Jadwal) atau CSV", type=["xlsx","csv"])

@st.cache_data
def load_df_from_bytes(bytes_io, fname):
    try:
        if fname.endswith(".csv"):
            return pd.read_csv(io.BytesIO(bytes_io))
        else:
            # expect sheet "Jadwal" with columns: Hari, Dokter, Poli, Jam, Kode (optional)
            xls = pd.ExcelFile(io.BytesIO(bytes_io))
            if "Jadwal" in xls.sheet_names:
                df = xls.parse("Jadwal")
            else:
                # try first sheet
                df = xls.parse(xls.sheet_names[0])
            return df
    except Exception as ex:
        st.error(f"Gagal membaca file: {ex}")
        return pd.DataFrame(columns=["Hari","Dokter","Poli","Jam","Kode"])

if uploaded:
    bytes_data = uploaded.getvalue()
    df = load_df_from_bytes(bytes_data, uploaded.name)
else:
    # empty sample
    df = pd.DataFrame(columns=["Hari","Dokter","Poli","Jam","Kode"])

# Normalize column names
df.columns = [c.strip() for c in df.columns]

# Ensure required columns exist
for c in ["Hari","Dokter","Poli","Jam"]:
    if c not in df.columns:
        df[c] = ""

# Fill Kode if missing
if "Kode" not in df.columns:
    df["Kode"] = df["Poli"].str.contains("reguler", case=False).map({True:"R", False:"E"})

# -------------------------
# Compute status: Count & Bentrok
# -------------------------
def compute_status(df):
    out = df.copy()
    out["Jam"] = out["Jam"].astype(str)
    out["Count"] = out.groupby(["Hari","Jam"])["Dokter"].transform("count")
    out["Status"] = ""
    out.loc[out["Count"] > 7, "Status"] = "Over Kuota"
    # Bentrok when same Hari+Jam and same Dokter duplicated
    dup_mask = out.duplicated(subset=["Hari","Jam","Dokter"], keep=False)
    out.loc[dup_mask, "Status"] = "Bentrok"
    # If multiple different doctors same slot -> mark E cards as Bentrok (per your rule)
    # Find slots where more than one unique doctor present
    grouped = out.groupby(["Hari","Jam"])["Dokter"].nunique().reset_index(name="nunique")
    clashes = grouped[(grouped["nunique"] > 1)]
    if not clashes.empty:
        for _, r in clashes.iterrows():
            h = r["Hari"]; j = r["Jam"]
            mask = (out["Hari"]==h) & (out["Jam"]==j) & (out["Kode"]=="E")
            out.loc[mask, "Status"] = "Bentrok"
    return out

df = compute_status(df)

# -------------------------
# UI: select Hari
# -------------------------
hari_list = sorted([h for h in df["Hari"].unique() if str(h).strip() != ""])
selected_day = st.selectbox("Pilih Hari untuk edit Kanban", ["--Pilih Hari--"] + hari_list)

# session store for current kanban grid
if "kanban_state" not in st.session_state:
    st.session_state["kanban_state"] = {}

# when day changes, initialize kanban lanes from df
def init_kanban_for_day(day):
    lanes = {}
    for s in SLOTS:
        # get rows for this day+slot
        rows = df[(df["Hari"]==day) & (df["Jam"]==s)].copy()
        # convert each row into a dict card
        cards = []
        for idx, r in rows.iterrows():
            cards.append({
                "id": f"{day}|{s}|{idx}",
                "Dokter": r["Dokter"],
                "Poli": r["Poli"],
                "Kode": r.get("Kode","E"),
                "Status": r.get("Status","")
            })
        lanes[s] = cards
    return lanes

if selected_day != "--Pilih Hari--":
    if selected_day not in st.session_state["kanban_state"]:
        st.session_state["kanban_state"][selected_day] = init_kanban_for_day(selected_day)

    st.subheader(f"Kanban: {selected_day}")
    st.markdown("Drag kartu antar slot untuk memindahkan jam. Jika drag tidak tersedia di environment, gunakan mode **Select & Move** di bawah.")

    # -------------------------
    # Render Kanban columns
    # -------------------------
    cols = st.columns(len(SLOTS))
    new_state_for_day = {s: list(st.session_state["kanban_state"][selected_day].get(s, [])) for s in SLOTS}

    if drag_available:
        # Best-effort attempt: render each slot as a sortable table.
        # sort_table returns (sorted_df, moved_flag) in some implementations.
        # We'll call sort_table on each column individually and capture changes.
        moved_any = False
        for i, s in enumerate(SLOTS):
            with cols[i]:
                st.markdown(f"**{s}**")
                cards = new_state_for_day[s]
                # Build small dataframe to feed the sorter
                if len(cards) == 0:
                    st.info("—")
                else:
                    # convert to dataframe
                    tab = pd.DataFrame(cards)
                    # show colored compact card list via sort_table
                    try:
                        sorted_tab, moved = sort_table(tab, key="id", height="300px")
                    except Exception:
                        # fallback: show as table
                        st.dataframe(tab[["Dokter","Poli","Kode","Status"]], use_container_width=True)
                        sorted_tab = tab
                        moved = False

                    # convert back to list of dicts
                    new_cards = sorted_tab.to_dict(orient="records")
                    new_state_for_day[s] = new_cards
                    if moved:
                        moved_any = True
        # If any move happened we need to update underlying df accordingly
        if moved_any:
            # Reconstruct df rows for this day from new_state_for_day
            # We'll remove existing rows for day, then append updated ones
            df_other = df[df["Hari"] != selected_day].copy()
            new_rows = []
            for s in SLOTS:
                cards = new_state_for_day[s]
                for c in cards:
                    new_rows.append({
                        "Hari": selected_day,
                        "Dokter": c.get("Dokter",""),
                        "Poli": c.get("Poli",""),
                        "Jam": s,
                        "Kode": c.get("Kode","E")
                    })
            df = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
            df = compute_status(df)
            st.session_state["kanban_state"][selected_day] = init_kanban_for_day(selected_day)
            st.success("Perubahan disimpan ke jadwal.")
    else:
        # Fallback: Select & Move mode
        st.warning("Drag & drop tidak tersedia. Mode fallback: klik kartu -> pilih slot tujuan.")

        # Show all cards and allow selection
        all_cards = []
        for s in SLOTS:
            for c in new_state_for_day[s]:
                card = c.copy(); card["Jam"] = s
                all_cards.append(card)
        if len(all_cards) == 0:
            st.info("Tidak ada kartu di hari ini.")
        else:
            sel_idx = st.selectbox("Pilih Dokter yang ingin dipindah", options=list(range(len(all_cards))),
                                   format_func=lambda i: f"{all_cards[i]['Dokter']} — {all_cards[i]['Poli']} @ {all_cards[i]['Jam']}")
            target_slot = st.selectbox("Pindah ke slot", SLOTS)
            if st.button("Pindahkan"):
                card = all_cards[sel_idx]
                # remove original
                orig = card["Jam"]
                st.session_state["kanban_state"][selected_day][orig] = [c for c in st.session_state["kanban_state"][selected_day][orig] if not (c["Dokter"]==card["Dokter"] and c["Poli"]==card["Poli"])]
                # append to target
                st.session_state["kanban_state"][selected_day][target_slot].append({
                    "id": f"{selected_day}|{target_slot}|{np.random.randint(1e9)}",
                    "Dokter": card["Dokter"],
                    "Poli": card["Poli"],
                    "Kode": card.get("Kode","E"),
                    "Status": ""
                })
                # reconstruct df accordingly
                df_other = df[df["Hari"] != selected_day].copy()
                new_rows = []
                for s in SLOTS:
                    cards = st.session_state["kanban_state"][selected_day][s]
                    for c in cards:
                        new_rows.append({
                            "Hari": selected_day,
                            "Dokter": c.get("Dokter",""),
                            "Poli": c.get("Poli",""),
                            "Jam": s,
                            "Kode": c.get("Kode","E")
                        })
                df = pd.concat([df_other, pd.DataFrame(new_rows)], ignore_index=True)
                df = compute_status(df)
                st.success("Dokter berhasil dipindah.")

    # show compact summary counts per slot
    st.markdown("---")
    counts = {s: len(new_state_for_day[s]) for s in SLOTS}
    cols_counts = st.columns(len(SLOTS))
    for i, s in enumerate(SLOTS):
        with cols_counts[i]:
            c = counts[s]
            if c > 7:
                st.markdown(f"**{s}**  \n🔴 {c}")
            else:
                st.markdown(f"**{s}**  \n🟢 {c}")

    # Save / Export controls
    st.markdown("---")
    if st.button("Simpan perubahan ke file (download CSV)"):
        # reconstruct full df from session state for all days
        df_other_days = df[df["Hari"] != selected_day].copy()
        new_rows = []
        for s in SLOTS:
            cards = new_state_for_day[s]
            for c in cards:
                new_rows.append({
                    "Hari": selected_day,
                    "Dokter": c.get("Dokter",""),
                    "Poli": c.get("Poli",""),
                    "Jam": s,
                    "Kode": c.get("Kode","E")
                })
        df_out = pd.concat([df_other_days, pd.DataFrame(new_rows)], ignore_index=True)
        df_out = compute_status(df_out)
        csv = df_out.to_csv(index=False).encode("utf-8")
        st.download_button("Download CSV Jadwal (updated)", csv, file_name=f"jadwal_{selected_day}.csv", mime="text/csv")
        st.success("File siap di-download.")

else:
    st.info("Pilih hari untuk mulai mengedit jadwal di Kanban.")
