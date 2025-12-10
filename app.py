import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

st.set_page_config(page_title="Jadwal Poli Final", layout="wide")

st.title("📅 Sistem Jadwal Poli – Final Version")


# ======================================================================================
# 1. NORMALISASI TOKENS WAKTU
# ======================================================================================

def _normalize_time_token(token: str) -> str:
    """Normalize token into HH:MM or return empty string if cannot."""
    if token is None:
        return ""
    t = str(token).strip()

    if t == "" or t.lower() in ["nan", "none"]:
        return ""

    t = t.replace(".", ":").replace("–", "-").replace("—", "-")
    t = t.lower().replace("am", "").replace("pm", "").strip()

    # Case: Only hour "7" → "07:00"
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
            mm_int = int(mm)
            if mm_int < 0 or mm_int > 59:
                return ""
        except:
            return ""
        return f"{hh}:{mm}"

    return ""


# ======================================================================================
# 2. PARSER RANGE WAKTU
# ======================================================================================

def expand_range_safe(range_str: str, interval_minutes: int = 30):
    """Safely expand time ranges."""
    if not isinstance(range_str, str) or range_str.strip() == "":
        return []

    text = range_str.replace(" ", "")

    # Coba pecah beberapa separator
    for sep in ["-", "–", "—", "to"]:
        if sep in text:
            parts = text.split(sep)
            break
    else:
        # no separator = single token
        tok = _normalize_time_token(text)
        return [tok] if tok else []

    if len(parts) < 2:
        return []

    start_tok = _normalize_time_token(parts[0])
    end_tok = _normalize_time_token(parts[1])

    if not start_tok or not end_tok:
        return []

    try:
        sdt = datetime.strptime(start_tok, "%H:%M")
        edt = datetime.strptime(end_tok, "%H:%M")
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


# ======================================================================================
# 3. UPLOAD FILE
# ======================================================================================

uploaded = st.file_uploader("Upload Excel (Template Jadwal)", type=["xlsx"])

if uploaded is None:
    st.info("Silakan upload file Excel dengan kolom: Hari | Range | Poli | Jenis | Dokter")
    st.stop()

raw = pd.read_excel(uploaded)
raw.columns = [c.strip() for c in raw.columns]


# ======================================================================================
# 4. NORMALISASI DATA
# ======================================================================================

# Normalisasi Jenis (case-insensitive)
raw["Jenis"] = raw["Jenis"].astype(str).str.strip().str.lower()
raw["Jenis"] = raw["Jenis"].replace({
    "reguler": "Reguler",
    "regular": "Reguler",
    "eksekutif": "Eksekutif",
    "executive": "Eksekutif",
    "poleks": "Eksekutif"
})

rows = []

for _, r in raw.iterrows():
    hari = str(r["Hari"]).strip()
    range_raw = str(r["Range"])
    poli = str(r["Poli"]).strip()
    jenis = str(r["Jenis"]).strip()
    dokter = str(r["Dokter"]).strip()

    slots = expand_range_safe(range_raw)

    if len(slots) == 0:
        tok = _normalize_time_token(range_raw)
        if tok:
            slots = [tok]

    for s in slots:
        rows.append({
            "Hari": hari,
            "Jam": s,
            "Poli": poli,
            "Jenis": jenis,
            "Dokter": dokter
        })

df = pd.DataFrame(rows)

# Kode R / E
df["Kode"] = df["Jenis"].apply(lambda x: "R" if x == "Reguler" else "E")


# ======================================================================================
# 5. OVER-KUOTA (>7 DOKTER EKSEKUTIF DALAM SATU JAM)
# ======================================================================================

df["Over_Kuota"] = False

for hari in df["Hari"].unique():
    df_hari = df[df["Hari"] == hari]

    for jam in df_hari["Jam"].unique():
        df_slot = df_hari[df_hari["Jam"] == jam]

        eksek = df_slot[df_slot["Jenis"] == "Eksekutif"]

        if len(eksek) > 7:
            df.loc[(df["Hari"] == hari) & (df["Jam"] == jam) & (df["Jenis"] == "Eksekutif"),
                   "Over_Kuota"] = True


# ======================================================================================
# 6. CEK BENTROK (Dokter sama tampil di slot sama pada poli berbeda)
# ======================================================================================

df["Bentrok"] = False
grouped = df.groupby(["Hari", "Jam", "Dokter"]).size()

for (hari, jam, dokter), count in grouped.items():
    if count > 1:
        df.loc[(df["Hari"] == hari) & (df["Jam"] == jam) & (df["Dokter"] == dokter),
               "Bentrok"] = True


# ======================================================================================
# 7. WARNA TABEL
# ======================================================================================

def color_row(row):
    if row["Over_Kuota"]:
        return ["background-color: red; color: white"] * len(row)
    if row["Bentrok"]:
        return ["background-color: orange; color: black"] * len(row)
    if row["Kode"] == "R":
        return ["background-color: lightgreen"] * len(row)
    return ["background-color: lightblue"] * len(row)


st.subheader("📋 Jadwal Final (Interval 30 menit)")

st.dataframe(
    df.style.apply(color_row, axis=1),
    use_container_width=True
)


# ======================================================================================
# 8. DASHBOARD
# ======================================================================================

st.subheader("📊 Dashboard – Jumlah Dokter per Jam/Hari")

pivot = df.pivot_table(
    index="Jam",
    columns="Hari",
    values="Dokter",
    aggfunc="count",
    fill_value=0
)

st.dataframe(pivot, use_container_width=True)


# ======================================================================================
# 9. DOWNLOAD
# ======================================================================================

st.download_button(
    label="📥 Download CSV",
    data=df.to_csv(index=False),
    file_name="jadwal_final.csv",
    mime="text/csv"
)
