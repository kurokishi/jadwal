import pandas as pd
import re


class DataCleaner:

    @staticmethod
    def clean(df, hari_list, jenis_poli, auto_fix=True):
        df = df.copy()

        required = ["Nama Dokter", "Poli Asal", "Jenis Poli"]
        for c in required:
            if c not in df.columns:
                df[c] = ""

        df["Jenis Poli"] = df["Jenis Poli"].fillna(jenis_poli)

        for h in hari_list:
            if h in df.columns and auto_fix:
                df[h] = df[h].apply(DataCleaner.fix_format)

        if hari_list:
            df = df[df[hari_list].notna().any(axis=1)]

        return df

    @staticmethod
    def fix_format(v):
        if pd.isna(v):
            return ""
        v = re.sub(r"[^0-9\.\:\-\s]", "", str(v)).replace(" ", "")
        p = v.split("-")
        if len(p) == 2:
            return p[0].replace(".", ":") + "-" + p[1].replace(".", ":")
        return v
