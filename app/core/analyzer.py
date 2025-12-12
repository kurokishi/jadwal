import pandas as pd
import re


class ErrorAnalyzer:

    def analyze_sheet(self, df, hari_list):
        report = {
            "is_valid": True,
            "errors": [],
            "warnings": [],
            "total_rows": len(df)
        }

        required = ["Nama Dokter", "Poli Asal", "Jenis Poli"]
        miss = [c for c in required if c not in df.columns]

        if miss:
            report["is_valid"] = False
            report["errors"].append(f"Missing columns: {miss}")

        pattern = re.compile(r"\d{1,2}[:\.]\d{2}\s*-\s*\d{1,2}[:\.]\d{2}")
        bad = 0
        for h in hari_list:
            if h in df.columns:
                for v in df[h].dropna().astype(str):
                    if not pattern.search(v):
                        bad += 1

        if bad:
            report["warnings"].append(f"{bad} invalid time format")

        return report

    def format_report(self, r):
        out = f"Valid: {r['is_valid']}\nRows: {r['total_rows']}\n"
        if r["errors"]:
            out += "Errors:\n" + "\n".join(r["errors"]) + "\n"
        if r["warnings"]:
            out += "Warnings:\n" + "\n".join(r["warnings"]) + "\n"
        return out
