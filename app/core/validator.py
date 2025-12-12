import pandas as pd
import re
from io import BytesIO


class Validator:

    REQUIRED_COLS = ["HARI", "DOKTER", "POLI", "JAM"]

    @staticmethod
    def validate(file):
        """
        Validasi file excel upload user.
        Format yang didukung:
        - File berisi satu sheet
        - Atau sheet 'Reguler' & 'Poleks' (opsional)

        Output:
            (is_valid: bool, message: str)
        """

        try:
            if file is None:
                return False, "Tidak ada file yang diupload"

            # Streamlit uploader menghasilkan BytesIO → harus reset posisi pointer
            file.seek(0)

            df = pd.read_excel(file)

        except Exception as e:
            return False, f"Gagal membaca file: {e}"

        # Normalisasi kolom
        df.columns = [c.strip().upper() for c in df.columns]

        # Cek kolom
        missing = [c for c in Validator.REQUIRED_COLS if c not in df.columns]

        if missing:
            return False, f"Kolom wajib tidak ditemukan: {missing}"

        # Validasi nama hari
        valid_days = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
        df["HARI"] = df["HARI"].astype(str).str.title()

        invalid_days = df[~df["HARI"].isin(valid_days)]

        if len(invalid_days) > 0:
            return False, f"Ditemukan hari tidak valid: {invalid_days['HARI'].unique()}"

        # Validasi format jam
        pattern = re.compile(r"^\d{1,2}[:\.]\d{2}\s*-\s*\d{1,2}[:\.]\d{2}$")

        bad_jam = df[~df["JAM"].astype(str).str.match(pattern)]

        if len(bad_jam) > 0:
            return False, f"Format jam invalid pada {len(bad_jam)} baris"

        return True, None


    @staticmethod
    def validate_df(df):
        """
        Validasi dataframe internal setelah cleaner.
        """

        missing = [c for c in Validator.REQUIRED_COLS if c not in df.columns]
        if missing:
            return False, f"Kolom hilang pada DF internal: {missing}"

        if df["HARI"].isna().any():
            return False, "Terdapat nilai HARI kosong"

        if df["DOKTER"].isna().any():
            return False, "Terdapat nilai DOKTER kosong"

        return True, None
