import pandas as pd
import re
from io import BytesIO


class Validator:

    # Kolom yang diharapkan untuk format file Excel Anda
    REQUIRED_COLS_UPLOAD = ["Nama Dokter", "Poli Asal", "Jenis Poli"]
    
    # Hari yang valid
    VALID_DAYS = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at", "Jumat", "Sabtu"]

    @staticmethod
    def validate_excel_file(file):
        """
        Validasi file excel upload user dengan format sheet Reguler dan Poleks.
        
        Args:
            file: Uploaded file dari Streamlit
            
        Returns:
            (is_valid: bool, message: str)
        """
        try:
            if file is None:
                return False, "Tidak ada file yang diupload"

            # Streamlit uploader menghasilkan BytesIO → reset posisi
            file.seek(0)

            # Baca sheet names untuk validasi
            try:
                excel_file = pd.ExcelFile(file)
                sheet_names = excel_file.sheet_names
                
                # Cek sheet required
                required_sheets = ["Reguler", "Poleks"]
                missing_sheets = [s for s in required_sheets if s not in sheet_names]
                
                if missing_sheets:
                    available_sheets = ", ".join(sheet_names)
                    return False, f"Sheet wajib tidak ditemukan: {missing_sheets}. Sheet yang ada: {available_sheets}"
                
                # Validasi setiap sheet
                for sheet_name in required_sheets:
                    df = pd.read_excel(excel_file, sheet_name=sheet_name)
                    
                    # Validasi kolom required
                    df.columns = [str(c).strip() for c in df.columns]
                    missing = [c for c in Validator.REQUIRED_COLS_UPLOAD 
                              if c not in df.columns]
                    
                    if missing:
                        return False, f"Sheet '{sheet_name}': Kolom wajib tidak ditemukan: {missing}"
                    
                    # Validasi minimal 1 baris data
                    if df.empty:
                        return False, f"Sheet '{sheet_name}' kosong"
                        
            except Exception as e:
                # Fallback: coba baca sebagai single sheet
                file.seek(0)
                df = pd.read_excel(file)
                
                # Validasi kolom untuk single sheet
                df.columns = [str(c).strip() for c in df.columns]
                missing = [c for c in Validator.REQUIRED_COLS_UPLOAD 
                          if c not in df.columns]
                
                if missing:
                    return False, f"Kolom wajib tidak ditemukan: {missing}. Format file tidak sesuai."
                
                # Cek kolom hari
                hari_cols = [col for col in df.columns if col in Validator.VALID_DAYS]
                if not hari_cols:
                    return False, f"Tidak ditemukan kolom hari. Kolom yang ada: {list(df.columns)}"

            return True, "File valid"

        except Exception as e:
            return False, f"Gagal membaca file: {str(e)}"

    @staticmethod
    def validate_time_format(time_str):
        """
        Validasi format waktu.
        Format yang diterima: "07.30-10.00", "07:30-10:00", "7.30 - 10.00", dll.
        """
        if pd.isna(time_str) or str(time_str).strip() == "":
            return True  # Kosong dianggap valid
        
        time_str = str(time_str).strip()
        
        # Pattern untuk format: digit[:.]digit - digit[:.]digit
        pattern = re.compile(r'^\s*\d{1,2}[:\.]\d{2}\s*[-–]\s*\d{1,2}[:\.]\d{2}\s*$')
        
        return bool(pattern.match(time_str))

    @staticmethod
    def validate_dataframe(df):
        """
        Validasi dataframe internal setelah diproses oleh cleaner.
        """
        # Kolom required setelah cleaning
        required_cols = ["Nama Dokter", "Poli Asal", "Jenis Poli"]
        
        # Cek kolom
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            return False, f"Kolom hilang pada DF internal: {missing}"
        
        # Cek data kosong
        if df.empty:
            return False, "Dataframe kosong"
        
        # Cek missing values pada kolom required
        for col in required_cols:
            if df[col].isna().any():
                return False, f"Terdapat nilai kosong pada kolom {col}"
        
        # Validasi format waktu di kolom hari
        hari_cols = [col for col in df.columns if col in Validator.VALID_DAYS]
        
        bad_time_formats = []
        for hari in hari_cols:
            if hari in df.columns:
                for idx, value in df[hari].items():
                    if not Validator.validate_time_format(value):
                        bad_time_formats.append(f"Baris {idx+2}, {hari}: '{value}'")
        
        if bad_time_formats:
            return False, f"Format waktu tidak valid:\n" + "\n".join(bad_time_formats[:10])
        
        return True, "Dataframe valid"

    @staticmethod
    def validate_grid_data(df_grid, slot_strings):
        """
        Validasi dataframe grid hasil scheduler.
        """
        required_cols = ["POLI", "JENIS", "HARI", "DOKTER"]
        
        # Cek kolom
        missing = [c for c in required_cols if c not in df_grid.columns]
        if missing:
            return False, f"Kolom hilang pada grid: {missing}"
        
        # Cek slot columns
        missing_slots = [slot for slot in slot_strings if slot not in df_grid.columns]
        if missing_slots:
            return False, f"Slot waktu hilang pada grid: {missing_slots[:5]}..."
        
        # Cek data valid di slot
        valid_codes = ['', 'R', 'E']
        invalid_codes = []
        
        for slot in slot_strings:
            if slot in df_grid.columns:
                invalid = df_grid[~df_grid[slot].isin(valid_codes)]
                if not invalid.empty:
                    invalid_codes.append(f"Slot {slot}: nilai tidak valid ditemukan")
        
        if invalid_codes:
            return False, f"Nilai tidak valid pada slot: {invalid_codes}"
        
        return True, "Grid data valid"

    @staticmethod
    def get_time_format_errors(df):
        """
        Dapatkan daftar error format waktu untuk ditampilkan ke user.
        """
        errors = []
        hari_cols = [col for col in df.columns if col in Validator.VALID_DAYS]
        
        for hari in hari_cols:
            if hari in df.columns:
                for idx, value in df[hari].items():
                    if pd.notna(value) and not Validator.validate_time_format(value):
                        errors.append({
                            'row': idx + 2,  # +2 karena header + 1-based index
                            'hari': hari,
                            'value': str(value),
                            'doctor': df.loc[idx, 'Nama Dokter'] if 'Nama Dokter' in df.columns else 'Unknown',
                            'poli': df.loc[idx, 'Poli Asal'] if 'Poli Asal' in df.columns else 'Unknown'
                        })
        
        return errors
