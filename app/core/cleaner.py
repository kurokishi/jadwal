# app/core/cleaner.py
import pandas as pd
import io
from typing import Union

class DataCleaner:
    def clean(self, df_or_file) -> pd.DataFrame:
        """
        Clean data dari berbagai sumber
        
        Args:
            df_or_file: Bisa berupa:
                      1. BytesIO (file upload Streamlit)
                      2. String path ke file
                      3. DataFrame
            
        Returns:
            DataFrame yang sudah dibersihkan
        """
        print(f"🔧 DataCleaner.clean() called with type: {type(df_or_file)}")
        
        try:
            # Case 1: BytesIO (file upload Streamlit)
            if isinstance(df_or_file, io.BytesIO):
                print("   Input is BytesIO, reading Excel file...")
                return self._clean_from_bytesio(df_or_file)
            
            # Case 2: String path
            elif isinstance(df_or_file, str):
                print(f"   Input is string path: {df_or_file}")
                if df_or_file.endswith(('.xlsx', '.xls')):
                    return self._clean_from_excel(df_or_file)
                else:
                    raise ValueError(f"File format not supported: {df_or_file}")
            
            # Case 3: DataFrame
            elif isinstance(df_or_file, pd.DataFrame):
                print("   Input is DataFrame, cleaning directly...")
                return self._clean_dataframe(df_or_file)
            
            # Case 4: Bytes (raw bytes)
            elif isinstance(df_or_file, bytes):
                print("   Input is bytes, converting to BytesIO...")
                return self._clean_from_bytesio(io.BytesIO(df_or_file))
            
            else:
                raise ValueError(f"Unsupported input type: {type(df_or_file)}")
                
        except Exception as e:
            print(f"❌ Error in DataCleaner.clean(): {e}")
            raise
    
    def _clean_from_bytesio(self, bytes_io: io.BytesIO) -> pd.DataFrame:
        """Clean data dari BytesIO"""
        try:
            # Reset stream position
            bytes_io.seek(0)
            
            # Baca sheet names
            excel_file = pd.ExcelFile(bytes_io)
            sheet_names = excel_file.sheet_names
            print(f"   Excel sheets found: {sheet_names}")
            
            # Gabungkan sheet Reguler dan Poleks
            all_data = []
            
            for sheet in ['Reguler', 'Poleks']:
                if sheet in sheet_names:
                    print(f"   Reading sheet: {sheet}")
                    df = pd.read_excel(excel_file, sheet_name=sheet)
                    df['Jenis Poli'] = sheet  # Tambahkan kolom jenis
                    all_data.append(df)
                else:
                    print(f"   ⚠️ Sheet '{sheet}' not found")
            
            if not all_data:
                # Fallback: baca sheet pertama
                print("   No Reguler/Poleks sheets, reading first sheet...")
                df = pd.read_excel(excel_file, sheet_name=0)
                all_data.append(df)
            
            # Gabungkan semua data
            combined_df = pd.concat(all_data, ignore_index=True)
            print(f"   Combined data: {len(combined_df)} rows")
            
            # Clean dataframe
            return self._clean_dataframe(combined_df)
            
        except Exception as e:
            print(f"❌ Error reading from BytesIO: {e}")
            raise
    
    def _clean_from_excel(self, file_path: str) -> pd.DataFrame:
        """Clean data dari file Excel"""
        try:
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            print(f"   Excel sheets: {sheet_names}")
            
            all_data = []
            
            for sheet in ['Reguler', 'Poleks']:
                if sheet in sheet_names:
                    df = pd.read_excel(excel_file, sheet_name=sheet)
                    df['Jenis Poli'] = sheet
                    all_data.append(df)
            
            if not all_data:
                df = pd.read_excel(excel_file, sheet_name=0)
                all_data.append(df)
            
            combined_df = pd.concat(all_data, ignore_index=True)
            return self._clean_dataframe(combined_df)
            
        except Exception as e:
            print(f"❌ Error reading Excel file: {e}")
            raise
    
    def _clean_dataframe(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean DataFrame yang sudah digabungkan"""
        print(f"   Cleaning DataFrame with {len(df)} rows, {len(df.columns)} columns")
        print(f"   Original columns: {list(df.columns)}")
        
        # Buat copy
        df_clean = df.copy()
        
        # 1. Normalisasi nama kolom
        df_clean.columns = [str(col).strip() for col in df_clean.columns]
        
        # 2. Mapping nama kolom
        column_mapping = {
            'Nama Dokter': 'Nama Dokter',
            'nama dokter': 'Nama Dokter',
            'Dokter': 'Nama Dokter',
            'Poli Asal': 'Poli Asal',
            'poli asal': 'Poli Asal',
            'Poli': 'Poli Asal',
            'Jenis Poli': 'Jenis Poli',
            'jenis poli': 'Jenis Poli',
            'Jenis': 'Jenis Poli',
            'Senin': 'Senin',
            'Selasa': 'Selasa',
            'Rabu': 'Rabu',
            'Kamis': 'Kamis',
            "Jum'at": "Jum'at",
            'Jumat': "Jum'at",
            'Sabtu': 'Sabtu'
        }
        
        # Rename kolom
        for old_name, new_name in column_mapping.items():
            if old_name in df_clean.columns and new_name not in df_clean.columns:
                df_clean.rename(columns={old_name: new_name}, inplace=True)
                print(f"   Renamed column: {old_name} -> {new_name}")
        
        # 3. Drop kolom yang tidak perlu
        cols_to_drop = ['No', 'Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3']
        for col in cols_to_drop:
            if col in df_clean.columns:
                df_clean.drop(columns=[col], inplace=True)
                print(f"   Dropped column: {col}")
        
        # 4. Validasi kolom required
        required_cols = ['Nama Dokter', 'Poli Asal']
        missing_cols = [col for col in required_cols if col not in df_clean.columns]
        
        if missing_cols:
            raise ValueError(f"Missing required columns: {missing_cols}. Available: {list(df_clean.columns)}")
        
        # 5. Clean data per kolom
        # Nama Dokter
        if 'Nama Dokter' in df_clean.columns:
            df_clean['Nama Dokter'] = df_clean['Nama Dokter'].astype(str).str.strip()
            # Remove empty doctor names
            df_clean = df_clean[df_clean['Nama Dokter'] != 'nan']
        
        # Poli Asal
        if 'Poli Asal' in df_clean.columns:
            df_clean['Poli Asal'] = df_clean['Poli Asal'].astype(str).str.strip()
        
        # Jenis Poli (jika ada)
        if 'Jenis Poli' in df_clean.columns:
            df_clean['Jenis Poli'] = df_clean['Jenis Poli'].astype(str).str.strip()
            # Normalisasi nilai
            df_clean['Jenis Poli'] = df_clean['Jenis Poli'].replace({
                'reguler': 'Reguler',
                'poleks': 'Poleks',
                'Polek': 'Poleks',
                'REGULER': 'Reguler',
                'POLEKS': 'Poleks'
            })
        else:
            # Jika tidak ada kolom Jenis Poli, tambahkan default
            df_clean['Jenis Poli'] = 'Reguler'
        
        # 6. Clean kolom hari
        hari_cols = ['Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu']
        
        for hari in hari_cols:
            if hari in df_clean.columns:
                # Konversi ke string dan clean
                df_clean[hari] = df_clean[hari].astype(str).str.strip()
                # Replace nilai kosong dengan None
                df_clean[hari] = df_clean[hari].replace(['nan', 'NaN', 'NaT', 'None', ''], None)
                print(f"   Cleaned column {hari}: {df_clean[hari].notna().sum()} non-empty values")
        
        # 7. Drop baris yang semua kolom harinya kosong
        hari_cols_exist = [h for h in hari_cols if h in df_clean.columns]
        if hari_cols_exist:
            initial_count = len(df_clean)
            mask = df_clean[hari_cols_exist].isnull().all(axis=1)
            df_clean = df_clean[~mask].reset_index(drop=True)
            print(f"   Removed {initial_count - len(df_clean)} rows with all empty days")
        
        # 8. Reset index
        df_clean = df_clean.reset_index(drop=True)
        
        print(f"   ✅ Cleaning complete: {len(df_clean)} rows remaining")
        print(f"   Final columns: {list(df_clean.columns)}")
        
        return df_clean
    
    def get_available_sheets(self, file_path):
        """
        Dapatkan list sheet yang tersedia di file Excel
        """
        try:
            xls = pd.ExcelFile(file_path)
            return xls.sheet_names
        except:
            return []
