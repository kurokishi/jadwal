# app/core/cleaner.py
import pandas as pd

class DataCleaner:
    def clean(self, df_or_file_path):
        """
        Clean data dari file Excel atau DataFrame
        Bisa menerima:
        1. Path file Excel (akan membaca sheet Reguler dan Poleks)
        2. DataFrame tunggal
        """
        if isinstance(df_or_file_path, str) and df_or_file_path.endswith(('.xlsx', '.xls')):
            # Baca dari file Excel
            return self._clean_from_excel(df_or_file_path)
        elif isinstance(df_or_file_path, pd.DataFrame):
            # Sudah DataFrame
            return self._clean_dataframe(df_or_file_path)
        else:
            raise ValueError("Input harus file Excel (.xlsx/.xls) atau DataFrame")
    
    def _clean_from_excel(self, file_path):
        """
        Baca dan gabungkan sheet Reguler dan Poleks
        """
        try:
            # Baca sheet Reguler
            df_reguler = pd.read_excel(file_path, sheet_name='Reguler')
            df_reguler['Jenis Poli'] = 'Reguler'
            
            # Baca sheet Poleks
            df_poleks = pd.read_excel(file_path, sheet_name='Poleks')
            df_poleks['Jenis Poli'] = 'Poleks'
            
            # Gabungkan kedua sheet
            df_combined = pd.concat([df_reguler, df_poleks], ignore_index=True)
            
            # Clean dataframe gabungan
            return self._clean_dataframe(df_combined)
            
        except Exception as e:
            # Fallback: baca semua sheet
            try:
                xls = pd.ExcelFile(file_path)
                all_dfs = []
                
                for sheet_name in xls.sheet_names:
                    if sheet_name not in ['Poli Asal', 'Jadwal']:  # Skip sheet non-data
                        df = pd.read_excel(xls, sheet_name=sheet_name)
                        df['Jenis Poli'] = 'Reguler' if 'Reguler' in sheet_name else 'Poleks' if 'Poleks' in sheet_name else 'Unknown'
                        all_dfs.append(df)
                
                if all_dfs:
                    df_combined = pd.concat(all_dfs, ignore_index=True)
                    return self._clean_dataframe(df_combined)
                else:
                    # Coba baca sebagai single sheet
                    df = pd.read_excel(file_path)
                    return self._clean_dataframe(df)
                    
            except Exception as e2:
                raise ValueError(f"Gagal membaca file Excel: {str(e2)}")
    
    def _clean_dataframe(self, df):
        """
        Clean DataFrame yang sudah digabungkan
        """
        # Salin DataFrame
        df_clean = df.copy()
        
        # 1. Normalisasi nama kolom
        df_clean.columns = df_clean.columns.str.strip()
        
        # Mapping nama kolom
        column_mapping = {
            'Nama Dokter': 'Nama Dokter',
            'Poli Asal': 'Poli Asal',
            'Jenis Poli': 'Jenis Poli',
            'Senin': 'Senin',
            'Selasa': 'Selasa',
            'Rabu': 'Rabu',
            'Kamis': 'Kamis',
            "Jum'at": "Jum'at",
            'Jumat': "Jum'at",
            'Sabtu': 'Sabtu'
        }
        
        # Rename kolom yang sesuai
        for old, new in column_mapping.items():
            if old in df_clean.columns:
                df_clean.rename(columns={old: new}, inplace=True)
        
        # 2. Drop kolom yang tidak perlu
        cols_to_drop = ['No', 'Unnamed: 0', 'Unnamed: 1', 'Unnamed: 2', 'Unnamed: 3']
        for col in cols_to_drop:
            if col in df_clean.columns:
                df_clean.drop(columns=[col], inplace=True)
        
        # 3. Validasi kolom required
        required_cols = ['Nama Dokter', 'Poli Asal', 'Jenis Poli']
        missing_cols = [col for col in required_cols if col not in df_clean.columns]
        if missing_cols:
            raise ValueError(f"Kolom wajib tidak ditemukan: {missing_cols}")
        
        # 4. Clean data per kolom
        # Nama Dokter
        if 'Nama Dokter' in df_clean.columns:
            df_clean['Nama Dokter'] = df_clean['Nama Dokter'].astype(str).str.strip()
        
        # Poli Asal
        if 'Poli Asal' in df_clean.columns:
            df_clean['Poli Asal'] = df_clean['Poli Asal'].astype(str).str.strip()
        
        # Jenis Poli
        if 'Jenis Poli' in df_clean.columns:
            df_clean['Jenis Poli'] = df_clean['Jenis Poli'].astype(str).str.strip()
            # Normalisasi nilai
            df_clean['Jenis Poli'] = df_clean['Jenis Poli'].replace({
                'reguler': 'Reguler',
                'poleks': 'Poleks',
                'Polek': 'Poleks'
            })
        
        # 5. Clean kolom hari
        hari_cols = ['Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at", 'Sabtu']
        for hari in hari_cols:
            if hari in df_clean.columns:
                # Konversi ke string dan clean
                df_clean[hari] = df_clean[hari].astype(str).str.strip()
                # Replace nilai kosong
                df_clean[hari] = df_clean[hari].replace(['nan', 'NaN', 'NaT', 'None', ''], None)
        
        # 6. Drop baris yang semua kolom harinya kosong
        hari_cols_exist = [h for h in hari_cols if h in df_clean.columns]
        if hari_cols_exist:
            # Hapus baris yang semua kolom hari kosong
            mask = df_clean[hari_cols_exist].isnull().all(axis=1)
            df_clean = df_clean[~mask].reset_index(drop=True)
        
        # 7. Reset index
        df_clean = df_clean.reset_index(drop=True)
        
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
