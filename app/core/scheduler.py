# app/core/scheduler.py
import pandas as pd
from typing import Dict, List, Tuple


class Scheduler:
    def __init__(self, parser, cleaner, config):
        self.parser = parser
        self.cleaner = cleaner
        self.config = config
        
    def process_dataframe(self, df_upload):
        """
        Proses dataframe mentah menjadi jadwal grid
        
        Args:
            df_upload: DataFrame dari file Excel upload
            
        Returns:
            Tuple: (grid_df, slot_strings, error_messages)
        """
        error_messages = []
        
        # 1. Clean data
        try:
            cleaned_df = self.cleaner.clean(df_upload)
        except Exception as e:
            error_messages.append(f"Error cleaning data: {str(e)}")
            return None, [], error_messages
        
        # 2. Validasi data
        if cleaned_df.empty:
            error_messages.append("Data setelah cleaning kosong")
            return None, [], error_messages
        
        # 3. Generate slot strings berdasarkan konfigurasi
        slot_strings = self._generate_slot_strings()
        
        # 4. Parse waktu ke slot
        try:
            slot_df = self._parse_time_to_slots(cleaned_df, slot_strings)
        except Exception as e:
            error_messages.append(f"Error parsing waktu: {str(e)}")
            return None, [], error_messages
        
        # 5. Create grid format
        try:
            grid_df = self._create_grid_format(slot_df, slot_strings)
        except Exception as e:
            error_messages.append(f"Error membuat grid: {str(e)}")
            return None, [], error_messages
        
        # 6. Validasi grid
        validation_errors = self._validate_grid(grid_df, slot_strings)
        if validation_errors:
            error_messages.extend(validation_errors)
        
        return grid_df, slot_strings, error_messages
    
    def _generate_slot_strings(self) -> List[str]:
        """
        Generate list slot waktu berdasarkan konfigurasi
        """
        slot_strings = []
        current_time = self.config.start_hour * 60 + self.config.start_minute
        
        while True:
            # Format: HH:MM
            hours = current_time // 60
            minutes = current_time % 60
            time_str = f"{hours:02d}:{minutes:02d}"
            
            slot_strings.append(time_str)
            
            # Cek apakah sudah melewati end time (14:30)
            if hours > 14 or (hours == 14 and minutes >= 30):
                break
                
            current_time += self.config.interval_minutes
        
        return slot_strings
    
    def _parse_time_to_slots(self, df: pd.DataFrame, slot_strings: List[str]) -> pd.DataFrame:
        """
        Parse waktu dari kolom hari ke slot
        """
        hari_list = self.config.hari_list
        
        # Copy dataframe untuk hasil
        result_rows = []
        
        for _, row in df.iterrows():
            nama_dokter = row.get('Nama Dokter', '')
            poli_asal = row.get('Poli Asal', '')
            jenis_poli = row.get('Jenis Poli', '')
            
            for hari in hari_list:
                if hari in row and pd.notna(row[hari]):
                    time_range = str(row[hari])
                    
                    # Parse waktu menggunakan time_parser
                    slots = self.parser.parse_time_range(time_range, slot_strings)
                    
                    if slots:
                        # Tentukan kode berdasarkan jenis poli
                        kode = 'E' if jenis_poli == 'Poleks' else 'R'
                        
                        for slot in slots:
                            result_rows.append({
                                'POLI': poli_asal,
                                'JENIS': jenis_poli,
                                'HARI': hari,
                                'DOKTER': nama_dokter,
                                'SLOT': slot,
                                'KODE': kode
                            })
        
        return pd.DataFrame(result_rows)
    
    def _create_grid_format(self, slot_df: pd.DataFrame, slot_strings: List[str]) -> pd.DataFrame:
        """
        Ubah format slot menjadi grid
        """
        if slot_df.empty:
            return pd.DataFrame()
        
        # Group by kombinasi unik
        grid_rows = []
        
        # Buat semua kombinasi unik
        unique_combos = slot_df[['POLI', 'JENIS', 'HARI', 'DOKTER']].drop_duplicates()
        
        for _, combo in unique_combos.iterrows():
            poli = combo['POLI']
            jenis = combo['JENIS']
            hari = combo['HARI']
            dokter = combo['DOKTER']
            
            # Filter data untuk kombinasi ini
            mask = (
                (slot_df['POLI'] == poli) &
                (slot_df['JENIS'] == jenis) &
                (slot_df['HARI'] == hari) &
                (slot_df['DOKTER'] == dokter)
            )
            combo_data = slot_df[mask]
            
            # Buat dictionary untuk semua slot
            row_data = {
                'POLI': poli,
                'JENIS': jenis,
                'HARI': hari,
                'DOKTER': dokter,
                'JAM': self._get_time_range(combo_data, slot_strings)
            }
            
            # Tambahkan kolom untuk setiap slot
            for slot in slot_strings:
                # Cek jika slot ada dalam data
                slot_mask = combo_data['SLOT'] == slot
                if slot_mask.any():
                    kode = combo_data[slot_mask]['KODE'].iloc[0]
                    row_data[slot] = kode
                else:
                    row_data[slot] = ''
            
            grid_rows.append(row_data)
        
        return pd.DataFrame(grid_rows)
    
    def _get_time_range(self, df: pd.DataFrame, slot_strings: List[str]) -> str:
        """
        Dapatkan range waktu dari slot yang aktif
        """
        if df.empty:
            return ''
        
        slots = sorted(df['SLOT'].tolist())
        if not slots:
            return ''
        
        # Gabungkan slot berurutan
        ranges = []
        start = slots[0]
        end = slots[0]
        
        for slot in slots[1:]:
            current_idx = slot_strings.index(slot)
            prev_idx = slot_strings.index(end)
            
            if current_idx == prev_idx + 1:
                end = slot
            else:
                ranges.append(f"{start}-{end}")
                start = slot
                end = slot
        
        ranges.append(f"{start}-{end}")
        return ", ".join(ranges)
    
    def _validate_grid(self, grid_df: pd.DataFrame, slot_strings: List[str]) -> List[str]:
        """
        Validasi grid untuk konflik dan batasan
        """
        errors = []
        
        if grid_df.empty:
            errors.append("Grid data kosong")
            return errors
        
        # 1. Validasi max_poleks_per_slot per hari
        hari_list = self.config.hari_list
        for hari in hari_list:
            hari_data = grid_df[grid_df['HARI'] == hari]
            
            for slot in slot_strings:
                poleks_count = 0
                
                for _, row in hari_data.iterrows():
                    if row.get(slot) == 'E':
                        poleks_count += 1
                
                if poleks_count > self.config.max_poleks_per_slot:
                    errors.append(
                        f"⚠️ Hari {hari}, Slot {slot}: "
                        f"Poleks melebihi batas ({poleks_count} > {self.config.max_poleks_per_slot})"
                    )
        
        # 2. Validasi konflik dokter (dokter yang sama di poli berbeda di slot yang sama)
        for hari in hari_list:
            hari_data = grid_df[grid_df['HARI'] == hari]
            doctors = hari_data['DOKTER'].unique()
            
            for dokter in doctors:
                doc_data = hari_data[hari_data['DOKTER'] == dokter]
                
                if len(doc_data) > 1:  # Dokter yang sama muncul di >1 poli
                    for slot in slot_strings:
                        active_polis = []
                        
                        for _, row in doc_data.iterrows():
                            if row.get(slot) in ['R', 'E']:
                                active_polis.append(row['POLI'])
                        
                        if len(active_polis) > 1:
                            errors.append(
                                f"⚠️ Konflik: Dr. {dokter} di {hari} jam {slot} "
                                f"berada di {len(active_polis)} poli: {', '.join(active_polis)}"
                            )
        
        return errors
    
    def get_slot_strings(self) -> List[str]:
        """
        Dapatkan list slot strings untuk digunakan di UI
        """
        return self._generate_slot_strings()
