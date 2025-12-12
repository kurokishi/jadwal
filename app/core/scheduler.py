"""
Scheduler - Modul utama untuk memproses jadwal dokter
Mengubah data Excel mentah (sheet Reguler & Poleks) menjadi grid jadwal
"""

import pandas as pd
import numpy as np
from typing import Dict, List, Tuple, Optional
from datetime import datetime, time
import re
import traceback


class Scheduler:
    def __init__(self, parser, cleaner, config):
        """
        Inisialisasi Scheduler
        
        Args:
            parser: TimeParser instance untuk parsing waktu
            cleaner: DataCleaner instance untuk cleaning data
            config: Config instance untuk konfigurasi
        """
        self.parser = parser
        self.cleaner = cleaner
        self.config = config
        
        # Debug info
        print(f"✅ Scheduler initialized")
        print(f"   - Start time: {config.start_hour:02d}:{config.start_minute:02d}")
        print(f"   - Interval: {config.interval_minutes} minutes")
        print(f"   - Max poleks per slot: {config.max_poleks_per_slot}")
        print(f"   - Days: {config.hari_list}")
    
    def process_dataframe(self, df_or_file) -> Tuple[Optional[pd.DataFrame], List[str], List[str]]:
        """
        Proses data dari file Excel atau DataFrame menjadi grid jadwal
        
        Args:
            df_or_file: Bisa berupa:
                      1. BytesIO (file upload Streamlit)
                      2. DataFrame (data sudah dibaca)
                      3. String path ke file
            
        Returns:
            Tuple: (grid_df, slot_strings, error_messages)
                   grid_df: DataFrame hasil dalam format grid
                   slot_strings: List string waktu slot (contoh: ["07:30", "08:00", ...])
                   error_messages: List pesan error/warning
        """
        error_messages = []
        
        try:
            print("=" * 50)
            print("🔄 STARTING DATA PROCESSING")
            print("=" * 50)
            
            # 1. CLEAN DATA
            print("\n1️⃣ CLEANING DATA...")
            cleaned_df = self.cleaner.clean(df_or_file)
            
            if cleaned_df.empty:
                error_msg = "❌ Data setelah cleaning kosong"
                print(error_msg)
                error_messages.append(error_msg)
                return None, [], error_messages
            
            print(f"   ✓ Data cleaned: {cleaned_df.shape[0]} rows × {cleaned_df.shape[1]} columns")
            print(f"   ✓ Columns: {list(cleaned_df.columns)}")
            
            # Preview beberapa baris
            if cleaned_df.shape[0] > 0:
                print(f"   ✓ Sample data:")
                for i in range(min(3, cleaned_df.shape[0])):
                    dokter = cleaned_df.iloc[i].get('Nama Dokter', 'N/A')
                    poli = cleaned_df.iloc[i].get('Poli Asal', 'N/A')
                    print(f"     - Dr. {dokter} | Poli {poli}")
            
            # 2. GENERATE TIME SLOTS
            print("\n2️⃣ GENERATING TIME SLOTS...")
            slot_strings = self._generate_slot_strings()
            
            if not slot_strings:
                error_msg = "❌ Gagal generate time slots"
                print(error_msg)
                error_messages.append(error_msg)
                return None, [], error_messages
            
            print(f"   ✓ Generated {len(slot_strings)} time slots")
            print(f"   ✓ First 5 slots: {slot_strings[:5]}")
            print(f"   ✓ Last 5 slots: {slot_strings[-5:]}")
            
            # 3. PARSE TIME TO SLOTS
            print("\n3️⃣ PARSING TIME RANGES TO SLOTS...")
            slot_df = self._parse_time_to_slots(cleaned_df, slot_strings)
            
            if slot_df.empty:
                error_msg = "❌ Tidak ada data waktu yang berhasil di-parse"
                print(error_msg)
                error_messages.append(error_msg)
                
                # Debug: coba lihat data waktu yang ada
                print("   Debug - Checking time data in cleaned_df:")
                hari_cols = [col for col in cleaned_df.columns if col in self.config.hari_list]
                for hari in hari_cols:
                    if hari in cleaned_df.columns:
                        non_empty = cleaned_df[hari].dropna()
                        if len(non_empty) > 0:
                            print(f"   - {hari}: {len(non_empty)} non-empty entries")
                            print(f"     Sample: {non_empty.iloc[0] if len(non_empty) > 0 else 'N/A'}")
                
                return None, [], error_messages
            
            print(f"   ✓ Parsed {len(slot_df)} time slot entries")
            print(f"   ✓ Unique doctors: {slot_df['DOKTER'].nunique()}")
            print(f"   ✓ Unique poli: {slot_df['POLI'].nunique()}")
            print(f"   ✓ Slot types - R: {(slot_df['KODE'] == 'R').sum()}, E: {(slot_df['KODE'] == 'E').sum()}")
            
            # 4. CREATE GRID FORMAT
            print("\n4️⃣ CREATING GRID FORMAT...")
            grid_df = self._create_grid_format(slot_df, slot_strings)
            
            if grid_df is None or grid_df.empty:
                error_msg = "❌ Grid data kosong setelah diproses"
                print(error_msg)
                error_messages.append(error_msg)
                return None, [], error_messages
            
            print(f"   ✓ Grid created: {grid_df.shape[0]} rows × {grid_df.shape[1]} columns")
            print(f"   ✓ Grid columns sample: {list(grid_df.columns)[:8]}...")
            
            # 5. VALIDATE GRID
            print("\n5️⃣ VALIDATING GRID...")
            validation_errors = self._validate_grid(grid_df, slot_strings)
            
            if validation_errors:
                print(f"   ⚠️ Found {len(validation_errors)} validation warnings")
                error_messages.extend(validation_errors)
            else:
                print(f"   ✓ Grid validation passed")
            
            # 6. FINAL CHECK
            print("\n6️⃣ FINAL CHECK...")
            stats = self._calculate_statistics(grid_df, slot_strings)
            
            print(f"   ✓ Statistics:")
            print(f"     - Total rows: {stats['total_rows']}")
            print(f"     - Total R slots: {stats['total_r']}")
            print(f"     - Total E slots: {stats['total_e']}")
            print(f"     - Total doctors: {stats['total_doctors']}")
            print(f"     - Total poli: {stats['total_poli']}")
            
            print("\n" + "=" * 50)
            print("✅ PROCESSING COMPLETE SUCCESSFULLY")
            print("=" * 50)
            
            return grid_df, slot_strings, error_messages
            
        except Exception as e:
            error_msg = f"❌ Error processing data: {str(e)}"
            print(error_msg)
            print(traceback.format_exc())
            error_messages.append(error_msg)
            return None, [], error_messages
    
    def _generate_slot_strings(self) -> List[str]:
        """
        Generate list slot waktu berdasarkan konfigurasi
        
        Returns:
            List string format "HH:MM" dari start_time sampai 14:30
        """
        try:
            slot_strings = []
            current_time = self.config.start_hour * 60 + self.config.start_minute
            end_time = 14 * 60 + 30  # Hard stop at 14:30
            
            while current_time < end_time:
                hours = current_time // 60
                minutes = current_time % 60
                time_str = f"{hours:02d}:{minutes:02d}"
                slot_strings.append(time_str)
                current_time += self.config.interval_minutes
            
            return slot_strings
            
        except Exception as e:
            print(f"❌ Error in _generate_slot_strings: {e}")
            return []
    
    def _parse_time_to_slots(self, df: pd.DataFrame, slot_strings: List[str]) -> pd.DataFrame:
        """
        Parse waktu dari kolom hari ke dalam slot-slot waktu
        
        Args:
            df: DataFrame yang sudah dibersihkan
            slot_strings: List slot waktu
            
        Returns:
            DataFrame dengan kolom: POLI, JENIS, HARI, DOKTER, SLOT, KODE
        """
        try:
            result_rows = []
            hari_list = self.config.hari_list
            
            print(f"   Parsing time for {len(df)} rows")
            print(f"   Days to parse: {hari_list}")
            
            for idx, row in df.iterrows():
                nama_dokter = str(row.get('Nama Dokter', '')).strip()
                poli_asal = str(row.get('Poli Asal', '')).strip()
                jenis_poli = str(row.get('Jenis Poli', '')).strip()
                
                # Skip jika data penting kosong
                if not nama_dokter or nama_dokter.lower() in ['nan', 'null', '']:
                    continue
                if not poli_asal or poli_asal.lower() in ['nan', 'null', '']:
                    continue
                
                for hari in hari_list:
                    if hari in row and pd.notna(row[hari]):
                        time_range = str(row[hari]).strip()
                        
                        # Skip jika kosong
                        if not time_range or time_range.lower() in ['nan', 'null', 'none', '']:
                            continue
                        
                        # Parse waktu
                        slots = self.parser.parse_time_range(time_range, slot_strings)
                        
                        if slots:
                            # Tentukan kode berdasarkan jenis poli
                            if 'poleks' in jenis_poli.lower():
                                kode = 'E'
                            else:
                                kode = 'R'
                            
                            for slot in slots:
                                result_rows.append({
                                    'POLI': poli_asal,
                                    'JENIS': jenis_poli,
                                    'HARI': hari,
                                    'DOKTER': nama_dokter,
                                    'SLOT': slot,
                                    'KODE': kode
                                })
                        else:
                            # Debug: waktu tidak bisa di-parse
                            print(f"   ⚠️ Could not parse time: '{time_range}' for Dr. {nama_dokter} on {hari}")
            
            result_df = pd.DataFrame(result_rows)
            
            if not result_df.empty:
                print(f"   ✓ Successfully parsed {len(result_df)} time slots")
                print(f"   ✓ Sample entries:")
                for i in range(min(3, len(result_df))):
                    entry = result_df.iloc[i]
                    print(f"     - {entry['HARI']} {entry['SLOT']}: Dr. {entry['DOKTER']} ({entry['KODE']})")
            
            return result_df
            
        except Exception as e:
            print(f"❌ Error in _parse_time_to_slots: {e}")
            print(traceback.format_exc())
            return pd.DataFrame()
    
    def _create_grid_format(self, slot_df: pd.DataFrame, slot_strings: List[str]) -> pd.DataFrame:
        """
        Ubah format slot menjadi grid (pivot format)
        
        Args:
            slot_df: DataFrame dari _parse_time_to_slots
            slot_strings: List slot waktu
            
        Returns:
            DataFrame grid dengan kolom: POLI, JENIS, HARI, DOKTER, JAM, [slot1], [slot2], ...
        """
        try:
            if slot_df.empty:
                print("   ⚠️ No data to create grid")
                return pd.DataFrame()
            
            print(f"   Creating grid from {len(slot_df)} slot entries")
            
            # Group by kombinasi unik
            unique_combos = slot_df[['POLI', 'JENIS', 'HARI', 'DOKTER']].drop_duplicates()
            print(f"   Found {len(unique_combos)} unique doctor-day combinations")
            
            grid_rows = []
            
            for idx, combo in unique_combos.iterrows():
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
                
                if combo_data.empty:
                    continue
                
                # Buat dictionary untuk row ini
                row_data = {
                    'POLI': poli,
                    'JENIS': jenis,
                    'HARI': hari,
                    'DOKTER': dokter,
                    'JAM': self._get_time_range(combo_data, slot_strings)
                }
                
                # Tambahkan kolom untuk setiap slot
                for slot in slot_strings:
                    slot_mask = combo_data['SLOT'] == slot
                    if slot_mask.any():
                        kode = combo_data[slot_mask]['KODE'].iloc[0]
                        row_data[slot] = kode
                    else:
                        row_data[slot] = ''
                
                grid_rows.append(row_data)
            
            # Buat DataFrame
            grid_df = pd.DataFrame(grid_rows)
            
            # Urutkan kolom: metadata dulu, lalu slot waktu
            meta_columns = ['POLI', 'JENIS', 'HARI', 'DOKTER', 'JAM']
            time_columns = [col for col in grid_df.columns if col not in meta_columns]
            sorted_columns = meta_columns + sorted(time_columns)
            
            grid_df = grid_df[sorted_columns]
            
            print(f"   ✓ Grid created with {len(grid_df)} rows")
            print(f"   ✓ Grid columns: {len(grid_df.columns)} total")
            print(f"   ✓ Time columns: {len(time_columns)} slots")
            
            return grid_df
            
        except Exception as e:
            print(f"❌ Error in _create_grid_format: {e}")
            print(traceback.format_exc())
            return pd.DataFrame()
    
    def _get_time_range(self, df: pd.DataFrame, slot_strings: List[str]) -> str:
        """
        Dapatkan range waktu dari slot yang aktif (untuk kolom JAM)
        
        Args:
            df: DataFrame untuk satu dokter di satu hari
            slot_strings: List semua slot waktu
            
        Returns:
            String format "HH:MM-HH:MM, HH:MM-HH:MM, ..."
        """
        if df.empty:
            return ''
        
        # Dapatkan slot yang aktif
        active_slots = sorted(df['SLOT'].tolist())
        if not active_slots:
            return ''
        
        # Gabungkan slot berurutan menjadi ranges
        ranges = []
        start = active_slots[0]
        end = active_slots[0]
        
        for i in range(1, len(active_slots)):
            current_slot = active_slots[i]
            current_idx = slot_strings.index(current_slot)
            prev_idx = slot_strings.index(end)
            
            # Jika slot berurutan (berdasarkan index)
            if current_idx == prev_idx + 1:
                end = current_slot
            else:
                # Tambahkan range saat ini
                if start == end:
                    ranges.append(start)
                else:
                    ranges.append(f"{start}-{end}")
                
                # Mulai range baru
                start = current_slot
                end = current_slot
        
        # Tambahkan range terakhir
        if start == end:
            ranges.append(start)
        else:
            ranges.append(f"{start}-{end}")
        
        return ", ".join(ranges)
    
    def _validate_grid(self, grid_df: pd.DataFrame, slot_strings: List[str]) -> List[str]:
        """
        Validasi grid untuk konflik dan batasan
        
        Args:
            grid_df: DataFrame grid
            slot_strings: List slot waktu
            
        Returns:
            List pesan error/warning
        """
        errors = []
        
        try:
            print(f"   Validating grid with {len(grid_df)} rows...")
            
            if grid_df.empty:
                errors.append("Grid data kosong")
                return errors
            
            # 1. Validasi max_poleks_per_slot per hari
            hari_list = self.config.hari_list
            
            for hari in hari_list:
                hari_data = grid_df[grid_df['HARI'] == hari]
                
                if hari_data.empty:
                    continue
                
                for slot in slot_strings:
                    if slot not in hari_data.columns:
                        continue
                    
                    # Hitung poleks di slot ini
                    poleks_count = (hari_data[slot] == 'E').sum()
                    
                    if poleks_count > self.config.max_poleks_per_slot:
                        errors.append(
                            f"⚠️ Hari {hari}, Slot {slot}: "
                            f"Poleks melebihi batas ({poleks_count} > {self.config.max_poleks_per_slot})"
                        )
            
            print(f"   ✓ Max poleks validation: {len([e for e in errors if 'Poleks melebihi batas' in e])} warnings")
            
            # 2. Validasi konflik dokter (dokter yang sama di poli berbeda di slot yang sama)
            doctor_conflicts = []
            
            for hari in hari_list:
                hari_data = grid_df[grid_df['HARI'] == hari]
                
                if hari_data.empty:
                    continue
                
                doctors = hari_data['DOKTER'].unique()
                
                for dokter in doctors:
                    doc_data = hari_data[hari_data['DOKTER'] == dokter]
                    
                    # Jika dokter muncul di lebih dari 1 baris (berarti di poli berbeda)
                    if len(doc_data) > 1:
                        for slot in slot_strings:
                            if slot not in doc_data.columns:
                                continue
                            
                            # Cek slot yang aktif untuk dokter ini
                            active_rows = doc_data[doc_data[slot].isin(['R', 'E'])]
                            
                            if len(active_rows) > 1:
                                # Dokter aktif di lebih dari 1 poli di slot yang sama
                                active_polis = active_rows['POLI'].tolist()
                                doctor_conflicts.append({
                                    'dokter': dokter,
                                    'hari': hari,
                                    'slot': slot,
                                    'polis': active_polis
                                })
            
            # Format conflict messages
            for conflict in doctor_conflicts[:5]:  # Limit to 5 conflicts
                errors.append(
                    f"⚠️ Konflik: Dr. {conflict['dokter']} di {conflict['hari']} jam {conflict['slot']} "
                    f"berada di {len(conflict['polis'])} poli: {', '.join(conflict['polis'])}"
                )
            
            if len(doctor_conflicts) > 5:
                errors.append(f"⚠️ ... dan {len(doctor_conflicts) - 5} konflik dokter lainnya")
            
            print(f"   ✓ Doctor conflict validation: {len(doctor_conflicts)} conflicts found")
            
            # 3. Validasi data kosong/tidak valid
            invalid_codes = []
            valid_codes = ['', 'R', 'E']
            
            for slot in slot_strings:
                if slot in grid_df.columns:
                    invalid_mask = ~grid_df[slot].isin(valid_codes)
                    if invalid_mask.any():
                        invalid_rows = grid_df[invalid_mask]
                        invalid_codes.extend(invalid_rows[slot].unique())
            
            if invalid_codes:
                errors.append(f"⚠️ Kode tidak valid ditemukan: {set(invalid_codes)}")
            
            print(f"   ✓ Data validation: {len(invalid_codes)} invalid codes found")
            
            return errors
            
        except Exception as e:
            errors.append(f"Error during validation: {str(e)}")
            return errors
    
    def _calculate_statistics(self, grid_df: pd.DataFrame, slot_strings: List[str]) -> Dict:
        """
        Hitung statistik dari grid data
        
        Returns:
            Dictionary berisi statistik
        """
        stats = {
            'total_rows': len(grid_df),
            'total_r': 0,
            'total_e': 0,
            'total_doctors': 0,
            'total_poli': 0,
            'total_slots': 0
        }
        
        try:
            if not grid_df.empty:
                # Hitung R dan E
                for slot in slot_strings:
                    if slot in grid_df.columns:
                        stats['total_r'] += (grid_df[slot] == 'R').sum()
                        stats['total_e'] += (grid_df[slot] == 'E').sum()
                
                # Hitung unik
                stats['total_doctors'] = grid_df['DOKTER'].nunique()
                stats['total_poli'] = grid_df['POLI'].nunique()
                stats['total_slots'] = len(grid_df) * len(slot_strings)
            
        except Exception as e:
            print(f"⚠️ Error calculating statistics: {e}")
        
        return stats
    
    def get_slot_strings(self) -> List[str]:
        """
        Dapatkan list slot strings untuk digunakan di UI
        
        Returns:
            List string slot waktu
        """
        return self._generate_slot_strings()
    
    def generate_sample_grid(self) -> pd.DataFrame:
        """
        Generate sample grid untuk testing
        
        Returns:
            DataFrame sample
        """
        try:
            # Generate slot strings
            slot_strings = self._generate_slot_strings()
            
            # Sample data
            sample_rows = []
            
            sample_data = [
                {
                    'POLI': 'Poli Anak',
                    'JENIS': 'Reguler',
                    'HARI': 'Senin',
                    'DOKTER': 'dr. Contoh',
                    'JAM': '07:30-10:00'
                },
                {
                    'POLI': 'Poli Dalam',
                    'JENIS': 'Poleks',
                    'HARI': 'Selasa',
                    'DOKTER': 'dr. Sample',
                    'JAM': '08:00-09:30'
                }
            ]
            
            for data in sample_data:
                row = data.copy()
                # Add time slots
                for slot in slot_strings[:10]:  # First 10 slots only for sample
                    row[slot] = 'R' if data['JENIS'] == 'Reguler' else 'E'
                sample_rows.append(row)
            
            return pd.DataFrame(sample_rows)
            
        except Exception as e:
            print(f"Error generating sample: {e}")
            return pd.DataFrame()
    
    def export_to_excel_format(self, grid_df: pd.DataFrame, slot_strings: List[str]) -> Dict:
        """
        Format data untuk export ke Excel
        
        Returns:
            Dictionary dengan berbagai format data untuk ExcelWriter
        """
        result = {
            'jadwal_grid': grid_df,
            'slot_strings': slot_strings,
            'rekap_layanan': None,
            'rekap_poli': None,
            'rekap_dokter': None
        }
        
        try:
            if grid_df.empty:
                return result
            
            # 1. Rekap Layanan
            rekap_rows = []
            for _, row in grid_df.iterrows():
                active_slots = []
                for slot in slot_strings:
                    if slot in row and row[slot] in ['R', 'E']:
                        active_slots.append(slot)
                
                if active_slots:
                    # Gabungkan slot berurutan
                    time_ranges = self._combine_slots_to_ranges(active_slots, slot_strings)
                    for time_range in time_ranges:
                        rekap_rows.append({
                            'POLI': row['POLI'],
                            'HARI': row['HARI'],
                            'DOKTER': row['DOKTER'],
                            'JENIS': row['JENIS'],
                            'WAKTU LAYANAN': time_range
                        })
            
            if rekap_rows:
                result['rekap_layanan'] = pd.DataFrame(rekap_rows)
            
            # 2. Rekap Poli
            poli_rows = []
            for (poli, hari), group in grid_df.groupby(['POLI', 'HARI']):
                total_r = total_e = 0
                for slot in slot_strings:
                    if slot in group.columns:
                        total_r += (group[slot] == 'R').sum()
                        total_e += (group[slot] == 'E').sum()
                
                # Konversi ke jam
                interval_hours = self.config.interval_minutes / 60
                jam_r = round(total_r * interval_hours, 2)
                jam_e = round(total_e * interval_hours, 2)
                total_jam = jam_r + jam_e
                
                poli_rows.append({
                    'POLI': poli,
                    'HARI': hari,
                    'REGULER (JAM)': jam_r,
                    'POLEKS (JAM)': jam_e,
                    'TOTAL': total_jam
                })
            
            if poli_rows:
                result['rekap_poli'] = pd.DataFrame(poli_rows)
            
            # 3. Rekap Dokter
            dokter_rows = []
            for (dokter, hari), group in grid_df.groupby(['DOKTER', 'HARI']):
                active_slots = []
                for slot in slot_strings:
                    if slot in group.columns:
                        if (group[slot].isin(['R', 'E'])).any():
                            active_slots.append(slot)
                
                if active_slots:
                    time_ranges = self._combine_slots_to_ranges(active_slots, slot_strings)
                    for time_range in time_ranges:
                        # Hitung durasi
                        slots_in_range = self._get_slots_in_range(time_range, slot_strings)
                        durasi = len(slots_in_range) * (self.config.interval_minutes / 60)
                        
                        dokter_rows.append({
                            'DOKTER': dokter,
                            'HARI': hari,
                            'SHIFT': time_range,
                            'TOTAL JAM': round(durasi, 2)
                        })
            
            if dokter_rows:
                result['rekap_dokter'] = pd.DataFrame(dokter_rows)
            
        except Exception as e:
            print(f"⚠️ Error in export formatting: {e}")
        
        return result
    
    def _combine_slots_to_ranges(self, slots: List[str], slot_strings: List[str]) -> List[str]:
        """Gabungkan slot menjadi range waktu"""
        if not slots:
            return []
        
        # Sort slots berdasarkan urutan di slot_strings
        slots_sorted = sorted(slots, key=lambda x: slot_strings.index(x))
        
        ranges = []
        start = slots_sorted[0]
        end = slots_sorted[0]
        
        for i in range(1, len(slots_sorted)):
            current_idx = slot_strings.index(slots_sorted[i])
            prev_idx = slot_strings.index(end)
            
            if current_idx == prev_idx + 1:
                end = slots_sorted[i]
            else:
                ranges.append(f"{start}-{end}")
                start = slots_sorted[i]
                end = slots_sorted[i]
        
        ranges.append(f"{start}-{end}")
        return ranges
    
    def _get_slots_in_range(self, time_range: str, slot_strings: List[str]) -> List[str]:
        """Dapatkan semua slot dalam range waktu"""
        if '-' not in time_range:
            return [time_range] if time_range in slot_strings else []
        
        start_str, end_str = time_range.split('-')
        start_idx = slot_strings.index(start_str) if start_str in slot_strings else -1
        end_idx = slot_strings.index(end_str) if end_str in slot_strings else -1
        
        if start_idx == -1 or end_idx == -1:
            return []
        
        return slot_strings[start_idx:end_idx + 1]
