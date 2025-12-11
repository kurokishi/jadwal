import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import time, timedelta, datetime
import io
import re
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from typing import Dict, List, Tuple, Optional, Any
import traceback
import warnings

# Suppress warnings
warnings.filterwarnings('ignore')

# ============================================================================
# KELAS UTAMA - Aplikasi Pengisi Jadwal Poli
# ============================================================================

class PoliSchedulerApp:
    """Kelas utama aplikasi Pengisi Jadwal Poli"""
    
    def __init__(self):
        """Inisialisasi aplikasi"""
        self.setup_page_config()
        self.initialize_session_state()
        self.initialize_constants()
        
    def setup_page_config(self):
        """Setup konfigurasi halaman Streamlit"""
        st.set_page_config(
            page_title="🏥 Pengisi Jadwal Poli Excel - OOP Version",
            page_icon="🏥",
            layout="wide",
            initial_sidebar_state="expanded"
        )
        
    def initialize_session_state(self):
        """Inisialisasi session state"""
        if 'processed_data' not in st.session_state:
            st.session_state.processed_data = None
        if 'error_logs' not in st.session_state:
            st.session_state.error_logs = []
        if 'processing_stats' not in st.session_state:
            st.session_state.processing_stats = {}
        if 'visualization_data' not in st.session_state:
            st.session_state.visualization_data = None
            
    def initialize_constants(self):
        """Inisialisasi konstan"""
        # Warna untuk sel
        self.FILL_R = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
        self.FILL_E = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
        self.FILL_OVERLIMIT = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
        
        # Time slots default
        self.TIME_SLOTS = [
            time(7, 30), time(8, 0), time(8, 30), time(9, 0), time(9, 30),
            time(10, 0), time(10, 30), time(11, 0), time(11, 30), time(12, 0),
            time(12, 30), time(13, 0), time(13, 30), time(14, 0), time(14, 30)
        ]
        self.TIME_SLOTS_STR = [t.strftime("%H:%M") for t in self.TIME_SLOTS]
        
        # Hari
        self.HARI_ORDER = {"Senin": 1, "Selasa": 2, "Rabu": 3, "Kamis": 4, "Jum'at": 5}
        self.HARI_INDONESIA = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]
        
        # Konfigurasi default
        self.config = {
            'start_hour': 7,
            'start_minute': 30,
            'interval_minutes': 30,
            'max_poleks_per_slot': 7,
            'auto_fix_errors': True
        }


# ============================================================================
# KELAS UTILITAS
# ============================================================================

class ErrorAnalyzer:
    """Kelas untuk menganalisis dan menangani error"""
    
    @staticmethod
    def analyze_excel_structure(df: pd.DataFrame, sheet_name: str) -> Dict[str, Any]:
        """Analisis struktur dataframe Excel"""
        errors = []
        warnings_list = []
        
        # Cek kolom yang diperlukan
        required_columns = ['Nama Dokter', 'Poli Asal', 'Jenis Poli']
        hari_columns = ['Senin', 'Selasa', 'Rabu', 'Kamis', "Jum'at"]
        
        missing_required = [col for col in required_columns if col not in df.columns]
        if missing_required:
            errors.append(f"Kolom yang diperlukan tidak ditemukan: {missing_required}")
        
        missing_hari = [col for col in hari_columns if col not in df.columns]
        if missing_hari:
            warnings_list.append(f"Kolom hari yang tidak ditemukan: {missing_hari}")
        
        # Cek data kosong
        empty_rows = df[required_columns].isnull().all(axis=1).sum()
        if empty_rows > 0:
            warnings_list.append(f"{empty_rows} baris kosong ditemukan")
        
        # Cek format waktu
        time_format_errors = 0
        for hari in hari_columns:
            if hari in df.columns:
                for time_str in df[hari].dropna():
                    if not ErrorAnalyzer._validate_time_format(str(time_str)):
                        time_format_errors += 1
        
        if time_format_errors > 0:
            errors.append(f"{time_format_errors} format waktu tidak valid ditemukan")
        
        return {
            'is_valid': len(errors) == 0,
            'errors': errors,
            'warnings': warnings_list,
            'total_rows': len(df),
            'missing_columns': missing_required + missing_hari,
            'time_format_errors': time_format_errors
        }
    
    @staticmethod
    def _validate_time_format(time_str: str) -> bool:
        """Validasi format waktu"""
        if pd.isna(time_str) or str(time_str).strip() == "":
            return True
        
        patterns = [
            r'^\d{1,2}\.\d{2}\s*-\s*\d{1,2}\.\d{2}$',  # 08.00 - 10.30
            r'^\d{1,2}:\d{2}\s*-\s*\d{1,2}:\d{2}$',    # 08:00 - 10:30
            r'^\d{1,2}\.\d{2}-\d{1,2}\.\d{2}$',        # 08.00-10.30
        ]
        
        clean_str = str(time_str).strip()
        return any(re.match(pattern, clean_str) for pattern in patterns)
    
    @staticmethod
    def create_error_report(error_data: Dict[str, Any]) -> str:
        """Buat laporan error"""
        report = "## 📋 Laporan Analisis Error\n\n"
        
        if error_data['errors']:
            report += "### ❌ Errors:\n"
            for error in error_data['errors']:
                report += f"- {error}\n"
            report += "\n"
        
        if error_data['warnings']:
            report += "### ⚠️ Warnings:\n"
            for warning in error_data['warnings']:
                report += f"- {warning}\n"
            report += "\n"
        
        report += f"### 📊 Statistik:\n"
        report += f"- Total Baris: {error_data['total_rows']}\n"
        report += f"- Format Waktu Error: {error_data['time_format_errors']}\n"
        report += f"- Status Valid: {'✅ Ya' if error_data['is_valid'] else '❌ Tidak'}\n"
        
        return report


class DataProcessor:
    """Kelas untuk memproses data jadwal"""
    
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        self.time_parser = TimeParser()
        self.data_cleaner = DataCleaner()
        
    def process_schedule(self, df: pd.DataFrame, jenis_poli: str) -> pd.DataFrame:
        """Proses dataframe jadwal"""
        # Bersihkan data terlebih dahulu
        df_cleaned = self.data_cleaner.clean_data(df, jenis_poli)
        
        results = []
        
        for (dokter, poli_asal), group in df_cleaned.groupby(['Nama Dokter', 'Poli Asal']):
            hari_schedules = self._extract_schedules_per_day(group)
            
            # Generate baris per hari
            for hari in ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]:
                if hari not in hari_schedules or not hari_schedules[hari]:
                    continue
                    
                merged_ranges = self._merge_time_ranges(hari_schedules[hari])
                row = self._create_schedule_row(poli_asal, jenis_poli, hari, dokter, merged_ranges)
                results.append(row)
        
        return pd.DataFrame(results) if results else pd.DataFrame()
    
    def _extract_schedules_per_day(self, group: pd.DataFrame) -> Dict[str, List[Tuple[time, time]]]:
        """Ekstrak jadwal per hari"""
        hari_schedules = {}
        
        for hari in ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]:
            if hari not in group.columns:
                continue
                
            time_ranges = []
            for time_str in group[hari]:
                start_time, end_time = self.time_parser.parse(time_str)
                if start_time and end_time:
                    # Potong maksimal sampai 14:30
                    if end_time > time(14, 30):
                        end_time = time(14, 30)
                    time_ranges.append((start_time, end_time))
            
            hari_schedules[hari] = time_ranges
        
        return hari_schedules
    
    def _merge_time_ranges(self, time_ranges: List[Tuple[time, time]]) -> List[List[time]]:
        """Gabungkan rentang waktu yang overlapping"""
        if not time_ranges:
            return []
        
        merged = []
        for start, end in sorted(time_ranges):
            if not merged:
                merged.append([start, end])
            else:
                last_start, last_end = merged[-1]
                if start <= last_end:
                    merged[-1][1] = max(last_end, end)
                else:
                    merged.append([start, end])
        return merged
    
    def _create_schedule_row(self, poli_asal: str, jenis_poli: str, 
                            hari: str, dokter: str, 
                            merged_ranges: List[List[time]]) -> Dict[str, Any]:
        """Buat baris jadwal"""
        row = {
            'POLI ASAL': poli_asal,
            'JENIS POLI': jenis_poli,
            'HARI': hari,
            'DOKTER': dokter
        }
        
        # Isi slot waktu
        for i, slot_time in enumerate(self.config['time_slots']):
            slot_start = slot_time
            slot_end = (datetime.combine(datetime.today(), slot_start) + 
                       timedelta(minutes=30)).time()
            
            # Cek overlap
            has_overlap = any(
                not (slot_end <= start or slot_start >= end)
                for start, end in merged_ranges
            )
            
            if has_overlap:
                row[self.config['time_slots_str'][i]] = 'R' if jenis_poli == 'Reguler' else 'E'
            else:
                row[self.config['time_slots_str'][i]] = ''
        
        return row


class TimeParser:
    """Kelas untuk parsing waktu"""
    
    @staticmethod
    def parse(time_str: Any) -> Tuple[Optional[time], Optional[time]]:
        """Parse rentang waktu dari string"""
        if pd.isna(time_str) or str(time_str).strip() == "":
            return None, None
        
        clean_str = str(time_str).strip()
        
        # Normalisasi format
        clean_str = clean_str.replace(' ', '').replace('.', ':')
        
        # Cari pola waktu
        pattern = r'(\d{1,2}:\d{2})-(\d{1,2}:\d{2})'
        match = re.search(pattern, clean_str)
        
        if not match:
            return None, None
        
        try:
            start_str, end_str = match.groups()
            start_hour, start_minute = map(int, start_str.split(':'))
            end_hour, end_minute = map(int, end_str.split(':'))
            
            start_time = time(start_hour, start_minute)
            end_time = time(end_hour, end_minute)
            
            if end_time < start_time:
                end_time = time(end_hour + 24, end_minute)
            
            return start_time, end_time
        except:
            return None, None


class DataCleaner:
    """Kelas untuk membersihkan data input"""
    
    @staticmethod
    def clean_data(df: pd.DataFrame, jenis_poli: str) -> pd.DataFrame:
        """Bersihkan dan perbaiki data input"""
        df_clean = df.copy()
        
        # Pastikan kolom yang diperlukan ada
        required_cols = ['Nama Dokter', 'Poli Asal', 'Jenis Poli']
        for col in required_cols:
            if col not in df_clean.columns:
                df_clean[col] = ''
        
        # Isi jenis poli jika kosong
        df_clean['Jenis Poli'] = df_clean['Jenis Poli'].fillna(jenis_poli)
        
        # Perbaiki format waktu
        hari_cols = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]
        for col in hari_cols:
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].apply(DataCleaner._fix_time_format)
        
        # Hapus baris yang tidak memiliki jadwal sama sekali
        if hari_cols[0] in df_clean.columns:
            has_schedule = df_clean[hari_cols].notna().any(axis=1)
            df_clean = df_clean[has_schedule]
        
        return df_clean
    
    @staticmethod
    def _fix_time_format(time_str: Any) -> str:
        """Perbaiki format waktu yang tidak standar"""
        if pd.isna(time_str):
            return ""
        
        str_time = str(time_str).strip()
        
        # Hapus karakter yang tidak perlu
        str_time = re.sub(r'[^\d.\-:]', '', str_time)
        
        # Ganti titik dengan titik dua untuk format jam
        if '.' in str_time and ':' not in str_time:
            parts = str_time.split('-')
            if len(parts) == 2:
                start = parts[0].replace('.', ':')
                end = parts[1].replace('.', ':')
                return f"{start}-{end}"
        
        return str_time


class ExcelStyler:
    """Kelas untuk styling Excel"""
    
    def __init__(self, config: Dict[str, Any]):
        self.config = config
        
    def apply_styles(self, ws, max_row: int):
        """Terapkan styling ke sheet Jadwal"""
        # Hitung jumlah E per hari per slot
        e_counts = self._count_poleks_per_slot(ws, max_row)
        
        # Terapkan warna dan highlight
        self._apply_colors_and_highlights(ws, max_row, e_counts)
    
    def _count_poleks_per_slot(self, ws, max_row: int) -> Dict[str, Dict[str, int]]:
        """Hitung jumlah Poleks per hari per slot"""
        e_counts = {hari: {slot: 0 for slot in self.config['time_slots_str']} 
                   for hari in ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]}
        
        for row in range(2, max_row + 1):
            hari = ws.cell(row=row, column=3).value
            if hari not in e_counts:
                continue
                
            for col_idx, slot in enumerate(self.config['time_slots_str'], start=5):
                cell = ws.cell(row=row, column=col_idx)
                if cell.value == 'E':
                    e_counts[hari][slot] += 1
        
        return e_counts
    
    def _apply_colors_and_highlights(self, ws, max_row: int, e_counts: Dict[str, Dict[str, int]]):
        """Terapkan warna dan highlight jika melebihi batas"""
        for row in range(2, max_row + 1):
            hari = ws.cell(row=row, column=3).value
            if hari not in e_counts:
                continue
                
            for col_idx, slot in enumerate(self.config['time_slots_str'], start=5):
                cell = ws.cell(row=row, column=col_idx)
                
                if cell.value == 'R':
                    cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
                elif cell.value == 'E':
                    cell.fill = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
                    
                    # Cek apakah melebihi batas
                    if e_counts[hari][slot] > self.config['max_poleks_per_slot']:
                        # Cari baris yang melebihi batas
                        e_rows_for_slot = []
                        for r in range(2, max_row + 1):
                            if (ws.cell(row=r, column=3).value == hari and 
                                ws.cell(row=r, column=col_idx).value == 'E'):
                                e_rows_for_slot.append(r)
                        
                        # Warnai baris yang melebihi batas
                        if len(e_rows_for_slot) > self.config['max_poleks_per_slot']:
                            if row in e_rows_for_slot[self.config['max_poleks_per_slot']:]:
                                cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")


class Visualizer:
    """Kelas untuk visualisasi data"""
    
    @staticmethod
    def create_heatmap(df: pd.DataFrame, time_slots: List[str]) -> go.Figure:
        """Buat heatmap jadwal"""
        if df.empty:
            return go.Figure()
        
        # Siapkan data untuk heatmap
        pivot_data = []
        
        for _, row in df.iterrows():
            for slot in time_slots:
                value = 0
                if row[slot] == 'R':
                    value = 1  # Reguler
                elif row[slot] == 'E':
                    value = 2  # Poleks
                
                pivot_data.append({
                    'Hari': row['HARI'],
                    'Waktu': slot,
                    'Poli': row['POLI ASAL'],
                    'Jenis': row['JENIS POLI'],
                    'Value': value,
                    'Label': row[slot] if row[slot] else ''
                })
        
        heatmap_df = pd.DataFrame(pivot_data)
        
        if heatmap_df.empty:
            return go.Figure()
        
        # Buat pivot table untuk heatmap
        pivot_table = heatmap_df.pivot_table(
            index='Hari',
            columns='Waktu',
            values='Value',
            aggfunc='max',
            fill_value=0
        )
        
        # Urutkan hari
        hari_order = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]
        pivot_table = pivot_table.reindex(hari_order)
        
        # Buat heatmap
        fig = px.imshow(
            pivot_table,
            labels=dict(x="Waktu", y="Hari", color="Status"),
            x=pivot_table.columns,
            y=pivot_table.index,
            aspect="auto",
            color_continuous_scale=["white", "green", "blue", "red"],
            text_auto=False
        )
        
        # Update layout
        fig.update_layout(
            title="Heatmap Jadwal Poli",
            xaxis_title="Waktu",
            yaxis_title="Hari",
            height=400
        )
        
        return fig
    
    @staticmethod
    def create_schedule_gantt(df: pd.DataFrame) -> go.Figure:
        """Buat Gantt chart untuk jadwal"""
        if df.empty:
            return go.Figure()
        
        # Siapkan data untuk Gantt chart
        gantt_data = []
        
        for _, row in df.iterrows():
            for i, slot in enumerate(row.index[4:]):  # Mulai dari kolom waktu
                if row[slot] in ['R', 'E']:
                    gantt_data.append({
                        'Task': f"{row['POLI ASAL']} - {row['DOKTER']}",
                        'Start': slot,
                        'Finish': slot,  # Akan diubah nanti
                        'Resource': row[slot],
                        'Hari': row['HARI']
                    })
        
        if not gantt_data:
            return go.Figure()
        
        gantt_df = pd.DataFrame(gantt_data)
        
        # Buat Gantt chart
        fig = px.timeline(
            gantt_df,
            x_start="Start",
            x_end="Finish",
            y="Task",
            color="Resource",
            facet_row="Hari",
            title="Gantt Chart Jadwal Poli"
        )
        
        fig.update_layout(
            height=600,
            showlegend=True,
            xaxis_title="Waktu",
            yaxis_title=""
        )
        
        return fig
    
    @staticmethod
    def create_statistics_charts(df: pd.DataFrame, time_slots: List[str]) -> Tuple[go.Figure, go.Figure]:
        """Buat chart statistik"""
        if df.empty:
            return go.Figure(), go.Figure()
        
        # Hitung statistik
        total_r = (df[time_slots] == 'R').sum().sum()
        total_e = (df[time_slots] == 'E').sum().sum()
        total_empty = (df[time_slots] == '').sum().sum()
        
        # Pie chart untuk distribusi slot
        fig_pie = go.Figure(data=[go.Pie(
            labels=['Reguler', 'Poleks', 'Kosong'],
            values=[total_r, total_e, total_empty],
            hole=.3,
            marker_colors=['green', 'blue', 'lightgray']
        )])
        
        fig_pie.update_layout(
            title="Distribusi Slot Waktu",
            height=300
        )
        
        # Bar chart untuk slot per hari
        hari_stats = []
        for hari in ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]:
            hari_df = df[df['HARI'] == hari]
            if not hari_df.empty:
                hari_r = (hari_df[time_slots] == 'R').sum().sum()
                hari_e = (hari_df[time_slots] == 'E').sum().sum()
                hari_stats.append({
                    'Hari': hari,
                    'Reguler': hari_r,
                    'Poleks': hari_e
                })
        
        stats_df = pd.DataFrame(hari_stats)
        
        fig_bar = go.Figure(data=[
            go.Bar(name='Reguler', x=stats_df['Hari'], y=stats_df['Reguler'], marker_color='green'),
            go.Bar(name='Poleks', x=stats_df['Hari'], y=stats_df['Poleks'], marker_color='blue')
        ])
        
        fig_bar.update_layout(
            title="Slot per Hari",
            barmode='stack',
            height=300,
            xaxis_title="Hari",
            yaxis_title="Jumlah Slot"
        )
        
        return fig_pie, fig_bar


# ============================================================================
# KELAS TAMPILAN (VIEW)
# ============================================================================

class TemplateManager:
    """Kelas untuk mengelola template Excel"""
    
    @staticmethod
    def create_template() -> io.BytesIO:
        """Buat template Excel"""
        wb = openpyxl.Workbook()
        
        # Sheet Poli Asal
        ws1 = wb.active
        ws1.title = "Poli Asal"
        ws1.append(["No", "Nama Poli", "kode sheet"])
        poli_list = [
            ["1", "Poli Anak", "ANAK"],
            ["2", "Poli Bedah", "BEDAH"],
            ["3", "Poli Dalam", "DALAM"],
            ["4", "Poli Obgyn", "OBGYN"],
            ["5", "Poli Jantung", "JANTUNG"],
            ["6", "Poli Ortho", "ORTHO"],
            ["7", "Poli Paru", "PARU"],
            ["8", "Poli Saraf", "SARAF"],
            ["9", "Poli THT", "THT"],
            ["10", "Poli Urologi", "URO"],
            ["11", "Poli Jiwa", "JIWA"],
            ["12", "Poli Kukel", "KUKEL"],
            ["13", "Poli Bedah Saraf", "BSARAF"],
            ["14", "Poli Gigi", "GIGI"],
            ["15", "Poli Mata", "MATA"],
            ["16", "Poli Rehab", "REHAB"]
        ]
        for poli in poli_list:
            ws1.append(poli)
        
        # Sheet Reguler
        ws2 = wb.create_sheet("Reguler")
        ws2.append(["Nama Dokter", "Poli Asal", "Jenis Poli", "Senin", "Selasa", 
                    "Rabu", "Kamis", "Jum'at"])
        ws2.append(["dr. Contoh Dokter, Sp.A", "Poli Anak", "Reguler", 
                    "08.00 - 10.30", "", "08.00 - 10.30", "", ""])
        
        # Sheet Poleks
        ws3 = wb.create_sheet("Poleks")
        ws3.append(["Nama Dokter", "Poli Asal", "Jenis Poli", "Senin", "Selasa", 
                    "Rabu", "Kamis", "Jum'at"])
        ws3.append(["dr. Contoh Dokter, Sp.A", "Poli Anak", "Poleks", 
                    "07.30 - 08.25", "", "07.30 - 08.25", "", ""])
        
        # Sheet Jadwal
        ws4 = wb.create_sheet("Jadwal")
        time_slots_str = [t.strftime("%H:%M") for t in [
            time(7, 30), time(8, 0), time(8, 30), time(9, 0), time(9, 30),
            time(10, 0), time(10, 30), time(11, 0), time(11, 30), time(12, 0),
            time(12, 30), time(13, 0), time(13, 30), time(14, 0), time(14, 30)
        ]]
        ws4.append(["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + time_slots_str)
        
        # Simpan ke buffer
        buffer = io.BytesIO()
        wb.save(buffer)
        buffer.seek(0)
        
        return buffer


class FileProcessor:
    """Kelas untuk memproses file"""
    
    def __init__(self, app: 'PoliSchedulerApp'):
        self.app = app
        self.error_analyzer = ErrorAnalyzer()
        
    def process_uploaded_file(self, uploaded_file, progress_bar) -> Optional[io.BytesIO]:
        """Proses file yang diupload"""
        try:
            # Langkah 1: Validasi file
            progress_bar.progress(10)
            validation_result = self._validate_file(uploaded_file)
            
            if not validation_result['is_valid']:
                raise ValueError(f"File tidak valid: {validation_result['errors']}")
            
            # Langkah 2: Analisis error
            progress_bar.progress(20)
            error_reports = self._analyze_sheets(uploaded_file)
            
            # Simpan error reports ke session state
            st.session_state.error_logs = error_reports
            
            # Langkah 3: Baca data
            progress_bar.progress(40)
            df_reguler, df_poleks = self._read_data(uploaded_file)
            
            # Langkah 4: Proses jadwal
            progress_bar.progress(60)
            df_jadwal = self._process_schedules(df_reguler, df_poleks)
            
            # Simpan ke session state
            st.session_state.processed_data = df_jadwal
            
            # Langkah 5: Buat visualisasi
            progress_bar.progress(70)
            st.session_state.visualization_data = df_jadwal
            
            # Langkah 6: Buat Excel dengan styling
            progress_bar.progress(80)
            result_buffer = self._create_styled_excel(uploaded_file, df_jadwal)
            
            progress_bar.progress(100)
            return result_buffer
            
        except Exception as e:
            st.error(f"❌ Error dalam memproses file: {str(e)}")
            st.error(traceback.format_exc())
            return None
    
    def _validate_file(self, uploaded_file) -> Dict[str, Any]:
        """Validasi file"""
        try:
            wb = load_workbook(uploaded_file)
            required_sheets = ['Reguler', 'Poleks']
            missing_sheets = [s for s in required_sheets if s not in wb.sheetnames]
            
            if missing_sheets:
                return {
                    'is_valid': False,
                    'errors': [f"Sheet berikut tidak ditemukan: {', '.join(missing_sheets)}"]
                }
            
            return {'is_valid': True, 'errors': []}
            
        except Exception as e:
            return {'is_valid': False, 'errors': [f"Error membaca file: {str(e)}"]}
    
    def _analyze_sheets(self, uploaded_file) -> List[Dict[str, Any]]:
        """Analisis sheet untuk error"""
        reports = []
        
        try:
            # Analisis sheet Reguler
            df_reguler = pd.read_excel(uploaded_file, sheet_name='Reguler')
            reguler_report = self.error_analyzer.analyze_excel_structure(df_reguler, 'Reguler')
            reports.append({'sheet': 'Reguler', **reguler_report})
            
            # Analisis sheet Poleks
            df_poleks = pd.read_excel(uploaded_file, sheet_name='Poleks')
            poleks_report = self.error_analyzer.analyze_excel_structure(df_poleks, 'Poleks')
            reports.append({'sheet': 'Poleks', **poleks_report})
            
        except Exception as e:
            reports.append({
                'sheet': 'Unknown',
                'is_valid': False,
                'errors': [f"Error dalam analisis: {str(e)}"]
            })
        
        return reports
    
    def _read_data(self, uploaded_file) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Baca data dari file"""
        df_reguler = pd.read_excel(uploaded_file, sheet_name='Reguler')
        df_poleks = pd.read_excel(uploaded_file, sheet_name='Poleks')
        
        return df_reguler, df_poleks
    
    def _process_schedules(self, df_reguler: pd.DataFrame, df_poleks: pd.DataFrame) -> pd.DataFrame:
        """Proses jadwal dari kedua sheet"""
        # Update konfigurasi
        time_slots = self._generate_time_slots()
        time_slots_str = [t.strftime("%H:%M") for t in time_slots]
        
        config = {
            'time_slots': time_slots,
            'time_slots_str': time_slots_str,
            'max_poleks_per_slot': self.app.config['max_poleks_per_slot']
        }
        
        # Proses data
        processor = DataProcessor(config)
        
        df_reguler_processed = processor.process_schedule(df_reguler, 'Reguler')
        df_poleks_processed = processor.process_schedule(df_poleks, 'Poleks')
        
        # Gabungkan hasil
        df_jadwal = pd.concat([df_reguler_processed, df_poleks_processed], ignore_index=True)
        
        # Urutkan
        df_jadwal['HARI_ORDER'] = df_jadwal['HARI'].map(self.app.HARI_ORDER)
        df_jadwal = df_jadwal.sort_values(['POLI ASAL', 'HARI_ORDER', 'DOKTER', 'JENIS POLI'])
        df_jadwal = df_jadwal.drop('HARI_ORDER', axis=1)
        
        return df_jadwal.reset_index(drop=True)
    
    def _generate_time_slots(self) -> List[time]:
        """Generate time slots berdasarkan konfigurasi"""
        slots = []
        current_time = time(
            self.app.config['start_hour'],
            self.app.config['start_minute']
        )
        
        interval = self.app.config['interval_minutes']
        end_time = time(14, 30)  # Batas akhir
        
        while current_time <= end_time:
            slots.append(current_time)
            
            # Tambah interval
            current_datetime = datetime.combine(datetime.today(), current_time)
            next_datetime = current_datetime + timedelta(minutes=interval)
            current_time = next_datetime.time()
        
        return slots
    
    def _create_styled_excel(self, uploaded_file, df_jadwal: pd.DataFrame) -> io.BytesIO:
        """Buat file Excel dengan styling"""
        new_wb = load_workbook(uploaded_file)
        
        # Hapus sheet Jadwal yang lama jika ada
        if 'Jadwal' in new_wb.sheetnames:
            std = new_wb['Jadwal']
            new_wb.remove(std)
        
        # Buat sheet Jadwal baru
        ws_jadwal = new_wb.create_sheet('Jadwal')
        
        # Tulis header
        time_slots_str = [t.strftime("%H:%M") for t in self._generate_time_slots()]
        headers = ['POLI ASAL', 'JENIS POLI', 'HARI', 'DOKTER'] + time_slots_str
        for col_idx, header in enumerate(headers, start=1):
            ws_jadwal.cell(row=1, column=col_idx, value=header)
        
        # Tulis data
        for row_idx, row_data in enumerate(df_jadwal.to_dict('records'), start=2):
            ws_jadwal.cell(row=row_idx, column=1, value=row_data['POLI ASAL'])
            ws_jadwal.cell(row=row_idx, column=2, value=row_data['JENIS POLI'])
            ws_jadwal.cell(row=row_idx, column=3, value=row_data['HARI'])
            ws_jadwal.cell(row=row_idx, column=4, value=row_data['DOKTER'])
            
            for col_idx, slot in enumerate(time_slots_str, start=5):
                ws_jadwal.cell(row=row_idx, column=col_idx, value=row_data.get(slot, ''))
        
        # Terapkan styling
        config = {
            'time_slots_str': time_slots_str,
            'max_poleks_per_slot': self.app.config['max_poleks_per_slot']
        }
        styler = ExcelStyler(config)
        styler.apply_styles(ws_jadwal, len(df_jadwal) + 1)
        
        # Simpan ke buffer
        result_buffer = io.BytesIO()
        new_wb.save(result_buffer)
        result_buffer.seek(0)
        
        return result_buffer


# ============================================================================
# TAMPILAN STREAMLIT
# ============================================================================

class StreamlitUI:
    """Kelas untuk mengelola tampilan Streamlit"""
    
    def __init__(self, app: 'PoliSchedulerApp'):
        self.app = app
        self.template_manager = TemplateManager()
        self.file_processor = FileProcessor(app)
        self.visualizer = Visualizer()
        
    def render_sidebar(self):
        """Render sidebar"""
        with st.sidebar:
            st.title("🏥 Pengisi Jadwal Poli OOP")
            st.markdown("---")
            
            st.subheader("📋 Panduan")
            st.markdown("""
            1. **Download** template untuk format yang benar
            2. **Upload** file Excel yang sudah diisi
            3. **Analisis** error dan perbaiki jika perlu
            4. **Proses** dan **download** hasil
            """)
            
            st.markdown("---")
            
            # Download template
            if st.button("📥 Download Template", use_container_width=True):
                template_buffer = self.template_manager.create_template()
                st.download_button(
                    label="Klik untuk download template",
                    data=template_buffer,
                    file_name="template_jadwal_poli_oop.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            st.markdown("---")
            
            # Auto-fix toggle
            self.app.config['auto_fix_errors'] = st.checkbox(
                "🔄 Auto-fix Errors",
                value=self.app.config['auto_fix_errors'],
                help="Secara otomatis memperbaiki error format yang umum"
            )
            
            st.markdown("---")
            
            # Info kontak/help
            st.caption("❓ Butuh bantuan?")
            st.caption("📧 support@example.com")
    
    def render_main_content(self):
        """Render konten utama"""
        st.title("🏥 Pengisi Jadwal Poli Excel - OOP Version")
        st.caption("Aplikasi berbasis OOP untuk mengisi jadwal poli secara otomatis")
        
        # Tabs
        tab1, tab2, tab3, tab4 = st.tabs([
            "📤 Upload & Proses", 
            "🔍 Error Analyzer", 
            "📊 Visualisasi", 
            "⚙️ Pengaturan"
        ])
        
        with tab1:
            self.render_upload_tab()
        
        with tab2:
            self.render_error_analyzer_tab()
        
        with tab3:
            self.render_visualization_tab()
        
        with tab4:
            self.render_settings_tab()
    
    def render_upload_tab(self):
        """Render tab upload"""
        col1, col2 = st.columns([2, 1])
        
        with col1:
            uploaded_file = st.file_uploader(
                "Upload file Excel (.xlsx)", 
                type=['xlsx'],
                help="Upload file dengan format yang sesuai",
                key="file_uploader"
            )
        
        with col2:
            if uploaded_file:
                file_size = len(uploaded_file.getvalue()) / 1024
                st.metric("📏 Ukuran File", f"{file_size:.1f} KB")
                
                try:
                    excel_file = pd.ExcelFile(uploaded_file)
                    st.success(f"✅ {len(excel_file.sheet_names)} sheet ditemukan")
                except:
                    st.error("❌ File tidak valid")
        
        if uploaded_file:
            st.markdown("---")
            
            # Preview sheet
            with st.expander("📄 Preview Sheet", expanded=False):
                sheet_names = pd.ExcelFile(uploaded_file).sheet_names
                selected_sheet = st.selectbox("Pilih sheet untuk preview:", sheet_names)
                
                if selected_sheet:
                    try:
                        df_preview = pd.read_excel(uploaded_file, sheet_name=selected_sheet, nrows=10)
                        st.dataframe(df_preview, use_container_width=True)
                    except:
                        st.warning(f"Tidak dapat membaca sheet {selected_sheet}")
            
            # Tombol proses
            col_proses1, col_proses2, col_proses3 = st.columns([1, 2, 1])
            with col_proses2:
                if st.button("🚀 Proses Jadwal", type="primary", use_container_width=True):
                    with st.spinner("Memproses..."):
                        progress_bar = st.progress(0)
                        
                        result_buffer = self.file_processor.process_uploaded_file(
                            uploaded_file, progress_bar
                        )
                        
                        if result_buffer:
                            progress_bar.progress(100)
                            st.success("✅ File berhasil diproses!")
                            
                            # Tombol download
                            st.download_button(
                                label="📥 Download Hasil",
                                data=result_buffer,
                                file_name="jadwal_hasil_oop.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
    
    def render_error_analyzer_tab(self):
        """Render tab error analyzer"""
        st.subheader("🔍 Error Analyzer")
        
        if 'error_logs' in st.session_state and st.session_state.error_logs:
            for error_report in st.session_state.error_logs:
                with st.expander(f"📋 Laporan untuk sheet: {error_report['sheet']}", expanded=True):
                    report_html = ErrorAnalyzer.create_error_report(error_report)
                    st.markdown(report_html, unsafe_allow_html=True)
                    
                    # Tampilkan rekomendasi perbaikan
                    if error_report['errors'] or error_report['warnings']:
                        st.markdown("### 🛠️ Rekomendasi Perbaikan:")
                        
                        for error in error_report['errors']:
                            if "kolom yang diperlukan" in error.lower():
                                st.markdown("- **Tambahkan kolom yang hilang** sesuai template")
                            elif "format waktu" in error.lower():
                                st.markdown("- **Perbaiki format waktu** menggunakan format HH.MM - HH.MM")
                        
                        if self.app.config['auto_fix_errors']:
                            st.success("✅ Auto-fix akan memperbaiki error format secara otomatis")
        else:
            st.info("ℹ️ Belum ada data error yang dianalisis. Upload file terlebih dahulu.")
    
    def render_visualization_tab(self):
        """Render tab visualisasi"""
        st.subheader("📊 Visualisasi Jadwal")
        
        if (st.session_state.processed_data is not None and 
            not st.session_state.processed_data.empty):
            
            df = st.session_state.processed_data
            time_slots_str = [col for col in df.columns if col not in 
                            ['POLI ASAL', 'JENIS POLI', 'HARI', 'DOKTER']]
            
            # Pilihan visualisasi
            viz_type = st.selectbox(
                "Pilih jenis visualisasi:",
                ["Heatmap", "Gantt Chart", "Statistik", "Tabel Interaktif"]
            )
            
            if viz_type == "Heatmap":
                fig = self.visualizer.create_heatmap(df, time_slots_str)
                st.plotly_chart(fig, use_container_width=True)
                
                # Legenda
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.markdown("⬜ **Kosong**")
                with col2:
                    st.markdown("🟩 **Reguler**")
                with col3:
                    st.markdown("🟦 **Poleks**")
                with col4:
                    st.markdown("🟥 **Over Limit**")
            
            elif viz_type == "Gantt Chart":
                fig = self.visualizer.create_schedule_gantt(df)
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.warning("Data tidak cukup untuk Gantt Chart")
            
            elif viz_type == "Statistik":
                fig_pie, fig_bar = self.visualizer.create_statistics_charts(df, time_slots_str)
                
                col1, col2 = st.columns(2)
                with col1:
                    st.plotly_chart(fig_pie, use_container_width=True)
                with col2:
                    st.plotly_chart(fig_bar, use_container_width=True)
                
                # Tampilkan statistik detail
                st.markdown("### 📈 Statistik Detail")
                total_r = (df[time_slots_str] == 'R').sum().sum()
                total_e = (df[time_slots_str] == 'E').sum().sum()
                total_slots = len(df) * len(time_slots_str)
                
                col_stat1, col_stat2, col_stat3 = st.columns(3)
                with col_stat1:
                    st.metric("Total Slot", total_slots)
                with col_stat2:
                    st.metric("Slot Reguler", total_r, 
                             f"{(total_r/total_slots*100):.1f}%" if total_slots > 0 else "0%")
                with col_stat3:
                    st.metric("Slot Poleks", total_e,
                             f"{(total_e/total_slots*100):.1f}%" if total_slots > 0 else "0%")
            
            elif viz_type == "Tabel Interaktif":
                # Filter data
                col_filter1, col_filter2, col_filter3 = st.columns(3)
                with col_filter1:
                    selected_poli = st.multiselect(
                        "Filter Poli:",
                        options=df['POLI ASAL'].unique(),
                        default=df['POLI ASAL'].unique()[:3]
                    )
                with col_filter2:
                    selected_hari = st.multiselect(
                        "Filter Hari:",
                        options=df['HARI'].unique(),
                        default=df['HARI'].unique()
                    )
                with col_filter3:
                    selected_jenis = st.multiselect(
                        "Filter Jenis:",
                        options=df['JENIS POLI'].unique(),
                        default=df['JENIS POLI'].unique()
                    )
                
                # Terapkan filter
                filtered_df = df[
                    df['POLI ASAL'].isin(selected_poli) &
                    df['HARI'].isin(selected_hari) &
                    df['JENIS POLI'].isin(selected_jenis)
                ]
                
                # Format tabel dengan warna
                def color_cells(val):
                    if val == 'R':
                        return 'background-color: green; color: white'
                    elif val == 'E':
                        return 'background-color: blue; color: white'
                    elif val == '':
                        return 'background-color: lightgray'
                    else:
                        return ''
                
                # Tampilkan tabel
                st.dataframe(
                    filtered_df.style.applymap(color_cells, subset=time_slots_str),
                    use_container_width=True,
                    height=400
                )
                
                # Download filtered data
                csv = filtered_df.to_csv(index=False).encode('utf-8')
                st.download_button(
                    label="📥 Download Data Tersaring (CSV)",
                    data=csv,
                    file_name="jadwal_tersaring.csv",
                    mime="text/csv"
                )
        else:
            st.info("ℹ️ Belum ada data yang diproses. Upload dan proses file terlebih dahulu.")
    
    def render_settings_tab(self):
        """Render tab pengaturan"""
        st.subheader("⚙️ Pengaturan Aplikasi")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### ⏰ Pengaturan Waktu")
            
            self.app.config['start_hour'] = st.slider(
                "Jam mulai",
                min_value=5,
                max_value=12,
                value=self.app.config['start_hour'],
                help="Jam mulai jadwal"
            )
            
            self.app.config['start_minute'] = st.slider(
                "Menit mulai",
                min_value=0,
                max_value=59,
                value=self.app.config['start_minute'],
                step=5
            )
            
            self.app.config['interval_minutes'] = st.selectbox(
                "Interval (menit)",
                options=[15, 20, 30, 60],
                index=[15, 20, 30, 60].index(self.app.config['interval_minutes'])
                if self.app.config['interval_minutes'] in [15, 20, 30, 60] else 2
            )
        
        with col2:
            st.markdown("### 📊 Pengaturan Batasan")
            
            self.app.config['max_poleks_per_slot'] = st.number_input(
                "Batas maksimal Poleks per slot",
                min_value=1,
                max_value=20,
                value=self.app.config['max_poleks_per_slot'],
                help="Jika lebih dari batas ini, akan ditandai merah"
            )
            
            st.markdown("### 🎨 Pengaturan Tampilan")
            
            # Pilihan tema warna (untuk visualisasi)
            color_theme = st.selectbox(
                "Tema warna visualisasi",
                ["Default", "Hijau-Biru", "Merah-Kuning", "Pastel"]
            )
            
            if color_theme != "Default":
                st.info(f"Tema {color_theme} akan diterapkan pada visualisasi")
        
        # Preview time slots
        st.markdown("### 👁️ Preview Time Slots")
        time_slots = self.file_processor._generate_time_slots()
        time_slots_str = [t.strftime("%H:%M") for t in time_slots]
        
        col_slots = st.columns(min(6, len(time_slots_str)))
        for i, slot in enumerate(time_slots_str):
            with col_slots[i % len(col_slots)]:
                st.info(f"**Slot {i+1}:** {slot}")
        
        st.info("⚙️ Pengaturan akan diterapkan pada proses selanjutnya")
        
        # Reset button
        if st.button("🔄 Reset ke Default", type="secondary"):
            self.app.initialize_constants()
            st.success("✅ Pengaturan telah direset ke default")
            st.rerun()


# ============================================================================
# MAIN EXECUTION
# ============================================================================

def main():
    """Fungsi utama"""
    # Inisialisasi aplikasi
    app = PoliSchedulerApp()
    
    # Inisialisasi UI
    ui = StreamlitUI(app)
    
    # Render aplikasi
    ui.render_sidebar()
    ui.render_main_content()


if __name__ == "__main__":
    main()
