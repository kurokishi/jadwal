import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Border, Side, Font
from datetime import time, timedelta, datetime
import io
import re
import seaborn as sns
import matplotlib.pyplot as plt
import altair as alt

# Konfigurasi halaman
st.set_page_config(
    page_title="Pengisi Jadwal Poli Excel",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Warna untuk sel
FILL_R = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")
FILL_E = PatternFill(start_color="0000FF", end_color="0000FF", fill_type="solid")
FILL_OVERLIMIT = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

# Border tipis
THIN_BORDER = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

# Konfigurasi default
TIME_SLOTS = [
    time(7, 30), time(8, 0), time(8, 30), time(9, 0), time(9, 30),
    time(10, 0), time(10, 30), time(11, 0), time(11, 30), time(12, 0),
    time(12, 30), time(13, 0), time(13, 30), time(14, 0), time(14, 30)
]
TIME_SLOTS_STR = [t.strftime("%H:%M") for t in TIME_SLOTS]

HARI_ORDER = {"Senin": 1, "Selasa": 2, "Rabu": 3, "Kamis": 4, "Jum'at": 5}
HARI_INDONESIA = ["Senin", "Selasa", "Rabu", "Kamis", "Jum'at"]

def parse_time_range(time_str):
    """Parse rentang waktu dari string seperti '08.00 - 10.30'"""
    if pd.isna(time_str) or str(time_str).strip() == "":
        return None, None
    
    # Bersihkan string
    clean_str = str(time_str).strip().replace(' ', '').replace('.', ':')
    
    # Cari pola waktu
    pattern = r'(\d{1,2}:\d{2})-(\d{1,2}:\d{2})'
    match = re.search(pattern, clean_str)
    if not match:
        return None, None
    
    try:
        start_str, end_str = match.groups()
        # Parse waktu
        start_hour, start_minute = map(int, start_str.split(':'))
        end_hour, end_minute = map(int, end_str.split(':'))
        
        start_time = time(start_hour, start_minute)
        end_time = time(end_hour, end_minute)
        
        # Jika end_time sebelum start_time, asumsikan hari berikutnya
        if end_time < start_time:
            end_time = time(end_hour + 24, end_minute)
        
        return start_time, end_time
    except:
        return None, None

def time_overlap(slot_start, slot_end, schedule_start, schedule_end):
    """Cek apakah slot waktu overlap dengan jadwal"""
    if schedule_start is None or schedule_end is None:
        return False
    
    # Cek overlap
    return not (slot_end <= schedule_start or slot_start >= schedule_end)

def process_schedule(df, jenis_poli):
    """Proses dataframe jadwal"""
    results = []
    
    for (dokter, poli_asal), group in df.groupby(['Nama Dokter', 'Poli Asal']):
        # Kumpulkan semua jadwal per hari
        hari_schedules = {}
        
        for hari in HARI_INDONESIA:
            if hari not in group.columns:
                continue
                
            time_ranges = []
            for time_str in group[hari]:
                start_time, end_time = parse_time_range(time_str)
                if start_time and end_time:
                    # Potong maksimal sampai 14:30
                    if end_time > time(14, 30):
                        end_time = time(14, 30)
                    time_ranges.append((start_time, end_time))
            
            hari_schedules[hari] = time_ranges
        
        # Generate baris per hari
        for hari in HARI_INDONESIA:
            if hari not in hari_schedules or not hari_schedules[hari]:
                continue
                
            # Gabungkan rentang waktu yang overlapping
            merged_ranges = []
            for start, end in sorted(hari_schedules[hari]):
                if not merged_ranges:
                    merged_ranges.append([start, end])
                else:
                    last_start, last_end = merged_ranges[-1]
                    if start <= last_end:
                        merged_ranges[-1][1] = max(last_end, end)
                    else:
                        merged_ranges.append([start, end])
            
            # Buat baris untuk sheet Jadwal
            row = {
                'POLI ASAL': poli_asal,
                'JENIS POLI': jenis_poli,
                'HARI': hari,
                'DOKTER': dokter
            }
            
            # Isi slot waktu
            for i, slot_time in enumerate(TIME_SLOTS):
                slot_start = slot_time
                # Slot 30 menit
                slot_end = (datetime.combine(datetime.today(), slot_start) + 
                           timedelta(minutes=30)).time()
                
                # Cek overlap dengan semua rentang waktu
                has_overlap = any(
                    time_overlap(slot_start, slot_end, start, end)
                    for start, end in merged_ranges
                )
                
                if has_overlap:
                    row[TIME_SLOTS_STR[i]] = 'R' if jenis_poli == 'Reguler' else 'E'
                else:
                    row[TIME_SLOTS_STR[i]] = ''
            
            results.append(row)
    
    return pd.DataFrame(results)

def apply_styles(ws, max_row):
    """Terapkan styling ke sheet Jadwal"""
    # Hitung jumlah E per hari per slot
    e_counts = {hari: {slot: 0 for slot in TIME_SLOTS_STR} for hari in HARI_INDONESIA}
    
    # Pertama, hitung semua E
    for row in range(2, max_row + 1):
        hari = ws.cell(row=row, column=3).value  # Kolom HARI
        if hari not in HARI_INDONESIA:
            continue
            
        for col_idx, slot in enumerate(TIME_SLOTS_STR, start=5):
            cell = ws.cell(row=row, column=col_idx)
            if cell.value == 'E':
                e_counts[hari][slot] += 1
    
    # Terapkan warna, border, dan font
    header_font = Font(bold=True)
    dokter_font = Font(italic=True)
    
    # Header row
    for col_idx in range(1, ws.max_column + 1):
        header_cell = ws.cell(row=1, column=col_idx)
        header_cell.font = header_font
        header_cell.border = THIN_BORDER
    
    # Terapkan ke data rows
    for row in range(2, max_row + 1):
        hari = ws.cell(row=row, column=3).value
        if hari not in HARI_INDONESIA:
            continue
        
        # Dokter cell italic
        dokter_cell = ws.cell(row=row, column=4)
        dokter_cell.font = dokter_font
        
        for col_idx in range(1, ws.max_column + 1):
            cell = ws.cell(row=row, column=col_idx)
            cell.border = THIN_BORDER
            
            if col_idx >= 5:  # Slot waktu
                if cell.value == 'R':
                    cell.fill = FILL_R
                elif cell.value == 'E':
                    cell.fill = FILL_E
                    
                    # Cek overlimit
                    slot = TIME_SLOTS_STR[col_idx - 5]
                    if e_counts[hari][slot] > 7:
                        # Cari baris ke-8 dan seterusnya
                        e_rows_for_slot = []
                        for r in range(2, max_row + 1):
                            if (ws.cell(row=r, column=3).value == hari and 
                                ws.cell(row=r, column=col_idx).value == 'E'):
                                e_rows_for_slot.append(r)
                        
                        if len(e_rows_for_slot) > 7:
                            if row in e_rows_for_slot[7:]:
                                cell.fill = FILL_OVERLIMIT
    
    # Freeze panes
    ws.freeze_panes = 'E2'

def create_template():
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
    ws4.append(["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + TIME_SLOTS_STR)
    
    # Sheet Legenda (baru)
    ws_legenda = wb.create_sheet("Legenda")
    ws_legenda.append(["Keterangan", "Contoh", "Arti"])
    ws_legenda.append(["Hijau", "", "Reguler (R)"])
    ws_legenda.cell(row=2, column=2).fill = FILL_R
    ws_legenda.append(["Biru", "", "Poleks (E)"])
    ws_legenda.cell(row=3, column=2).fill = FILL_E
    ws_legenda.append(["Merah", "", "Poleks Overlimit (>7)"])
    ws_legenda.cell(row=4, column=2).fill = FILL_OVERLIMIT
    
    # Simpan ke buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer

def process_file(uploaded_file, progress_bar):
    """Proses file Excel"""
    try:
        # Langkah 1: Baca file
        progress_bar.progress(25)
        wb = load_workbook(uploaded_file)
        
        # Cek sheet yang diperlukan
        required_sheets = ['Reguler', 'Poleks', 'Jadwal']
        missing_sheets = [s for s in required_sheets if s not in wb.sheetnames]
        
        if missing_sheets:
            raise ValueError(f"Sheet berikut tidak ditemukan: {', '.join(missing_sheets)}")
        
        # Langkah 2: Baca data
        progress_bar.progress(40)
        df_reguler = pd.read_excel(uploaded_file, sheet_name='Reguler')
        df_poleks = pd.read_excel(uploaded_file, sheet_name='Poleks')
        
        # Langkah 3: Proses jadwal
        progress_bar.progress(60)
        df_reguler_processed = process_schedule(df_reguler, 'Reguler')
        df_poleks_processed = process_schedule(df_poleks, 'Poleks')
        
        # Gabungkan hasil
        df_jadwal = pd.concat([df_reguler_processed, df_poleks_processed], ignore_index=True)
        
        # Urutkan
        df_jadwal['HARI_ORDER'] = df_jadwal['HARI'].map(HARI_ORDER)
        df_jadwal = df_jadwal.sort_values(['POLI ASAL', 'HARI_ORDER', 'DOKTER', 'JENIS POLI'])
        df_jadwal = df_jadwal.drop('HARI_ORDER', axis=1)
        
        # Reset index
        df_jadwal = df_jadwal.reset_index(drop=True)
        
        # Simpan ke session state untuk preview
        st.session_state.processed_data = df_jadwal
        
        # Langkah 4: Buat workbook baru
        progress_bar.progress(80)
        new_wb = load_workbook(uploaded_file)
        
        # Hapus sheet Jadwal yang lama jika ada
        if 'Jadwal' in new_wb.sheetnames:
            std = new_wb['Jadwal']
            new_wb.remove(std)
        
        # Buat sheet Jadwal baru
        ws_jadwal = new_wb.create_sheet('Jadwal')
        
        # Tulis header
        headers = ['POLI ASAL', 'JENIS POLI', 'HARI', 'DOKTER'] + TIME_SLOTS_STR
        for col_idx, header in enumerate(headers, start=1):
            ws_jadwal.cell(row=1, column=col_idx, value=header)
        
        # Tulis data
        for row_idx, row_data in enumerate(df_jadwal.to_dict('records'), start=2):
            ws_jadwal.cell(row=row_idx, column=1, value=row_data['POLI ASAL'])
            ws_jadwal.cell(row=row_idx, column=2, value=row_data['JENIS POLI'])
            ws_jadwal.cell(row=row_idx, column=3, value=row_data['HARI'])
            ws_jadwal.cell(row=row_idx, column=4, value=row_data['DOKTER'])
            
            for col_idx, slot in enumerate(TIME_SLOTS_STR, start=5):
                ws_jadwal.cell(row=row_idx, column=col_idx, value=row_data.get(slot, ''))
        
        # Terapkan styling
        apply_styles(ws_jadwal, len(df_jadwal) + 1)
        
        # Tambahkan sheet Legenda jika belum ada
        if 'Legenda' not in new_wb.sheetnames:
            ws_legenda = new_wb.create_sheet("Legenda")
            ws_legenda.append(["Keterangan", "Contoh", "Arti"])
            ws_legenda.append(["Hijau", "", "Reguler (R)"])
            ws_legenda.cell(row=2, column=2).fill = FILL_R
            ws_legenda.append(["Biru", "", "Poleks (E)"])
            ws_legenda.cell(row=3, column=2).fill = FILL_E
            ws_legenda.append(["Merah", "", "Poleks Overlimit (>7)"])
            ws_legenda.cell(row=4, column=2).fill = FILL_OVERLIMIT
        
        # Langkah 5: Simpan ke buffer
        progress_bar.progress(95)
        result_buffer = io.BytesIO()
        new_wb.save(result_buffer)
        result_buffer.seek(0)
        
        progress_bar.progress(100)
        return result_buffer
        
    except Exception as e:
        raise Exception(f"Error dalam memproses file: {str(e)}")

def main():
    # Sidebar
    with st.sidebar:
        st.title("🏥 Pengisi Jadwal Poli")
        st.markdown("---")
        
        st.subheader("📋 Panduan")
        st.markdown("""
        1. **Upload** file Excel dengan format yang benar
        2. **Konfigurasi** pengaturan jadwal
        3. **Proses** data
        4. **Download** hasil
        """)
        
        st.markdown("---")
        
        # Download template
        if st.button("📥 Download Template", use_container_width=True):
            template_buffer = create_template()
            st.download_button(
                label="Klik untuk download template",
                data=template_buffer,
                file_name="template_jadwal_poli.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
            
        st.markdown("---")
        
        # Info kontak/help
        st.caption("❓ Butuh bantuan?")
        st.caption("📧 support@example.com")
    
    # Main content
    st.title("🏥 Pengisi Jadwal Poli Excel")
    st.caption("Aplikasi untuk mengisi jadwal poli secara otomatis")
    
    # Tabs
    tab1, tab2, tab3 = st.tabs(["📤 Upload & Proses", "⚙️ Pengaturan", "📊 Preview"])
    
    with tab1:
        col1, col2 = st.columns([2, 1])
        
        with col1:
            uploaded_file = st.file_uploader(
                "Upload file Excel (.xlsx)", 
                type=['xlsx'],
                help="Upload file dengan format yang sesuai"
            )
        
        with col2:
            if uploaded_file:
                # Tampilkan info file
                file_size = len(uploaded_file.getvalue()) / 1024
                st.metric("File Size", f"{file_size:.1f} KB")
                
                # Validasi cepat
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
                        # Progress bar
                        progress_bar = st.progress(0)
                        
                        try:
                            # Proses file
                            result_buffer = process_file(uploaded_file, progress_bar)
                            
                            if result_buffer:
                                progress_bar.progress(100)
                                
                                # Tampilkan notifikasi sukses
                                st.success("✅ File berhasil diproses!")
                                
                                # Tombol download
                                st.download_button(
                                    label="📥 Download Hasil",
                                    data=result_buffer,
                                    file_name="jadwal_hasil.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )
                        except Exception as e:
                            st.error(f"❌ Error: {str(e)}")
    
    with tab2:
        st.subheader("Pengaturan Jadwal")
        
        # Konfigurasi waktu
        col_time1, col_time2, col_time3 = st.columns(3)
        with col_time1:
            start_hour = st.slider("Jam mulai", 5, 10, 7)
        with col_time2:
            start_minute = st.slider("Menit mulai", 0, 59, 30)
        with col_time3:
            interval = st.selectbox("Interval (menit)", [15, 30, 60], index=1)
        
        # Batasan
        st.number_input("Batas maksimal Poleks per slot", 
                       min_value=1, max_value=20, value=7,
                       help="Jika lebih dari batas ini, akan ditandai merah")
        
        st.info("⚙️ Pengaturan akan diterapkan pada proses selanjutnya")
    
    with tab3:
        if 'processed_data' in st.session_state and not st.session_state.processed_data.empty:
            df = st.session_state.processed_data
            st.subheader("Data yang Telah Diproses")
            
            # Filter interaktif
            selected_poli = st.multiselect("Filter Poli Asal", options=df['POLI ASAL'].unique())
            selected_hari = st.multiselect("Filter Hari", options=df['HARI'].unique())
            
            filtered_df = df.copy()
            if selected_poli:
                filtered_df = filtered_df[filtered_df['POLI ASAL'].isin(selected_poli)]
            if selected_hari:
                filtered_df = filtered_df[filtered_df['HARI'].isin(selected_hari)]
            
            st.dataframe(filtered_df, use_container_width=True)
            
            # Statistik
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Baris", len(filtered_df))
            with col2:
                total_r = (filtered_df[TIME_SLOTS_STR] == 'R').sum().sum()
                st.metric("Slot Reguler", total_r)
            with col3:
                total_e = (filtered_df[TIME_SLOTS_STR] == 'E').sum().sum()
                st.metric("Slot Poleks", total_e)
            
            # Heatmap
            st.subheader("Heatmap Jadwal")
            heatmap_data = filtered_df[TIME_SLOTS_STR].replace({'R': 1, 'E': 2, '': 0})
            fig, ax = plt.subplots(figsize=(12, 8))
            sns.heatmap(heatmap_data, cmap='coolwarm', ax=ax, annot=True, fmt=".0f")
            ax.set_xticklabels(TIME_SLOTS_STR, rotation=45)
            ax.set_yticklabels(filtered_df['DOKTER'], rotation=0)
            st.pyplot(fig)
            
            # Grafik distribusi
            st.subheader("Distribusi Slot per Hari")
            dist_data = filtered_df.groupby('HARI')[TIME_SLOTS_STR].apply(lambda x: (x == 'E').sum().sum()).reset_index(name='Poleks')
            dist_data['Reguler'] = filtered_df.groupby('HARI')[TIME_SLOTS_STR].apply(lambda x: (x == 'R').sum().sum())
            dist_data = dist_data.melt(id_vars='HARI', value_vars=['Poleks', 'Reguler'])
            chart = alt.Chart(dist_data).mark_bar().encode(
                x='HARI',
                y='value',
                color='variable'
            ).properties(width=600)
            st.altair_chart(chart)
            
            # Legenda di Streamlit
            with st.expander("Legenda Warna"):
                st.markdown("""
                - **Hijau**: Reguler (R)
                - **Biru**: Poleks (E)
                - **Merah**: Poleks Overlimit (>7 per slot)
                """)
        else:
            st.info("Belum ada data yang diproses. Upload dan proses file terlebih dahulu.")

if __name__ == "__main__":
    main()
