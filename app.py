import streamlit as st
import pandas as pd
import openpyxl
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from datetime import time, timedelta, datetime
import io
import re

# Konfigurasi halaman
st.set_page_config(
    page_title="Pengisi Jadwal Poli Excel",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS kustom
st.markdown("""
<style>
    .stButton > button {
        width: 100%;
        margin-top: 10px;
    }
    .success-box {
        padding: 20px;
        background-color: #d4edda;
        border-radius: 5px;
        border-left: 5px solid #155724;
        margin: 10px 0;
    }
    .warning-box {
        padding: 20px;
        background-color: #fff3cd;
        border-radius: 5px;
        border-left: 5px solid #856404;
        margin: 10px 0;
    }
</style>
""", unsafe_allow_html=True)

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
            create_template()
            
        st.markdown("---")
        
        # Info kontak/help
        st.caption("❓ Butuh bantuan?")
        st.caption("📧 contact@example.com")
    
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
                st.metric("File Size", f"{len(uploaded_file.getvalue()) / 1024:.1f} KB")
                
                # Validasi cepat
                try:
                    excel_file = pd.ExcelFile(uploaded_file)
                    st.success(f"✅ {len(excel_file.sheet_names)} sheet ditemukan")
                except:
                    st.error("❌ File tidak valid")
        
        if uploaded_file:
            st.markdown("---")
            
            # Preview sheet
            with st.expander("📄 Preview Sheet", expanded=True):
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
                            result = process_file(uploaded_file, progress_bar)
                            
                            if result:
                                progress_bar.progress(100)
                                
                                # Tampilkan notifikasi sukses
                                st.markdown('<div class="success-box">✅ File berhasil diproses!</div>', 
                                          unsafe_allow_html=True)
                                
                                # Tombol download
                                st.download_button(
                                    label="📥 Download Hasil",
                                    data=result,
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
    
    with tab3:
        if 'processed_data' in st.session_state:
            st.dataframe(st.session_state.processed_data, use_container_width=True)
        else:
            st.info("Belum ada data yang diproses")

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
        # ... tambahkan lainnya
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
    ws4.append(["POLI ASAL", "JENIS POLI", "HARI", "DOKTER"] + 
               [f"{h:02d}:{m:02d}" for h in range(7, 15) for m in [30, 0] if not (h == 14 and m == 0)][:15])
    
    # Simpan ke buffer
    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    # Download button
    st.download_button(
        label="Download Template",
        data=buffer,
        file_name="template_jadwal_poli.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def process_file(uploaded_file, progress_bar):
    """Proses file Excel"""
    # ... implementasi processing seperti sebelumnya
    # dengan progress bar updates
    progress_bar.progress(25)
    # Baca file
    # ...
    progress_bar.progress(50)
    # Proses data
    # ...
    progress_bar.progress(75)
    # Buat output
    # ...
    
    return result_buffer

if __name__ == "__main__":
    main()
