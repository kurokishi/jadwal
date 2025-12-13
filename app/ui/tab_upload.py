# app/ui/tab_upload.py
import streamlit as st
import pandas as pd
import io
import traceback
from datetime import datetime  # ✅ IMPORT datetime di sini

def render_upload_tab(scheduler, writer, analyzer, validator, config):
    st.subheader("📤 Upload & Proses Jadwal")
    
    # ======================================================
    # UPLOAD FILE SECTION (SELALU TAMPIL)
    # ======================================================
    uploaded_file = st.file_uploader(
        "Upload file Excel (Format: sheet Reguler & Poleks)",
        type=['xlsx', 'xls'],
        help="File harus memiliki sheet 'Reguler' dan 'Poleks'",
        key="file_uploader"
    )
    
    # ======================================================
    # FILE PROCESSING SECTION
    # ======================================================
    if uploaded_file is not None:
        # Simpan file bytes ke session state JIKA belum ada atau file berbeda
        if ("uploaded_file_bytes" not in st.session_state or 
            st.session_state.get("uploaded_file_name") != uploaded_file.name):
            
            st.session_state["uploaded_file_bytes"] = uploaded_file.getvalue()
            st.session_state["uploaded_file_name"] = uploaded_file.name
            st.session_state["processed_data"] = None
            st.session_state["slot_strings"] = None
            print(f"✅ File saved to session: {uploaded_file.name}")
        
        # Tampilkan file info
        st.success(f"✅ File terupload: **{uploaded_file.name}**")
        
        # Preview file
        with st.expander("📄 Preview File Upload", expanded=False):
            try:
                file_stream = io.BytesIO(st.session_state["uploaded_file_bytes"])
                excel_data = pd.ExcelFile(file_stream)
                st.write(f"**Sheet yang ditemukan:** {excel_data.sheet_names}")
                
                for sheet in ['Reguler', 'Poleks']:
                    if sheet in excel_data.sheet_names:
                        file_stream.seek(0)
                        df_sheet = pd.read_excel(file_stream, sheet_name=sheet)
                        st.write(f"**Sheet {sheet}:** {len(df_sheet)} baris")
                        st.dataframe(df_sheet.head(3), width='stretch')
            except Exception as e:
                st.warning(f"Tidak bisa preview file: {e}")
        
        # ======================================================
        # PROCESS BUTTON (SELALU TAMPIL JIKA ADA FILE)
        # ======================================================
        if st.button("🚀 Proses Jadwal", type="primary", width='stretch', 
                    key="process_button"):
            
            with st.spinner("Memproses data... Mohon tunggu"):
                try:
                    # Validasi file
                    file_stream = io.BytesIO(st.session_state["uploaded_file_bytes"])
                    is_valid, message = validator.validate_excel_file(file_stream)
                    
                    if not is_valid:
                        st.error(f"❌ File tidak valid: {message}")
                        st.stop()
                    
                    # Proses data
                    file_stream.seek(0)
                    grid_df, slot_strings, errors = scheduler.process_dataframe(file_stream)
                    
                    if grid_df is not None:
                        # Simpan hasil ke session state
                        st.session_state["processed_data"] = grid_df
                        st.session_state["slot_strings"] = slot_strings
                        st.session_state["processing_errors"] = errors
                        
                        st.success(f"✅ Data berhasil diproses! ({len(grid_df)} baris, {len(slot_strings)} slot waktu)")
                        
                        # Tampilkan preview
                        with st.expander("📋 Preview Hasil Proses", expanded=True):
                            st.write(f"**Dimensi data:** {grid_df.shape[0]} baris × {grid_df.shape[1]} kolom")
                            st.dataframe(grid_df.head(), width='stretch')
                        
                        # Tampilkan errors jika ada
                        if errors:
                            st.warning(f"⚠️ **{len(errors)} peringatan:**")
                            for error in errors[:3]:
                                st.write(f"- {error}")
                            if len(errors) > 3:
                                st.write(f"- ... dan {len(errors) - 3} lainnya")
                    
                    else:
                        st.error("❌ Gagal memproses data")
                        if errors:
                            for error in errors:
                                st.write(f"- {error}")
                
                except Exception as e:
                    st.error(f"❌ Error saat memproses: {str(e)}")
                    st.code(traceback.format_exc())
    
    # ======================================================
    # RESULTS SECTION (SELALU TAMPIL JIKA ADA DATA DI SESSION)
    # ======================================================
    if ("processed_data" in st.session_state and 
        st.session_state["processed_data"] is not None):
        
        st.divider()
        st.subheader("📊 Hasil Proses")
        
        grid_df = st.session_state["processed_data"]
        slot_strings = st.session_state["slot_strings"]
        
        st.write(f"✅ **Data tersedia:** {len(grid_df)} baris, {len(slot_strings)} slot waktu")
        
        # ======================================================
        # DOWNLOAD BUTTONS (SELALU TAMPIL JIKA ADA DATA)
        # ======================================================
        st.subheader("💾 Download Hasil")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if st.button("📥 Download Excel Hasil", width='stretch', key="download_excel"):
                st.session_state["download_clicked"] = True
        
        with col2:
            if st.button("📄 Download Template", width='stretch', key="download_template"):
                st.session_state["download_template"] = True
        
        with col3:
            if st.button("🔄 Proses Ulang", width='stretch', key="reprocess"):
                for key in ["processed_data", "slot_strings", "processing_errors"]:
                    if key in st.session_state:
                        del st.session_state[key]
                st.rerun()
        
        # ======================================================
        # ACTUAL DOWNLOAD HANDLING
        # ======================================================
        
        # Handle Excel download
        if st.session_state.get("download_clicked", False):
            try:
                st.session_state["download_clicked"] = False
                
                with st.spinner("Membuat file Excel..."):
                    # Buat stream baru
                    file_stream = io.BytesIO(st.session_state["uploaded_file_bytes"])
                    
                    # Generate Excel - gunakan datetime dari import global
                    output_buffer = writer.write(
                        source_file=file_stream,
                        df_grid=grid_df,
                        slot_str=slot_strings
                    )
                    
                    # Buat nama file dengan timestamp
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    filename = f"jadwal_hasil_{timestamp}.xlsx"
                    
                    # Tampilkan download button
                    st.download_button(
                        label=f"⬇️ Download: {filename}",
                        data=output_buffer,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        width='stretch',
                        key="excel_download_button"
                    )
                    
                    st.success("✅ File Excel siap di-download!")
            
            except Exception as e:
                st.error(f"❌ Gagal membuat file Excel: {str(e)}")
                st.code(traceback.format_exc())
        
        # Handle Template download
        if st.session_state.get("download_template", False):
            try:
                st.session_state["download_template"] = False
                
                template_buffer = writer.generate_template(slot_strings)
                
                st.download_button(
                    label="⬇️ Klik untuk download Template",
                    data=template_buffer,
                    file_name="template_jadwal.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width='stretch',
                    key="template_download_button"
                )
                
                st.success("✅ Template siap di-download!")
            
            except Exception as e:
                st.error(f"❌ Gagal membuat template: {str(e)}")
    
    # ======================================================
    # NO FILE UPLOADED STATE
    # ======================================================
    elif uploaded_file is None:
        st.info("📤 Silakan upload file Excel dengan sheet Reguler dan Poleks")
        
        with st.expander("ℹ️ Panduan Format File"):
            st.markdown("""
            **Format file Excel harus memiliki:**
            
            1. **Sheet 'Reguler'** - Berisi jadwal reguler
            2. **Sheet 'Poleks'** - Berisi jadwal poleks
            
            **Kolom yang harus ada di setiap sheet:**
            - `Nama Dokter` - Nama lengkap dokter
            - `Poli Asal` - Nama poli
            - `Jenis Poli` - "Reguler" atau "Poleks"
            - `Senin` - Format: "07.30-10.00" atau "07:30-10:00"
            - `Selasa` - Format sama
            - `Rabu` - Format sama
            - `Kamis` - Format sama
            - `Jum'at` - Format sama
            """)
