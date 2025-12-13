# app/ui/tab_upload.py
import streamlit as st
import pandas as pd
import io
import traceback

def render_upload_tab(scheduler, writer, analyzer, validator, config):
    st.subheader("📤 Upload & Proses Jadwal")
    
    # ======================================================
    # UPLOAD FILE
    # ======================================================
    uploaded_file = st.file_uploader(
        "Upload file Excel (Format: sheet Reguler & Poleks)",
        type=['xlsx', 'xls'],
        help="File harus memiliki sheet 'Reguler' dan 'Poleks'"
    )
    
    if uploaded_file is not None:
        try:
            # Simpan file bytes untuk digunakan kembali
            file_bytes = uploaded_file.getvalue()
            
            # Validasi file
            st.info("🔍 Memvalidasi file...")
            
            # Buat BytesIO dari file bytes
            file_stream = io.BytesIO(file_bytes)
            file_stream.seek(0)  # Pastikan di awal
            
            is_valid, message = validator.validate_excel_file(file_stream)
            
            if not is_valid:
                st.error(f"❌ File tidak valid: {message}")
                return
            
            st.success("✅ File valid!")
            
            # Preview file
            with st.expander("📄 Preview File Upload"):
                try:
                    # Reset stream untuk preview
                    file_stream.seek(0)
                    excel_data = pd.ExcelFile(file_stream)
                    st.write(f"**Sheet yang ditemukan:** {excel_data.sheet_names}")
                    
                    for sheet in ['Reguler', 'Poleks']:
                        if sheet in excel_data.sheet_names:
                            # Reset stream lagi
                            file_stream.seek(0)
                            df_sheet = pd.read_excel(file_stream, sheet_name=sheet)
                            st.write(f"**Sheet {sheet}:** {len(df_sheet)} baris")
                            st.dataframe(df_sheet.head(3), width='stretch')
                except Exception as e:
                    st.warning(f"Tidak bisa preview file: {e}")
            
            # ======================================================
            # PROSES DATA
            # ======================================================
            if st.button("🚀 Proses Jadwal", type="primary", width='stretch'):
                with st.spinner("Memproses data... Mohon tunggu"):
                    try:
                        # Debug info
                        st.write("🔄 **Memulai proses...**")
                        
                        # Reset stream untuk proses
                        file_stream.seek(0)
                        
                        # Proses dengan scheduler
                        st.write("1. Membersihkan data...")
                        grid_df, slot_strings, errors = scheduler.process_dataframe(file_stream)
                        
                        # Debug output
                        st.write(f"2. Hasil: grid_df={'Ada' if grid_df is not None else 'Tidak'}, "
                               f"slots={len(slot_strings) if slot_strings else 0}, "
                               f"errors={len(errors) if errors else 0}")
                        
                        if grid_df is not None:
                            # Simpan ke session state
                            st.session_state["processed_data"] = grid_df
                            st.session_state["slot_strings"] = slot_strings
                            st.session_state["uploaded_file_bytes"] = file_bytes  # Simpan bytes, bukan stream
                            
                            st.success(f"✅ Data berhasil diproses! ({len(grid_df)} baris, {len(slot_strings)} slot waktu)")
                            
                            # Preview hasil
                            with st.expander("📋 Preview Hasil Proses"):
                                st.write(f"**Dimensi data:** {grid_df.shape[0]} baris × {grid_df.shape[1]} kolom")
                                st.dataframe(grid_df.head(), width='stretch')
                            
                            # Tampilkan errors/warnings
                            if errors:
                                st.warning(f"⚠️ **{len(errors)} peringatan selama pemrosesan:**")
                                for error in errors[:5]:
                                    st.write(f"- {error}")
                                if len(errors) > 5:
                                    st.write(f"- ... dan {len(errors) - 5} peringatan lainnya")
                            
                            # ======================================================
                            # DOWNLOAD HASIL
                            # ======================================================
                            st.subheader("💾 Download Hasil")
                            
                            # Tampilkan tombol download hanya jika proses berhasil
                            if grid_df is not None and slot_strings:
                                col1, col2 = st.columns(2)
                                
                                with col1:
                                    if st.button("📥 Download Excel Hasil", width='stretch'):
                                        with st.spinner("Membuat file Excel..."):
                                            try:
                                                # Buat stream baru dari bytes yang disimpan
                                                new_stream = io.BytesIO(st.session_state["uploaded_file_bytes"])
                                                
                                                output_buffer = writer.write(
                                                    source_file=new_stream,
                                                    df_grid=grid_df,
                                                    slot_str=slot_strings
                                                )
                                                
                                                # Tampilkan download button
                                                st.download_button(
                                                    label="⬇️ Klik untuk download file Excel",
                                                    data=output_buffer,
                                                    file_name="jadwal_hasil.xlsx",
                                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                    width='stretch'
                                                )
                                            except Exception as e:
                                                st.error(f"❌ Gagal membuat file Excel: {str(e)}")
                                                st.code(traceback.format_exc())
                                
                                with col2:
                                    if st.button("📄 Download Template", width='stretch'):
                                        try:
                                            template_buffer = writer.generate_template(slot_strings)
                                            st.download_button(
                                                label="⬇️ Download Template",
                                                data=template_buffer,
                                                file_name="template_jadwal.xlsx",
                                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                                width='stretch'
                                            )
                                        except Exception as e:
                                            st.error(f"❌ Gagal membuat template: {str(e)}")
                        
                        else:
                            st.error("❌ Gagal memproses data")
                            if errors:
                                st.error("**Detail error:**")
                                for error in errors:
                                    st.write(f"- {error}")
                            
                    except Exception as e:
                        st.error(f"❌ Error saat memproses: {str(e)}")
                        st.code(traceback.format_exc())
                        
        except Exception as e:
            st.error(f"❌ Error membaca file: {str(e)}")
            st.code(traceback.format_exc())
    
    else:
        st.info("📤 Silakan upload file Excel dengan sheet Reguler dan Poleks")
        
        # ======================================================
        # TEMPLATE & CONTOH
        # ======================================================
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
            
            **Contoh data:**
            | Nama Dokter | Poli Asal | Jenis Poli | Senin | Selasa |
            |-------------|-----------|------------|-------|--------|
            | dr. Contoh | Poli Anak | Reguler | 08.00-10.00 | 09.00-11.00 |
            """)
    
    # ======================================================
    # DATA YANG SUDAH DIPROSES
    # ======================================================
    if "processed_data" in st.session_state and st.session_state["processed_data"] is not None:
        st.divider()
        st.subheader("📊 Data Hasil Proses")
        
        df_processed = st.session_state["processed_data"]
        st.write(f"📈 **Statistik:** {len(df_processed)} baris data tersimpan")
        
        # Tombol download langsung jika data sudah ada
        if st.button("📥 Download Hasil Sekarang", width='stretch'):
            if "uploaded_file_bytes" in st.session_state and "slot_strings" in st.session_state:
                with st.spinner("Membuat file Excel..."):
                    try:
                        grid_df = st.session_state["processed_data"]
                        slot_strings = st.session_state["slot_strings"]
                        
                        # Buat stream baru
                        new_stream = io.BytesIO(st.session_state["uploaded_file_bytes"])
                        
                        output_buffer = writer.write(
                            source_file=new_stream,
                            df_grid=grid_df,
                            slot_str=slot_strings
                        )
                        
                        st.download_button(
                            label="⬇️ Klik untuk download",
                            data=output_buffer,
                            file_name="jadwal_hasil.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            width='stretch'
                        )
                    except Exception as e:
                        st.error(f"❌ Gagal membuat file: {str(e)}")
        
        if st.button("🔄 Reset Data", width='stretch'):
            for key in ["uploaded_file_bytes", "processed_data", "slot_strings"]:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
