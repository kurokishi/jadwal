# app/ui/tab_upload.py
import streamlit as st
import pandas as pd
import io

def render_upload_tab(scheduler, writer, analyzer, config):
    st.subheader("📤 Upload & Proses Jadwal")
    
    # ======================================================
    # UPLOAD FILE
    # ======================================================
    uploaded_file = st.file_uploader(
        "Upload file Excel (Format: sheet Reguler & Poleks)",
        type=['xlsx', 'xls']
    )
    
    if uploaded_file is not None:
        try:
            # Simpan file ke session state
            st.session_state["uploaded_file"] = uploaded_file
            
            # Baca file untuk preview
            excel_data = pd.ExcelFile(uploaded_file)
            
            # Tampilkan sheet yang tersedia
            st.write("📄 **Sheet yang ditemukan:**", excel_data.sheet_names)
            
            # Preview sheet Reguler
            if 'Reguler' in excel_data.sheet_names:
                with st.expander("👁️ Preview Sheet Reguler"):
                    df_reguler = pd.read_excel(excel_data, sheet_name='Reguler')
                    st.dataframe(df_reguler.head(), use_container_width=True)
                    st.caption(f"Total baris: {len(df_reguler)}")
            
            # Preview sheet Poleks
            if 'Poleks' in excel_data.sheet_names:
                with st.expander("👁️ Preview Sheet Poleks"):
                    df_poleks = pd.read_excel(excel_data, sheet_name='Poleks')
                    st.dataframe(df_poleks.head(), use_container_width=True)
                    st.caption(f"Total baris: {len(df_poleks)}")
            
            # ======================================================
            # PROSES DATA
            # ======================================================
            if st.button("🚀 Proses Jadwal", type="primary", use_container_width=True):
                with st.spinner("Memproses data..."):
                    try:
                        # Simpan file temporary
                        file_bytes = uploaded_file.getvalue()
                        
                        # Proses dengan scheduler
                        grid_df, slot_strings, errors = scheduler.process_dataframe(file_bytes)
                        
                        if grid_df is not None:
                            # Simpan ke session state
                            st.session_state["processed_data"] = grid_df
                            st.session_state["slot_strings"] = slot_strings
                            
                            # Tampilkan preview hasil
                            st.success("✅ Data berhasil diproses!")
                            
                            with st.expander("📋 Preview Hasil Proses"):
                                st.dataframe(grid_df.head(), use_container_width=True)
                                st.caption(f"Total baris: {len(grid_df)}")
                                st.caption(f"Slot waktu: {', '.join(slot_strings[:5])}...")
                            
                            # Tampilkan error/warning jika ada
                            if errors:
                                st.warning("⚠️ **Peringatan selama pemrosesan:**")
                                for error in errors:
                                    st.write(f"- {error}")
                            
                            # ======================================================
                            # DOWNLOAD HASIL
                            # ======================================================
                            st.subheader("💾 Download Hasil")
                            
                            # Tombol download Excel
                            if st.button("📥 Download Excel Hasil"):
                                with st.spinner("Membuat file Excel..."):
                                    try:
                                        # Buat file template/output
                                        output_buffer = writer.write(
                                            source_file=file_bytes,
                                            df_grid=grid_df,
                                            slot_str=slot_strings
                                        )
                                        
                                        st.download_button(
                                            label="⬇️ Klik untuk download file Excel",
                                            data=output_buffer,
                                            file_name="jadwal_hasil.xlsx",
                                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                        )
                                    except Exception as e:
                                        st.error(f"Gagal membuat file Excel: {str(e)}")
                            
                            # Tombol download template
                            if st.button("📄 Download Template Kosong"):
                                template_buffer = writer.generate_template(slot_strings)
                                st.download_button(
                                    label="⬇️ Download Template",
                                    data=template_buffer,
                                    file_name="template_jadwal.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                        
                        else:
                            st.error("❌ Gagal memproses data")
                            if errors:
                                for error in errors:
                                    st.write(f"- {error}")
                                    
                    except Exception as e:
                        st.error(f"❌ Error saat memproses: {str(e)}")
                        
        except Exception as e:
            st.error(f"❌ Error membaca file: {str(e)}")
    
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
            - `Senin` - Format waktu: "07.30-10.00" atau "07:30-10:00"
            - `Selasa` - Format waktu sama
            - `Rabu` - Format waktu sama
            - `Kamis` - Format waktu sama
            - `Jum'at` - Format waktu sama
            - `Sabtu` (opsional) - Format waktu sama
            
            **Contoh format waktu:**
            - 07.30-10.00
            - 07:30-10:00
            - 08.00 - 12.00
            """)
    
    # ======================================================
    # DATA YANG SUDAH DIPROSES
    # ======================================================
    if "processed_data" in st.session_state and st.session_state["processed_data"] is not None:
        st.divider()
        st.subheader("📊 Data Hasil Proses")
        
        df_processed = st.session_state["processed_data"]
        st.write(f"📈 **Statistik:** {len(df_processed)} baris data")
        
        if st.button("🔄 Reset Data"):
            for key in ["uploaded_file", "processed_data", "slot_strings"]:
                if key in st.session_state:
                    del st.session_state[key]
            st.rerun()
