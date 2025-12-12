# app/ui/tab_upload.py
import streamlit as st
import pandas as pd
from app.core.validator import Validator

def render_upload_tab(scheduler, writer, analyzer, config):
    st.subheader("📤 Upload Jadwal")
    st.info("Silakan upload file Excel berformat Reguler & Poleks.")

    # ====== siapkan slot_str dari scheduler (list "HH:MM") ======
    try:
        slots_dt = scheduler.generate_slots()  # list of datetime.time
        slot_str = [t.strftime("%H:%M") for t in slots_dt]
    except Exception:
        # fallback: gunakan beberapa slot default jika generator gagal
        slot_str = [f"{h:02d}:{m:02d}" for h in range(7, 15) for m in (0,30)]
        slot_str = sorted(list(dict.fromkeys(slot_str)))  # dedup & keep order

    # ================= TEMPLATE DOWNLOAD =================
    st.subheader("📄 Download Template Excel")
    if st.button("📥 Download Template Jadwal"):
        template_buf = writer.generate_template(slot_str)
        st.download_button(
            label="Klik untuk Download Template",
            data=template_buf,
            file_name="template_jadwal.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    st.write("---")

    # ================== FILE UPLOADER =====================
    uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

    if not uploaded:
        return

    ok, err = Validator.validate(uploaded)
    if not ok:
        st.error(f"❌ File tidak valid: {err}")
        return

    try:
        xl = pd.ExcelFile(uploaded)
        st.success(f"File valid. Sheets: {xl.sheet_names}")
    except Exception as e:
        st.error(f"Gagal membaca file Excel: {e}")
        return

    if "Reguler" in xl.sheet_names and st.checkbox("Preview sheet Reguler"):
        try:
            st.dataframe(pd.read_excel(uploaded, sheet_name="Reguler", nrows=10))
        except Exception as e:
            st.warning(f"Gagal preview sheet Reguler: {e}")

    # ================== PROSES ============================
    if st.button("🚀 Proses Jadwal"):

        try:
            df_reg = xl.parse("Reguler") if "Reguler" in xl.sheet_names else pd.DataFrame()
        except Exception:
            df_reg = pd.DataFrame()

        try:
            df_pol = xl.parse("Poleks") if "Poleks" in xl.sheet_names else pd.DataFrame()
        except Exception:
            df_pol = pd.DataFrame()

        if df_reg.empty and df_pol.empty:
            st.warning("File tidak berisi data di sheet 'Reguler' atau 'Poleks'.")
            return

        with st.spinner("Memproses jadwal..."):
            try:
                df_r = scheduler.process_schedule(df_reg, "Reguler") if not df_reg.empty else pd.DataFrame()
                df_e = scheduler.process_schedule(df_pol, "Poleks") if not df_pol.empty else pd.DataFrame()
                df_all = pd.concat([df_r, df_e], ignore_index=True) if (not df_r.empty or not df_e.empty) else pd.DataFrame()
            except Exception as e:
                st.error(f"Gagal memproses jadwal: {e}")
                return

        if df_all.empty:
            st.warning("Hasil proses kosong. Periksa kembali input Anda.")
            return

        # simpan ke session menggunakan slot_str yang valid
        st.session_state["processed_data"] = df_all
        st.session_state["time_slots"] = slot_str

        st.success("✅ Jadwal berhasil diproses!")
        st.dataframe(df_all, use_container_width=True)

        # SAVE -> gunakan slot_str
        try:
            buf = writer.write(uploaded, df_all, slot_str)
            st.download_button(
                "📥 Download Jadwal Hasil",
                data=buf,
                file_name="jadwal_hasil.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Gagal membuat file Excel hasil: {e}")
