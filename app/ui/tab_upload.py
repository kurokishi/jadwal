# app/ui/tab_upload.py
import streamlit as st
import pandas as pd
from app.core.validator import Validator


def render_upload_tab(scheduler, writer, analyzer, config):

    st.subheader("📤 Upload Jadwal")
    st.info("Upload file Excel yang berisi sheet **Reguler** dan **Poleks**.")

    # ============================================================
    # SLOT GENERATOR (aman)
    # ============================================================
    try:
        slot_times = scheduler.generate_slots()       # list of datetime.time
        slot_str = [t.strftime("%H:%M") for t in slot_times]
    except Exception:
        # fallback aman
        slot_str = [f"{h:02d}:{m:02d}" for h in range(7, 15) for m in (0, 30)]
        slot_str = sorted(list(dict.fromkeys(slot_str)))

    # ============================================================
    # TEMPLATE DOWNLOAD
    # ============================================================
    st.subheader("📄 Download Template Jadwal")

    if st.button("📥 Download Template Jadwal"):
        try:
            template_buf = writer.generate_template(slot_str)
            st.download_button(
                label="Klik untuk download template",
                data=template_buf,
                file_name="template_jadwal.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Gagal membuat template: {e}")

    st.markdown("---")

    # ============================================================
    # FILE UPLOADER
    # ============================================================
    uploaded_file = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

    if uploaded_file is None:
        return

    # ============================================================
    # VALIDASI FILE
    # ============================================================
    ok, err = Validator.validate(uploaded_file)
    if not ok:
        st.error(f"❌ File tidak valid: {err}")
        return

    # ============================================================
    # BACA EXCEL
    # ============================================================
    try:
        xl = pd.ExcelFile(uploaded_file)
        st.success(f"File valid. Sheets tersedia: {xl.sheet_names}")
    except Exception as e:
        st.error(f"Gagal membaca workbook: {e}")
        return

    # PREVIEW OPTIONAL
    if "Reguler" in xl.sheet_names and st.checkbox("Preview sheet Reguler"):
        try:
            prev = pd.read_excel(uploaded_file, sheet_name="Reguler", nrows=12)
            st.dataframe(prev, use_container_width=True)
        except Exception as e:
            st.warning(f"Gagal preview sheet Reguler: {e}")

    st.write("---")

    # ============================================================
    # PROSES JADWAL
    # ============================================================
    if st.button("🚀 Proses Jadwal"):

        # Load sheet Reguler
        try:
            df_reg = xl.parse("Reguler") if "Reguler" in xl.sheet_names else pd.DataFrame()
        except Exception:
            df_reg = pd.DataFrame()

        # Load sheet Poleks
        try:
            df_pol = xl.parse("Poleks") if "Poleks" in xl.sheet_names else pd.DataFrame()
        except Exception:
            df_pol = pd.DataFrame()

        if df_reg.empty and df_pol.empty:
            st.warning("Tidak ditemukan data pada sheet **Reguler** maupun **Poleks**.")
            return

        # ============================================================
        # Jalankan scheduler (safe)
        # ============================================================
        with st.spinner("⏳ Memproses jadwal..."):
            try:
                df_R = scheduler.process_schedule(df_reg, "Reguler") if not df_reg.empty else pd.DataFrame()
                df_E = scheduler.process_schedule(df_pol, "Poleks") if not df_pol.empty else pd.DataFrame()

                df_all = pd.concat([df_R, df_E], ignore_index=True)
            except Exception as e:
                st.error(f"Gagal memproses jadwal: {e}")
                return

        if df_all.empty:
            st.warning("Hasil proses kosong. Periksa format input Anda.")
            return

        # simpan ke session state
        st.session_state["processed_data"] = df_all
        st.session_state["time_slots"] = slot_str

        st.success("✅ Jadwal berhasil diproses!")

        st.dataframe(df_all, use_container_width=True)

        # ============================================================
        # EXPORT EXCEL
        # ============================================================
        try:
            buf = writer.write(uploaded_file, df_all, slot_str)
            st.download_button(
                "📥 Download Hasil Jadwal",
                data=buf,
                file_name="jadwal_hasil.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"Gagal membuat output Excel: {e}")
