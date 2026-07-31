import streamlit as st
import datetime
import calendar
from etl_kpi import run_etl_kpi
from etl_raw_sts import run_etl_raw
from etl_ban import run_etl_ban
from etl_sts import run_etl_sts
from etl_sparepart import run_etl_sparepart
import pandas as pd
import io

# Gunakan cache untuk pembacaan excel agar tidak berat saat ganti sheet
@st.cache_data
def get_dfs(bytes_obj):
    return pd.read_excel(io.BytesIO(bytes_obj), dtype=str, sheet_name=None)

# Konfigurasi File Path Default
DEFAULT_INPUT_FILE = 'data/MEKANIK - STSFORM.csv'
DEFAULT_MASTER_FILE = 'data/KPI 95% - Available Car Report - 2026.xlsx'

st.set_page_config(page_title="KPI & Fleet Dashboard", layout="wide")
st.title("Aplikasi Laporan Fleet & Analisis Ban")

tab_kpi, tab_ban, tab_sts, tab_sparepart = st.tabs(["Dashboard KPI & STS", "Analisis Pencapaian KM Ban", "Capaian KM STS", "ETL Sparepart"])

# =======================
# TAB 1: KPI & STS
# =======================
with tab_kpi:
    st.markdown("Dashboard interaktif untuk mengatur rentang tanggal dan men-generate laporan KPI secara otomatis berdasarkan data Servis Mekanik dan Master Data.")
    st.info("Data servis akan otomatis dibaca mundur 1 tahun dari tanggal awal yang Anda pilih agar histori perbaikan mobil lama (belum keluar) tidak terpotong.")

    # UI Layout: 3 Columns
    col_date, col_sts, col_master = st.columns(3)

    with col_date:
        st.subheader("Rentang Tanggal")
        # Default ke bulan ini
        today = datetime.date.today()
        default_start = today.replace(day=1)
        # Bulan terakhir
        last_day = calendar.monthrange(today.year, today.month)[1]
        default_end = today.replace(day=last_day)
        
        start_date = st.date_input("Tanggal Mulai", default_start)
        end_date = st.date_input("Tanggal Akhir", default_end)

    with col_sts:
        st.subheader("Data Servis (STS)")
        custom_file = st.file_uploader("Upload MEKANIK - STSFORM.csv terbaru (Opsional)", type=['csv'], key='sts_upload')
        st.caption("Jika dikosongkan, sistem akan menggunakan data default di server.")

    with col_master:
        st.subheader("Data Master Identitas")
        custom_master = st.file_uploader("Upload Data Master Identitas (Opsional)", type=['csv', 'xlsx'], key='master_upload')
        st.caption("Gunakan file Master Aset Kendaraan Baru untuk sinkronisasi identitas.")

    st.markdown("---")

    # Inisialisasi Session State KPI
    if 'bytes_kpi' not in st.session_state:
        st.session_state.bytes_kpi = None
    if 'bytes_raw' not in st.session_state:
        st.session_state.bytes_raw = None
    if 'report_dates' not in st.session_state:
        st.session_state.report_dates = ("", "")

    if st.button("Generate Reports", width='stretch', type="primary", key="btn_kpi"):
        if start_date > end_date:
            st.error("Tanggal Awal tidak boleh lebih besar dari Tanggal Akhir!")
        else:
            with st.spinner("Memproses data KPI & STS..."):
                try:
                    # Path file data masukan (Pakai file lokal default jika tidak upload)
                    input_path = DEFAULT_INPUT_FILE
                    master_path = DEFAULT_MASTER_FILE
                    
                    if custom_file is not None:
                        with open("data/TEMP_STSFORM.csv", "wb") as f:
                            f.write(custom_file.getbuffer())
                        input_path = "data/TEMP_STSFORM.csv"
                        
                    if custom_master is not None:
                        ext = custom_master.name.split('.')[-1]
                        with open(f"data/TEMP_MASTER.{ext}", "wb") as f:
                            f.write(custom_master.getbuffer())
                        master_path = f"data/TEMP_MASTER.{ext}"
                    
                    # Format string untuk skrip
                    ts = start_date.strftime("%Y-%m-%d")
                    te = end_date.strftime("%Y-%m-%d")
                    
                    # Simpan ke session state
                    st.session_state.bytes_kpi = run_etl_kpi(ts, te, input_path, master_path)
                    st.session_state.bytes_raw = run_etl_raw(ts, te, input_path)
                    st.session_state.report_dates = (ts, te)
                    
                    st.success("Laporan berhasil dibuat.")

                except Exception as e:
                    import traceback
                    st.error(f"Terjadi kesalahan: {e}")
                    st.session_state.bytes_kpi = None
                    st.session_state.bytes_raw = None
                    st.code(traceback.format_exc())

    # Bagian Display & Preview (Hanya muncul jika data ada di session state)
    if st.session_state.bytes_kpi is not None:
        ts, te = st.session_state.report_dates
        
        st.markdown("---")
        dl_col1, dl_col2 = st.columns(2)
        
        filename_kpi = f"Output_KPI_Report_{ts}_to_{te}_FINAL.xlsx"
        dl_col1.download_button(
            label="Download Laporan Utama",
            data=st.session_state.bytes_kpi,
            file_name=filename_kpi,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch'
        )
        
        filename_raw = f"Output_RAW_STS_Report_{ts}_to_{te}.xlsx"
        dl_col2.download_button(
            label="Download Laporan Audit Mekanik",
            data=st.session_state.bytes_raw,
            file_name=filename_raw,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch'
        )
        
        # --- PREVIEW UI ---
        st.markdown("---")
        st.subheader("Preview Laporan KPI")
        st.info("Tips: Pilih sheet di bawah ini untuk melihat datanya. Anda dapat menyalin data langsung dari tabel.")
        
        dfs_preview = get_dfs(st.session_state.bytes_kpi)
        
        sheet_names = list(dfs_preview.keys())
        selected_sheet = st.selectbox("Tampilkan Sheet:", sheet_names)
        
        def color_status(val):
            if pd.isna(val): return ''
            val_str = str(val).strip().upper()
            if val_str == 'A':
                return 'background-color: #c3e6cb; color: #155724;'
            elif val_str in ['B-INT', 'B - INT', 'AB-INT', 'AB - INT']:
                return 'background-color: #FFA500; color: black;'
            elif val_str in ['B-EXT', 'B - EXT', 'AB-EXT', 'AB - EXT']:
                return 'background-color: #FF4500; color: white;'
            elif val_str in ['B-INS', 'B - INS']:
                return 'background-color: #8A2BE2; color: white;'
            elif val_str == 'R':
                return 'background-color: #d6d8db; color: #383d41;'
            return ''
        
        df_to_show = dfs_preview[selected_sheet]
        
        if selected_sheet == 'SUMMARY_USIA':
            st.dataframe(df_to_show, width='stretch', hide_index=True)
        else:
            try:
                styled_df = df_to_show.style.map(color_status)
            except AttributeError:
                styled_df = df_to_show.style.applymap(color_status)
            
            st.dataframe(styled_df, width='stretch', hide_index=True)


# =======================
# TAB 2: ANALISIS BAN
# =======================
with tab_ban:
    st.markdown("Dashboard untuk menganalisis pencapaian KM Ban menggunakan data **BAN SHARE - Entry Data**. Upload file dengan baris ke-7 sebagai header utama.")
    
    ban_col1, ban_col2 = st.columns(2)
    with ban_col1:
        ban_start_date = st.date_input("Tanggal Awal (Pelepasan)", datetime.date.today().replace(day=1), key="ban_start")
    with ban_col2:
        ban_end_date = st.date_input("Tanggal Akhir (Pelepasan)", datetime.date.today(), key="ban_end")

    ban_file = st.file_uploader("Upload File BAN SHARE (Excel / CSV)", type=['xlsx', 'xls', 'csv'], key='upload_ban')
    
    if 'bytes_ban' not in st.session_state:
        st.session_state.bytes_ban = None
        st.session_state.df_ban = None
        
    if st.button("Generate Analisis Ban", type="primary", key="btn_ban_run", width="stretch"):
        if ban_file is None:
            st.error("Silakan upload file BAN SHARE terlebih dahulu.")
        else:
            with st.spinner("Memproses Analisis Ban..."):
                try:
                    bytes_b, df_b = run_etl_ban(ban_file.getvalue(), filename=ban_file.name, start_date=ban_start_date, end_date=ban_end_date)
                    st.session_state.bytes_ban = bytes_b
                    st.session_state.df_ban = df_b
                    st.success("Laporan Analisis Ban berhasil dibuat.")
                except Exception as e:
                    import traceback
                    st.error(f"Terjadi kesalahan saat memproses Ban Share: {e}")
                    st.code(traceback.format_exc())
                    
    if st.session_state.bytes_ban is not None:
        st.markdown("---")
        st.download_button(
            label="Download Laporan Analisis Ban (Excel)",
            data=st.session_state.bytes_ban,
            file_name="Laporan_Analisis_KM_Ban.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
            key='dl_ban'
        )
        
        st.subheader("Preview Jarak Tempuh per Ban")
        st.dataframe(st.session_state.df_ban, width='stretch', hide_index=True)


# =======================
# TAB 3: CAPAIAN KM STS
# =======================
with tab_sts:
    st.markdown("Dashboard untuk menganalisis total pencapaian KM per Nopol berdasarkan file **STSFORM** pada rentang waktu tertentu.")
    
    sts_col1, sts_col2 = st.columns(2)
    with sts_col1:
        sts_start_date = st.date_input("Tanggal Awal", datetime.date.today().replace(day=1), key="sts_start")
    with sts_col2:
        sts_end_date = st.date_input("Tanggal Akhir", datetime.date.today(), key="sts_end")
        
    sts_file = st.file_uploader("Upload File STSFORM (Excel / CSV)", type=['xlsx', 'xls', 'csv'], key='upload_sts_km')
    
    if 'bytes_sts_km' not in st.session_state:
        st.session_state.bytes_sts_km = None
        st.session_state.df_sts_km = None
        
    if st.button("Generate Analisis Capaian KM STS", type="primary", key="btn_sts_run", width="stretch"):
        if sts_file is None:
            st.error("Silakan upload file STSFORM terlebih dahulu.")
        else:
            with st.spinner("Memproses Data STS..."):
                try:
                    bytes_s, df_s = run_etl_sts(sts_file.getvalue(), filename=sts_file.name, start_date=sts_start_date, end_date=sts_end_date)
                    st.session_state.bytes_sts_km = bytes_s
                    st.session_state.df_sts_km = df_s
                    st.success("Laporan Capaian KM STS berhasil dibuat.")
                except Exception as e:
                    import traceback
                    st.error(f"Terjadi kesalahan saat memproses STS: {e}")
                    st.code(traceback.format_exc())
                    
    if st.session_state.bytes_sts_km is not None:
        st.markdown("---")
        st.download_button(
            label="Download Laporan Capaian KM STS (Excel)",
            data=st.session_state.bytes_sts_km,
            file_name="Laporan_Capaian_KM_STS.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
            key='dl_sts_km'
        )
        
        st.subheader("Preview Jarak Tempuh per Nopol")
        st.dataframe(st.session_state.df_sts_km, width='stretch', hide_index=True)


# =======================
# TAB 4: ETL SPAREPART
# =======================
with tab_sparepart:
    st.markdown("Dashboard untuk menganalisis jumlah pemakaian sparepart dan total harga.")
    
    sp_col1, sp_col2 = st.columns(2)
    with sp_col1:
        sp_start_date = st.date_input("Tanggal Awal (RF)", datetime.date.today().replace(day=1), key="sp_start")
    with sp_col2:
        sp_end_date = st.date_input("Tanggal Akhir (RF)", datetime.date.today(), key="sp_end")
        
    sp_file = st.file_uploader("Upload File Sparepart (Excel / CSV)", type=['xlsx', 'xls', 'csv'], key='upload_sp')
    
    if 'bytes_sp' not in st.session_state:
        st.session_state.bytes_sp = None
        st.session_state.df_sp = None
        
    if st.button("Generate Analisis Sparepart", type="primary", key="btn_sp_run", width="stretch"):
        if sp_file is None:
            st.error("Silakan upload file Sparepart terlebih dahulu.")
        else:
            with st.spinner("Memproses Data Sparepart..."):
                try:
                    bytes_s, df_s = run_etl_sparepart(sp_file.getvalue(), filename=sp_file.name, start_date=sp_start_date, end_date=sp_end_date)
                    st.session_state.bytes_sp = bytes_s
                    st.session_state.df_sp = df_s
                    st.success("Laporan Sparepart berhasil dibuat.")
                except Exception as e:
                    import traceback
                    st.error(f"Terjadi kesalahan saat memproses Sparepart: {e}")
                    st.code(traceback.format_exc())
                    
    if st.session_state.bytes_sp is not None:
        st.markdown("---")
        st.download_button(
            label="Download Laporan Sparepart (Excel)",
            data=st.session_state.bytes_sp,
            file_name="Laporan_ETL_Sparepart.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
            key='dl_sp'
        )
        
        st.subheader("Preview Data Sparepart")
        st.dataframe(st.session_state.df_sp, width='stretch', hide_index=True)
