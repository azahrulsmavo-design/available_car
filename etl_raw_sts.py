import pandas as pd
import numpy as np
from datetime import datetime
import io

def run_etl_raw(target_start_date_str, target_end_date_str, input_file):
    start_date = pd.to_datetime(target_start_date_str)
    end_date = pd.to_datetime(target_end_date_str)
    
    # 1 YEAR LOOKBACK
    fetch_start_date = start_date - pd.DateOffset(years=1)

    # ==========================================
    # 2. EKSTRAKSI & PEMBERSIHAN DATA RAW (CSV / EXCEL)
    # ==========================================
    is_excel = False
    if isinstance(input_file, str) and (input_file.endswith('.xlsx') or input_file.endswith('.xls')):
        is_excel = True
        
    df_raw = None
    for skip in [0, 1, 2, 3, 4, 5, 17, 18]:
        try:
            if is_excel:
                temp_df = pd.read_excel(input_file, skiprows=skip, nrows=5)
            else:
                temp_df = pd.read_csv(input_file, skiprows=skip, nrows=5, low_memory=False)
            cols_check = [str(c).replace('\n', ' ').strip().upper() for c in temp_df.columns]
            if any('NOPOL' in c for c in cols_check) or (any('STATUS' in c for c in cols_check) and any('MASUK' in c for c in cols_check)):
                if is_excel:
                    df_raw = pd.read_excel(input_file, skiprows=skip)
                else:
                    df_raw = pd.read_csv(input_file, skiprows=skip, low_memory=False)
                break
        except Exception:
            continue

    if df_raw is None:
        if is_excel:
            df_raw = pd.read_excel(input_file)
        else:
            df_raw = pd.read_csv(input_file, low_memory=False)

    df_raw.columns = [str(c).replace('\n', ' ').strip() for c in df_raw.columns]

    # Map target columns flexibly
    col_mapping = {}
    for c in df_raw.columns:
        c_clean = str(c).upper().replace('\n', ' ').replace('_', ' ').strip()
        if 'BU' not in col_mapping.values() and ('BU MASTER' in c_clean or c_clean == 'BU'):
            col_mapping[c] = 'BU'
        elif 'DEPT' not in col_mapping.values() and ('DEPT' in c_clean or 'DEPARTMENT' in c_clean or 'SECTION' in c_clean):
            col_mapping[c] = 'DEPT'
        elif 'LOKASI' not in col_mapping.values() and ('CABANG MASTER' in c_clean or 'LOKASI' in c_clean or 'LOCATION' in c_clean):
            col_mapping[c] = 'LOKASI'
        elif 'JENIS_MOBIL' not in col_mapping.values() and ('JENIS KENDARAAN' in c_clean or 'JENIS MOBIL' in c_clean):
            col_mapping[c] = 'JENIS_MOBIL'
        elif 'MERK' not in col_mapping.values() and ('MEREK KENDARAAN' in c_clean or 'MERK' in c_clean or 'MEREK' in c_clean):
            col_mapping[c] = 'MERK'
        elif 'NOPOL' not in col_mapping.values() and ('NOPOL' in c_clean or 'NO POL' in c_clean or 'POLISI' in c_clean):
            col_mapping[c] = 'NOPOL'
        elif 'USIA' not in col_mapping.values() and ('USIA KENDARAAN' in c_clean or 'USIA' in c_clean):
            col_mapping[c] = 'USIA'
        elif 'STATUS_BENGKEL' not in col_mapping.values() and ('STATUS BENGKEL' in c_clean or 'STATUS_BENGKEL' in c_clean or c_clean == 'STATUS'):
            col_mapping[c] = 'STATUS_BENGKEL'
        elif 'TGL_MASUK' not in col_mapping.values() and ('TGL MASUK' in c_clean or 'TANGGAL MASUK' in c_clean or 'MASUK BENGKEL' in c_clean):
            col_mapping[c] = 'TGL_MASUK'
        elif 'TGL_KELUAR' not in col_mapping.values() and ('TGLKELUAR' in c_clean or 'TGL KELUAR' in c_clean or 'TANGGAL KELUAR' in c_clean or 'KELUAR BENGKEL' in c_clean):
            col_mapping[c] = 'TGL_KELUAR'

    df = df_raw[list(col_mapping.keys())].copy()
    df.rename(columns=col_mapping, inplace=True)

    for req in ['BU', 'DEPT', 'LOKASI', 'JENIS_MOBIL', 'MERK', 'NOPOL', 'USIA', 'STATUS_BENGKEL', 'TGL_MASUK', 'TGL_KELUAR']:
        if req not in df.columns:
            df[req] = np.nan

    # Clean Text
    text_cols = ['BU', 'DEPT', 'LOKASI', 'JENIS_MOBIL', 'MERK', 'NOPOL', 'STATUS_BENGKEL']
    for col in text_cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().str.upper()

    df['TGL_MASUK'] = pd.to_datetime(df['TGL_MASUK'], errors='coerce', format='mixed', dayfirst=True)
    df['TGL_KELUAR'] = pd.to_datetime(df['TGL_KELUAR'], errors='coerce', format='mixed', dayfirst=True)
    df = df.dropna(subset=['NOPOL', 'TGL_MASUK'])
    
    # MUNDUR DUA BULAN MAKSIMAL HINGGA END_DATE
    df = df[(df['TGL_MASUK'] >= fetch_start_date) & (df['TGL_MASUK'] <= end_date)]

    df['TGL_KELUAR_FILLED'] = df['TGL_KELUAR'].fillna(end_date)

    # ==========================================
    # 3. PEMETAAN STATUS
    # ==========================================
    def map_status(row):
        status = str(row['STATUS_BENGKEL'])
        masuk = row['TGL_MASUK']
        keluar = row['TGL_KELUAR']
        is_same_day = pd.notna(keluar) and (masuk.date() == keluar.date())
        
        if status == 'R': return 'R'
        elif 'ASURANSI' in status or 'INSURANCE' in status: return 'B - INS'
        elif 'STORING HO' in status or 'STORING MKS' in status: return 'B - INT'
        elif 'INTERNAL' in status: return 'AB - INT' if is_same_day else 'B - INT'
        elif 'EKSTERNAL' in status or 'EXTERNAL' in status: return 'AB - EXT' if is_same_day else 'B - EXT'
        else: return 'A'

    df['STATUS_CODE'] = df.apply(map_status, axis=1)

    # ==========================================
    # 4. EXPLODE TANGGAL & FILTER
    # ==========================================
    def get_date_range(row):
        try:
            s_val = pd.to_datetime(row['TGL_MASUK'])
            e_val = pd.to_datetime(row['TGL_KELUAR_FILLED'])
            
            s = max(s_val, start_date)
            e = min(e_val, end_date)
            
            if s > e: return []
            return pd.date_range(s, e).date.tolist()
        except:
            return []

    df['DATE'] = df.apply(get_date_range, axis=1)
    df = df.explode('DATE')
    df = df.dropna(subset=['DATE'])
    df['DATE'] = pd.to_datetime(df['DATE'])

    df = df.sort_values(by=['LOKASI', 'NOPOL', 'DATE', 'TGL_MASUK'])
    df = df.drop_duplicates(subset=['LOKASI', 'NOPOL', 'DATE'], keep='last')

    # ==========================================
    # 5. PIVOT & EXPORT
    # ==========================================
    list_lokasi = df['LOKASI'].unique()
    date_range = pd.date_range(start=start_date, end=end_date)
    date_cols_str = [f"{d.day}/{d.month}" for d in date_range]
    total_work_days = sum(1 for d in date_range if d.dayofweek != 6)
    if total_work_days == 0: total_work_days = 1

    identity_cols = ['BU', 'DEPT', 'LOKASI', 'JENIS_MOBIL', 'MERK', 'NOPOL', 'USIA']

    # Pastikan kolom identitas ada
    for col in identity_cols:
        if col not in df.columns:
            df[col] = '-'

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for lokasi in list_lokasi:
            df_lokasi = df[df['LOKASI'] == lokasi]
            if df_lokasi.empty: continue
                
            pivot_df = df_lokasi.pivot_table(
                index=identity_cols, 
                columns='DATE', 
                values='STATUS_CODE', 
                aggfunc='first'
            )
            
            pivot_df = pivot_df.reindex(columns=date_range).fillna('A')
            
            # Kosongkan status di hari Minggu
            for dt in date_range:
                if dt.dayofweek == 6:
                    pivot_df[dt] = ''
                    
            pivot_df.columns = date_cols_str
            
            status_list = ['A', 'AB - INT', 'AB - EXT', 'B - INT', 'B - EXT', 'B - INS', 'R']
            total_days = total_work_days
            
            for status in status_list:
                pivot_df[f'TOTAL {status}'] = (pivot_df[date_cols_str] == status).sum(axis=1)
                persentase = (pivot_df[f'TOTAL {status}'] / total_days)
                pivot_df[f'% {status}'] = persentase.map(lambda x: f"{x:.2%}")
            
            # Sortir alphabetical
            pivot_df = pivot_df.reset_index().sort_values(by='NOPOL')
            sheet_name = str(lokasi)[:31].replace('[', '').replace(']', '').replace(':', '').replace('*', '').replace('?', '').replace('\\', '').replace('/', '')
            pivot_df.to_excel(writer, sheet_name=sheet_name, index=False)
            
    return output.getvalue()
