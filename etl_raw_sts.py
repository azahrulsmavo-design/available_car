import pandas as pd
import numpy as np
from datetime import datetime
import io

def run_etl_raw(target_start_date_str, target_end_date_str, input_file):
    start_date = pd.to_datetime(target_start_date_str)
    end_date = pd.to_datetime(target_end_date_str)
    
    # 1 MONTH LOOKBACK
    fetch_start_date = start_date - pd.DateOffset(months=1)

    df_raw = pd.read_csv(input_file, skiprows=3, low_memory=False)
    df_raw.columns = [str(c).replace('\n', ' ').strip() for c in df_raw.columns]

    # ==========================================
    # 2. EKSTRAKSI DATA
    # ==========================================
    cols_to_keep = {
        'BU MASTER': 'BU',
        'DEPT': 'DEPT',
        'CABANG MASTER': 'LOKASI',
        'JENIS KENDARAAN': 'JENIS_MOBIL',
        'MEREK KENDARAAN': 'MERK',
        'NOPOL': 'NOPOL',
        'USIA KENDARAAN': 'USIA',
        'STATUS BENGKEL': 'STATUS_BENGKEL',
        'TGL MASUK BENGKEL': 'TGL_MASUK',
        'TGLKELUAR BENGKEL': 'TGL_KELUAR'
    }

    available_cols = [c for c in cols_to_keep.keys() if c in df_raw.columns]
    df = df_raw[available_cols].copy()
    df.rename(columns=cols_to_keep, inplace=True)

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
