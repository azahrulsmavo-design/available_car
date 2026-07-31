import pandas as pd
import numpy as np
import io
import re

def run_etl_ban(input_file, filename="", start_date=None, end_date=None):
    """
    Membaca data BAN SHARE, menggunakan baris 7 sebagai header (skiprows=6),
    mengekstrak Nopol, Serial Number, KM Pemasangan, KM Pengajuan, dan KM Lepas.
    Menghitung Jarak Tempuh.
    """
    if isinstance(input_file, bytes):
        input_file = io.BytesIO(input_file)
        
    try:
        if str(filename).endswith('.csv'):
            df_raw = pd.read_csv(input_file, skiprows=6, dtype=str)
        else:
            df_raw = pd.read_excel(input_file, skiprows=6, dtype=str)
    except Exception:
        # Fallback if first attempt fails
        input_file.seek(0)
        try:
            df_raw = pd.read_csv(input_file, skiprows=6, dtype=str)
        except Exception:
            input_file.seek(0)
            df_raw = pd.read_excel(input_file, skiprows=6, dtype=str)
            
    # Normalisasi nama kolom untuk memudahkan pencarian fleksibel
    columns_mapped = {c: str(c).strip().upper() for c in df_raw.columns}
    df_raw.rename(columns=columns_mapped, inplace=True)
    
    col_nopol = None
    col_serial = None
    col_km_pasang = None
    col_km_pengajuan = None
    col_km_lepas = None
    col_tgl_pasang = None
    col_tgl_pengajuan = None
    col_tgl_lepas = None

    # Pencarian kolom fleksibel
    for c in df_raw.columns:
        c_str = str(c).upper()
        if 'NOPOL' in c_str or 'NO POL' in c_str or 'POLISI' in c_str:
            if not col_nopol: col_nopol = c
        elif 'SERIAL' in c_str or 'SERI' in c_str or 'S/N' in c_str:
            if not col_serial: col_serial = c
        elif 'PASANG' in c_str or 'PEMASANGAN' in c_str:
            if 'TGL' in c_str or 'TANGGAL' in c_str:
                if not col_tgl_pasang: col_tgl_pasang = c
            elif 'KM' in c_str:
                if not col_km_pasang: col_km_pasang = c
        elif 'PENGAJUAN' in c_str:
            if 'TGL' in c_str or 'TANGGAL' in c_str:
                if not col_tgl_pengajuan: col_tgl_pengajuan = c
            elif 'KM' in c_str:
                if not col_km_pengajuan: col_km_pengajuan = c
        elif 'LEPAS' in c_str:
            if 'TGL' in c_str or 'TANGGAL' in c_str:
                if not col_tgl_lepas: col_tgl_lepas = c
            elif 'KM' in c_str:
                if not col_km_lepas: col_km_lepas = c

    # Ambil kolom yang dibutuhkan
    cols_to_extract = [col_nopol, col_serial]
    if col_km_pasang: cols_to_extract.append(col_km_pasang)
    if col_km_pengajuan: cols_to_extract.append(col_km_pengajuan)
    if col_km_lepas: cols_to_extract.append(col_km_lepas)
    if col_tgl_pasang: cols_to_extract.append(col_tgl_pasang)
    if col_tgl_pengajuan: cols_to_extract.append(col_tgl_pengajuan)
    if col_tgl_lepas: cols_to_extract.append(col_tgl_lepas)
    
    cols_to_extract = [c for c in cols_to_extract if c is not None]
    
    df = df_raw[cols_to_extract].copy()
    
    rename_dict = {
        col_nopol: 'NOPOL',
        col_serial: 'SERIAL_NUMBER'
    }
    if col_km_pasang: rename_dict[col_km_pasang] = 'KM_PEMASANGAN'
    if col_km_pengajuan: rename_dict[col_km_pengajuan] = 'KM_PENGAJUAN'
    if col_km_lepas: rename_dict[col_km_lepas] = 'KM_LEPAS'
    if col_tgl_pasang: rename_dict[col_tgl_pasang] = 'TGL_PEMASANGAN'
    if col_tgl_pengajuan: rename_dict[col_tgl_pengajuan] = 'TGL_PENGAJUAN'
    if col_tgl_lepas: rename_dict[col_tgl_lepas] = 'TGL_LEPAS'
    
    df.rename(columns=rename_dict, inplace=True)

    # Bersihkan Data
    def format_nopol(nopol):
        nopol = str(nopol).upper().replace(' ', '')
        m = re.match(r'^([A-Z]{1,2})(\d{1,4})([A-Z]{0,3})$', nopol)
        if m:
            return ' '.join([p for p in m.groups() if p])
        return nopol

    df = df.dropna(subset=['NOPOL', 'SERIAL_NUMBER'])
    df['NOPOL'] = df['NOPOL'].apply(format_nopol)
    df['SERIAL_NUMBER'] = df['SERIAL_NUMBER'].astype(str).str.strip().str.upper()

    # Ekstrak angka saja dari KM (Karena pemisah ribuan bisa berupa titik atau koma, kita hapus semua karakter selain angka)
    for col in ['KM_PEMASANGAN', 'KM_PENGAJUAN', 'KM_LEPAS']:
        if col in df.columns:
            # Ambil hanya digit, ini otomatis menghapus separator ribuan seperti titik atau koma
            df[col + '_STR'] = df[col].astype(str).str.replace(r'[^\d]', '', regex=True)
            df[col] = pd.to_numeric(df[col + '_STR'], errors='coerce').fillna(0)

    # Deteksi pola anomali seperti '12345' atau angka berulang
    def is_anomalous_pattern(val):
        val = str(val)
        if not val or val == '0': return False
        if val in ['123', '1234', '12345', '123456', '1234567', '12345678', '123456789']: return True
        if len(val) >= 3 and len(set(val)) == 1: return True
        return False

    # Deteksi Anomali (KM > 0 dan KM < 1000 atau memiliki pola tertentu)
    # Berlaku untuk semua kolom KM. Jika ada satu saja yang anomali di baris tersebut, masuk ke sheet anomali
    cond_pasang = (df['KM_PEMASANGAN'] > 0) & (df['KM_PEMASANGAN'] < 1000) if 'KM_PEMASANGAN' in df.columns else False
    cond_pengajuan = (df['KM_PENGAJUAN'] > 0) & (df['KM_PENGAJUAN'] < 1000) if 'KM_PENGAJUAN' in df.columns else False
    cond_lepas = (df['KM_LEPAS'] > 0) & (df['KM_LEPAS'] < 1000) if 'KM_LEPAS' in df.columns else False
    
    if 'KM_PEMASANGAN_STR' in df.columns:
        if isinstance(cond_pasang, bool) and cond_pasang == False: cond_pasang = df['KM_PEMASANGAN_STR'].apply(is_anomalous_pattern)
        else: cond_pasang = cond_pasang | df['KM_PEMASANGAN_STR'].apply(is_anomalous_pattern)
        
    if 'KM_PENGAJUAN_STR' in df.columns:
        if isinstance(cond_pengajuan, bool) and cond_pengajuan == False: cond_pengajuan = df['KM_PENGAJUAN_STR'].apply(is_anomalous_pattern)
        else: cond_pengajuan = cond_pengajuan | df['KM_PENGAJUAN_STR'].apply(is_anomalous_pattern)
        
    if 'KM_LEPAS_STR' in df.columns:
        if isinstance(cond_lepas, bool) and cond_lepas == False: cond_lepas = df['KM_LEPAS_STR'].apply(is_anomalous_pattern)
        else: cond_lepas = cond_lepas | df['KM_LEPAS_STR'].apply(is_anomalous_pattern)
    
    anomaly_mask = cond_pasang | cond_pengajuan | cond_lepas
    df_anomali = df[anomaly_mask].copy()
    df_valid = df[~anomaly_mask].copy()

    # Hapus kolom temporary
    for col in ['KM_PEMASANGAN_STR', 'KM_PENGAJUAN_STR', 'KM_LEPAS_STR']:
        if col in df_valid.columns: df_valid = df_valid.drop(columns=[col])
        if col in df_anomali.columns: df_anomali = df_anomali.drop(columns=[col])

    # Hitung Jarak Tempuh dan Gabung Tanggal untuk data yang valid
    if 'KM_PENGAJUAN' in df_valid.columns and 'KM_LEPAS' in df_valid.columns:
        df_valid['KM_PELEPASAN'] = df_valid[['KM_PENGAJUAN', 'KM_LEPAS']].max(axis=1)
    elif 'KM_PENGAJUAN' in df_valid.columns:
        df_valid['KM_PELEPASAN'] = df_valid['KM_PENGAJUAN']
    elif 'KM_LEPAS' in df_valid.columns:
        df_valid['KM_PELEPASAN'] = df_valid['KM_LEPAS']
    else:
        df_valid['KM_PELEPASAN'] = 0

    if 'TGL_PENGAJUAN' in df_valid.columns and 'TGL_LEPAS' in df_valid.columns:
        df_valid['TGL_PELEPASAN'] = df_valid['TGL_LEPAS'].fillna(df_valid['TGL_PENGAJUAN'])
    elif 'TGL_PENGAJUAN' in df_valid.columns:
        df_valid['TGL_PELEPASAN'] = df_valid['TGL_PENGAJUAN']
    elif 'TGL_LEPAS' in df_valid.columns:
        df_valid['TGL_PELEPASAN'] = df_valid['TGL_LEPAS']
    else:
        df_valid['TGL_PELEPASAN'] = ''

    if 'TGL_PEMASANGAN' not in df_valid.columns:
        df_valid['TGL_PEMASANGAN'] = ''

    # Pastikan serial number ada isinya
    df_valid = df_valid[df_valid['SERIAL_NUMBER'].str.len() > 0]

    # Agregasi data berdasarkan NOPOL dan SERIAL_NUMBER
    agg_df = df_valid.groupby(['NOPOL', 'SERIAL_NUMBER']).agg({
        'KM_PEMASANGAN': 'max',
        'KM_PELEPASAN': 'max',
        'TGL_PEMASANGAN': lambda x: next(iter([v for v in x if pd.notna(v) and str(v).strip() != '']), ''),
        'TGL_PELEPASAN': lambda x: next(iter([v for v in x if pd.notna(v) and str(v).strip() != '']), '')
    }).reset_index()

    # Tentukan status kelengkapan data
    def get_status(row):
        pasang = row['KM_PEMASANGAN'] > 0
        lepas = row['KM_PELEPASAN'] > 0
        if pasang and lepas:
            return 'Lengkap'
        elif pasang and not lepas:
            return 'Hanya Pemasangan'
        elif not pasang and lepas:
            return 'Hanya Pelepasan'
        else:
            return 'Tidak Ada KM'
            
    agg_df['STATUS_DATA'] = agg_df.apply(get_status, axis=1)
    
    # Deteksi anomali ban berbeda (KM / Tanggal mundur, sama, atau Jarak Tempuh <= 0)
    def is_invalid_pair(row):
        if row['STATUS_DATA'] == 'Lengkap':
            if row['KM_PEMASANGAN'] >= row['KM_PELEPASAN']:
                return True
            try:
                tgl_pasang = pd.to_datetime(row['TGL_PEMASANGAN'], dayfirst=True)
                tgl_lepas = pd.to_datetime(row['TGL_PELEPASAN'], dayfirst=True)
                if tgl_pasang >= tgl_lepas:
                    return True
            except:
                pass
        return False

    invalid_mask = agg_df.apply(is_invalid_pair, axis=1)
    
    if invalid_mask.any():
        df_invalid = agg_df[invalid_mask].copy()
        
        # Pisah jadi Hanya Pemasangan
        df_pasang_split = df_invalid.copy()
        df_pasang_split['KM_PELEPASAN'] = 0
        df_pasang_split['TGL_PELEPASAN'] = ''
        df_pasang_split['STATUS_DATA'] = 'Hanya Pemasangan'
        
        # Pisah jadi Hanya Pelepasan
        df_lepas_split = df_invalid.copy()
        df_lepas_split['KM_PEMASANGAN'] = 0
        df_lepas_split['TGL_PEMASANGAN'] = ''
        df_lepas_split['STATUS_DATA'] = 'Hanya Pelepasan'
        
        agg_df = agg_df[~invalid_mask]
        agg_df = pd.concat([agg_df, df_pasang_split, df_lepas_split], ignore_index=True)

    # Hitung Jarak Tempuh
    agg_df['JARAK_TEMPUH'] = agg_df['KM_PELEPASAN'] - agg_df['KM_PEMASANGAN']
    
    # Nol-kan Jarak Tempuh jika data tidak lengkap
    agg_df.loc[agg_df['STATUS_DATA'] != 'Lengkap', 'JARAK_TEMPUH'] = 0

    # Filter berdasarkan rentang tanggal pelepasan jika start_date dan end_date diberikan
    if start_date and end_date:
        start_pd = pd.to_datetime(start_date)
        end_pd = pd.to_datetime(end_date)
        
        def is_within_range(val):
            if pd.isna(val) or str(val).strip() == '':
                return False
            try:
                dt = pd.to_datetime(val, dayfirst=True)
                return start_pd <= dt <= end_pd
            except:
                return False
                
        date_mask = agg_df['TGL_PELEPASAN'].apply(is_within_range)
        # Selalu sertakan 'Hanya Pemasangan' karena mereka belum memiliki tanggal pelepasan
        date_mask = date_mask | (agg_df['STATUS_DATA'] == 'Hanya Pemasangan')
        agg_df = agg_df[date_mask]

    # Urutkan berdasarkan Nopol dan Serial Number
    agg_df = agg_df.sort_values(by=['NOPOL', 'SERIAL_NUMBER', 'JARAK_TEMPUH'], ascending=[True, True, False])
    
    # Pisahkan menjadi beberapa dataframe berdasarkan status
    df_lengkap = agg_df[agg_df['STATUS_DATA'] == 'Lengkap']
    df_pemasangan = agg_df[agg_df['STATUS_DATA'] == 'Hanya Pemasangan']
    df_pelepasan = agg_df[agg_df['STATUS_DATA'] == 'Hanya Pelepasan']
    df_tidak_ada = agg_df[agg_df['STATUS_DATA'] == 'Tidak Ada KM']
    
    # Rapikan df_anomali agar format kolomnya selaras dengan sheet lain
    if not df_anomali.empty:
        # Hitung KM Pelepasan untuk anomali
        if 'KM_PENGAJUAN' in df_anomali.columns and 'KM_LEPAS' in df_anomali.columns:
            df_anomali['KM_PELEPASAN'] = df_anomali[['KM_PENGAJUAN', 'KM_LEPAS']].max(axis=1)
        elif 'KM_PENGAJUAN' in df_anomali.columns:
            df_anomali['KM_PELEPASAN'] = df_anomali['KM_PENGAJUAN']
        elif 'KM_LEPAS' in df_anomali.columns:
            df_anomali['KM_PELEPASAN'] = df_anomali['KM_LEPAS']
        else:
            df_anomali['KM_PELEPASAN'] = 0
            
        # TGL Pelepasan untuk anomali
        if 'TGL_PENGAJUAN' in df_anomali.columns and 'TGL_LEPAS' in df_anomali.columns:
            df_anomali['TGL_PELEPASAN'] = df_anomali['TGL_LEPAS'].fillna(df_anomali['TGL_PENGAJUAN'])
        elif 'TGL_PENGAJUAN' in df_anomali.columns:
            df_anomali['TGL_PELEPASAN'] = df_anomali['TGL_PENGAJUAN']
        elif 'TGL_LEPAS' in df_anomali.columns:
            df_anomali['TGL_PELEPASAN'] = df_anomali['TGL_LEPAS']
        else:
            df_anomali['TGL_PELEPASAN'] = ''

        if 'TGL_PEMASANGAN' not in df_anomali.columns:
            df_anomali['TGL_PEMASANGAN'] = ''
            
        df_anomali['STATUS_DATA'] = 'Anomali (< 1000 KM atau Pola Tidak Valid)'
        df_anomali['JARAK_TEMPUH'] = 0
        
        # Ambil kolom yang konsisten dengan sheet lain
        cols_anomali = ['NOPOL', 'SERIAL_NUMBER', 'KM_PEMASANGAN', 'KM_PELEPASAN', 'TGL_PEMASANGAN', 'TGL_PELEPASAN', 'STATUS_DATA', 'JARAK_TEMPUH']
        df_anomali = df_anomali[[c for c in cols_anomali if c in df_anomali.columns]]
        
        if start_date and end_date:
            start_pd = pd.to_datetime(start_date)
            end_pd = pd.to_datetime(end_date)
            
            def is_within_range_anomali(val):
                if pd.isna(val) or str(val).strip() == '':
                    return False
                try:
                    dt = pd.to_datetime(val, dayfirst=True)
                    return start_pd <= dt <= end_pd
                except:
                    return False
                    
            date_mask_anomali = df_anomali['TGL_PELEPASAN'].apply(is_within_range_anomali)
            # Selalu sertakan anomali yang belum memiliki tanggal pelepasan
            date_mask_anomali = date_mask_anomali | (df_anomali['TGL_PELEPASAN'] == '')
            df_anomali = df_anomali[date_mask_anomali]
    
    # Export ke format Excel di memory dengan beberapa sheet
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not df_lengkap.empty:
            df_lengkap.to_excel(writer, sheet_name='Lengkap', index=False)
        if not df_pemasangan.empty:
            df_pemasangan.to_excel(writer, sheet_name='Hanya Pemasangan', index=False)
        if not df_pelepasan.empty:
            df_pelepasan.to_excel(writer, sheet_name='Hanya Pelepasan', index=False)
        if not df_tidak_ada.empty:
            df_tidak_ada.to_excel(writer, sheet_name='Tidak Ada KM', index=False)
        if not df_anomali.empty:
            df_anomali.to_excel(writer, sheet_name='Anomali', index=False)
            
        # Format lebar kolom
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('A:A', 15)  # Nopol
            worksheet.set_column('B:B', 30)  # Serial Number
            worksheet.set_column('C:D', 20)  # KM
            worksheet.set_column('E:F', 20)  # TGL
            worksheet.set_column('G:G', 25)  # Status
            worksheet.set_column('H:H', 20)  # Jarak Tempuh
    
    # Untuk return df yang ditampilkan di UI, kita bisa gabung yang valid saja atau semuanya
    return output.getvalue(), agg_df
