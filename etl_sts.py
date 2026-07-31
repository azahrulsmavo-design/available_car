import pandas as pd
import numpy as np
import io
import datetime

def run_etl_sts(file_bytes, filename, start_date, end_date):
    """
    Menjalankan ETL untuk file STS.
    """
    # Deteksi format file
    try:
        if filename.endswith('.csv'):
            df_raw = None
            for skip in [3, 17, 0, 1, 2]:
                temp_df = pd.read_csv(io.BytesIO(file_bytes), skiprows=skip, dtype=str, on_bad_lines='skip')
                if 'Timestamp' in temp_df.columns or any('TIMESTAMP' in str(c).upper() for c in temp_df.columns):
                    df_raw = temp_df
                    break
            if df_raw is None:
                # Fallback to no skip if loop didn't find it
                df_raw = pd.read_csv(io.BytesIO(file_bytes), dtype=str, on_bad_lines='skip')
                
        elif filename.endswith('.xlsx') or filename.endswith('.xls'):
            df_raw = None
            for skip in [3, 17, 0, 1, 2]:
                try:
                    temp_df = pd.read_excel(io.BytesIO(file_bytes), skiprows=skip, dtype=str)
                    if 'Timestamp' in temp_df.columns or any('TIMESTAMP' in str(c).upper() for c in temp_df.columns):
                        df_raw = temp_df
                        break
                except:
                    continue
            if df_raw is None:
                df_raw = pd.read_excel(io.BytesIO(file_bytes), dtype=str)
        else:
            raise ValueError("Format file tidak didukung. Gunakan CSV atau Excel.")
    except Exception as e:
        raise Exception(f"Gagal membaca file: {str(e)}")

    # Mapping Kolom
    col_nopol = None
    col_tgl = None
    col_km = None

    for c in df_raw.columns:
        c_str = str(c).upper().strip()
        if 'NOPOL' in c_str or 'NO POL' in c_str or 'POLISI' in c_str:
            if not col_nopol: col_nopol = c
        elif 'TIMESTAMP' in c_str or 'TANGGAL' in c_str:
            # Utamakan Timestamp jika ada
            if 'TIMESTAMP' in c_str: col_tgl = c
            elif not col_tgl: col_tgl = c
        elif 'KM KENDARAAN' in c_str or 'KM' in c_str:
            if 'KM KENDARAAN' in c_str: col_km = c
            elif not col_km: col_km = c

    col_id = None
    for c in df_raw.columns:
        c_str = str(c).upper().strip()
        if c_str == 'ID' or c_str == 'ID FS':
            col_id = c
            break

    if not all([col_nopol, col_tgl, col_km]):
        raise ValueError(f"Gagal menemukan kolom NOPOL, Timestamp, atau KM KENDARAAN. Kolom ditemukan: NOPOL={col_nopol}, TGL={col_tgl}, KM={col_km}. Daftar Kolom: {df_raw.columns.tolist()[:15]}")

    if col_id:
        df = df_raw[[col_nopol, col_tgl, col_km, col_id]].copy()
        df.columns = ['NOPOL', 'TIMESTAMP', 'KM', 'ID']
    else:
        df = df_raw[[col_nopol, col_tgl, col_km]].copy()
        df.columns = ['NOPOL', 'TIMESTAMP', 'KM']
        df['ID'] = ''

    df = df.dropna(subset=['NOPOL', 'TIMESTAMP'])

    # Cleaning Nopol
    df['NOPOL'] = df['NOPOL'].astype(str).str.replace(' ', '').str.upper()

    # Cleaning KM
    df['KM_STR'] = df['KM'].astype(str).str.replace(r'[^\d]', '', regex=True)
    df['KM'] = pd.to_numeric(df['KM_STR'], errors='coerce').fillna(0)

    # Pisahkan Anomali (KM < 1000 atau KM > 2000000 dianggap tidak wajar, atau pola ketikan salah)
    bad_exacts = ['123', '123456', '123455', '12345678', '123456789', '123461', '1234566', '879387', '1233456', '1234567']
    pola_anomali = df['KM_STR'].isin(bad_exacts) | df['KM_STR'].str.endswith('123') | df['KM_STR'].str.endswith('123456') | df['KM_STR'].str.startswith('123')
    
    cond_anomali = (df['KM'] > 0) & ((df['KM'] < 1000) | (df['KM'] > 2000000) | pola_anomali)
    df_anomali = df[cond_anomali].copy()
    df_valid = df[~cond_anomali].copy()
    
    # Filter valid KM (pastikan yang 0 juga dibuang dari valid jika terlewat)
    df_valid = df_valid[(df_valid['KM'] >= 1000) & (df_valid['KM'] <= 2000000)]

    # Parsing Tanggal
    # Format di file STS biasanya DD/MM/YYYY HH:MM:SS, kita gunakan dayfirst=True
    df_valid['TIMESTAMP'] = pd.to_datetime(df_valid['TIMESTAMP'], dayfirst=True, errors='coerce')
    df_valid = df_valid.dropna(subset=['TIMESTAMP'])
    
    df_anomali['TIMESTAMP'] = pd.to_datetime(df_anomali['TIMESTAMP'], dayfirst=True, errors='coerce')

    # Siapkan datetime filter
    # start_date dan end_date adalah datetime.date. Kita ubah ke pd.Timestamp
    start_dt = pd.to_datetime(start_date) + pd.Timedelta(days=1) - pd.Timedelta(seconds=1) # Akhir hari start_date
    end_dt = pd.to_datetime(end_date) # Awal hari end_date

    # Urutkan
    df_valid = df_valid.sort_values(by=['NOPOL', 'TIMESTAMP'])

    results = []

    nopols = df_valid['NOPOL'].unique()
    for nopol in nopols:
        df_nopol = df_valid[df_valid['NOPOL'] == nopol].copy()
        
        # --- DETEKSI ANOMALI HISTORIS (FLUKTUATIF) ---
        df_nopol = df_nopol.sort_values(by='TIMESTAMP')
        diff_km = df_nopol['KM'].diff()
        
        jumps_down = (diff_km < -5000).sum()
        jumps_up_extreme = (diff_km > 200000).sum()
        
        history_anomalous = False
        if (jumps_down + jumps_up_extreme) >= 2:
            history_anomalous = True
        # ---------------------------------------------
        
        median_km = df_nopol['KM'].median()
        jarak = 0
        
        while len(df_nopol) > 0:
            # Cari tanggal paling mendekati start_dt (Absolute Closest)
            idx_start = (df_nopol['TIMESTAMP'] - start_dt).abs().idxmin()
            row_start = df_nopol.loc[idx_start]
            km_start = row_start['KM']
            tgl_start = row_start['TIMESTAMP']
            id_start = row_start['ID']

            # Cari tanggal paling mendekati end_dt (Absolute Closest)
            idx_end = (df_nopol['TIMESTAMP'] - end_dt).abs().idxmin()
            row_end = df_nopol.loc[idx_end]
            km_end = row_end['KM']
            tgl_end = row_end['TIMESTAMP']
            id_end = row_end['ID']

            if idx_start == idx_end:
                jarak = 0
                break
                
            jarak_temp = km_end - km_start
            
            is_anomaly = False
            
            # 1. Tanggal start bernilai sama atau lebih besar dari Tanggal Akhir
            if tgl_start >= tgl_end:
                is_anomaly = True
                
            # 2. Jarak capaian KM bernilai 0 atau negatif
            if jarak_temp <= 0:
                is_anomaly = True
                
            # 3. Angka anomali karena salah ketik (Capaian tidak masuk akal, misal >= 1,000,000 KM)
            if jarak_temp >= 1000000:
                is_anomaly = True
                
            if is_anomaly:
                # Buang data yang salah (yang nilainya lebih jauh dari median KM) dan coba lagi
                if abs(km_start - median_km) > abs(km_end - median_km):
                    df_nopol = df_nopol.drop(idx_start)
                else:
                    df_nopol = df_nopol.drop(idx_end)
            else:
                jarak = jarak_temp
                break

        if len(df_nopol) == 0:
            km_start, km_end, tgl_start, tgl_end, id_start, id_end = 0, 0, None, None, '', ''
            jarak = 0
            status = 'Data Tidak Lengkap / Terindikasi Anomali Semua'
        else:
            ket_awal = '' if tgl_start <= start_dt else '(Setelah Tgl Awal)'
            ket_akhir = '' if tgl_end >= end_dt else '(Sebelum Tgl Akhir)'
            status = 'Lengkap' if (km_start > 0 and km_end > 0) else 'Data Tidak Lengkap di Rentang Waktu'
            
            tambahan = f"{ket_awal} {ket_akhir}".strip()
            if tambahan and km_start > 0 and km_end > 0:
                status = f"{status} {tambahan}"
                
        if history_anomalous:
            status = f"Anomali (Riwayat Fluktuatif) - {status}"
        
        results.append({
            'NOPOL': nopol,
            'TANGGAL AWAL DITEMUKAN': tgl_start.strftime('%d/%m/%Y %H:%M:%S') if tgl_start else 'Tidak Ada',
            'ID FS AWAL': id_start,
            'KM AWAL': km_start,
            'TANGGAL AKHIR DITEMUKAN': tgl_end.strftime('%d/%m/%Y %H:%M:%S') if tgl_end else 'Tidak Ada',
            'ID FS AKHIR': id_end,
            'KM AKHIR': km_end,
            'STATUS': status,
            'CAPAIAN KM': jarak
        })

    result_df = pd.DataFrame(results)
    
    import re
    def format_nopol(n):
        # Format nopol: 1-2 huruf, 1-4 angka, 0-3 huruf
        m = re.match(r'^([A-Z]{1,2})(\d{1,4})([A-Z]{0,3})$', str(n))
        if m:
            return f"{m.group(1)} {m.group(2)} {m.group(3)}".strip()
        return str(n)

    # Sort & Format
    if not result_df.empty:
        result_df['NOPOL'] = result_df['NOPOL'].apply(format_nopol)
        result_df = result_df.sort_values(by=['CAPAIAN KM', 'NOPOL'], ascending=[False, True])

    # Export Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        if not result_df.empty:
            df_lengkap = result_df[result_df['STATUS'] == 'Lengkap']
            df_tidak_lengkap = result_df[result_df['STATUS'] != 'Lengkap']
            
            if not df_lengkap.empty:
                df_lengkap.to_excel(writer, sheet_name='Data Lengkap', index=False)
            if not df_tidak_lengkap.empty:
                df_tidak_lengkap.to_excel(writer, sheet_name='Data Tidak Lengkap', index=False)
        else:
            # Fallback jika kosong semua
            result_df.to_excel(writer, sheet_name='Data', index=False)
            
        if not df_anomali.empty:
            df_anomali['TIMESTAMP'] = df_anomali['TIMESTAMP'].dt.strftime('%d/%m/%Y %H:%M:%S')
            df_anomali.to_excel(writer, sheet_name='Anomali', index=False)

        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            worksheet.set_column('A:A', 15)  # Nopol
            worksheet.set_column('B:B', 25)  # Tgl Awal
            worksheet.set_column('C:C', 15)  # ID Awal
            worksheet.set_column('D:D', 15)  # KM Awal
            worksheet.set_column('E:E', 25)  # Tgl Akhir
            worksheet.set_column('F:F', 15)  # ID Akhir
            worksheet.set_column('G:G', 15)  # KM Akhir
            worksheet.set_column('H:H', 30)  # Status
            worksheet.set_column('I:I', 20)  # Capaian KM
            
    return output.getvalue(), result_df
