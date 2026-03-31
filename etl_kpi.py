import pandas as pd
import numpy as np
from datetime import datetime
from custom_order import CUSTOM_NOPOL_ORDER
import io

def run_etl_kpi(target_start_date_str, target_end_date_str, input_file, master_file):
    start_date = pd.to_datetime(target_start_date_str)
    end_date = pd.to_datetime(target_end_date_str)
    
    # 1 YEAR LOOKBACK
    # Agar data servis yang belum ditutup bisa terdeteksi
    fetch_start_date = start_date - pd.DateOffset(years=1)
    
    # ==========================================
    # 2. LOAD MASTER DATA (Source of Truth)
    # ==========================================
    if master_file.endswith('.csv'):
        # Tangani file CSV master format baru
        df_master_all = pd.read_csv(master_file, low_memory=False)
        # Jika header nyasar (baris pertama kosong/title), baca ulang dengan skiprows=1
        if 'NOPOL' not in df_master_all.columns and 'DEPARTMENT' not in df_master_all.columns and 'BU' not in df_master_all.columns:
            df_master_all = pd.read_csv(master_file, skiprows=1, low_memory=False)
    else:
        xl = pd.ExcelFile(master_file)
        target_sheet = None
        # Cari sheet yang mengandung keyword 'MASTER'
        for s in xl.sheet_names:
            s_up = s.upper()
            if 'MASTER OKT' in s_up or 'MASTER ASET' in s_up or s_up == 'MASTER' or 'MASTER' in s_up:
                target_sheet = s
                break
        if not target_sheet: 
            target_sheet = xl.sheet_names[0]
        df_master_all = xl.parse(target_sheet)
        
    df_master_all.columns = [str(c).replace('\n', ' ').strip().upper() for c in df_master_all.columns]
    
    # Deteksi penamaan kolom baru atau lama (Excel vs CSV Text yg diberikan pengguna)
    master_cols_map = {}
    if 'DEPARTMENT' in df_master_all.columns and 'LOCATION' in df_master_all.columns:
        # Format CSV baru
        master_cols_map = {
            'NOPOL': 'NOPOL',
            'DEPARTMENT': 'BU', 
            'SECTION': 'DEPT',
            'LOCATION': 'LOKASI',
            'JENIS  MOBIL': 'JENIS_MOBIL', # Sesuai teks user (ada double spasi)
            'MERK': 'MERK',
            'TAHUN PEMBUATAN': 'TAHUN_PEMBUATAN',
            'USIA': 'USIA'
        }
        # Coba perbaiki spasi ganda
        for col in df_master_all.columns:
            if 'JENIS' in col and 'MOBIL' in col:
                master_cols_map[col] = 'JENIS_MOBIL'
    else:
        # Format Excel Lama / CSV OKT ex-Excel
        master_cols_map = {
            'NOPOL': 'NOPOL',
            'BU': 'BU',
            'DEPT': 'DEPT',
            'JENIS MOBIL': 'JENIS_MOBIL',
            'MERK': 'MERK',
            'TAHUN PEMBUATAN': 'TAHUN_PEMBUATAN',
            'USIA': 'USIA'
        }
        if 'DETAIL LOCATION' in df_master_all.columns:
            master_cols_map['DETAIL LOCATION'] = 'LOKASI'
        elif 'LOKASI' in df_master_all.columns:
            master_cols_map['LOKASI'] = 'LOKASI'
        elif 'LOCATION' in df_master_all.columns:
             master_cols_map['LOCATION'] = 'LOKASI'
    
    available_master_cols = [c for c in master_cols_map.keys() if c in df_master_all.columns]
    df_master = df_master_all[available_master_cols].copy()
    df_master.rename(columns=master_cols_map, inplace=True)
    
    # Validasi jika masih ada kurang
    for missing_col in ['NOPOL', 'BU', 'DEPT', 'LOKASI', 'JENIS_MOBIL', 'MERK', 'TAHUN_PEMBUATAN', 'USIA']:
        if missing_col not in df_master.columns:
            df_master[missing_col] = '-'
            
    for col in ['NOPOL', 'BU', 'DEPT', 'LOKASI', 'JENIS_MOBIL', 'MERK']:
        if col in df_master.columns:
            df_master[col] = df_master[col].astype(str).str.strip().str.upper()
    
    # ==========================================
    # 3. EKSTRAKSI & PEMBERSIHAN DATA SERVIS (CSV)
    # ==========================================
    df_raw = pd.read_csv(input_file, skiprows=3, low_memory=False)
    df_raw.columns = [str(c).replace('\n', ' ').strip() for c in df_raw.columns]
    
    csv_cols_map = {
        'NOPOL': 'NOPOL',
        'STATUS BENGKEL': 'STATUS_BENGKEL',
        'TGL MASUK BENGKEL': 'TGL_MASUK',
        'TGLKELUAR BENGKEL': 'TGL_KELUAR'
    }
    available_csv_cols = [c for c in csv_cols_map.keys() if c in df_raw.columns]
    df_servis = df_raw[available_csv_cols].copy()
    df_servis.rename(columns=csv_cols_map, inplace=True)
    
    df_servis['NOPOL'] = df_servis['NOPOL'].astype(str).str.strip().str.upper()
    df_servis['STATUS_BENGKEL'] = df_servis['STATUS_BENGKEL'].astype(str).str.strip().str.upper()
    
    df_servis['TGL_MASUK'] = pd.to_datetime(df_servis['TGL_MASUK'], errors='coerce', format='mixed', dayfirst=True)
    df_servis['TGL_KELUAR'] = pd.to_datetime(df_servis['TGL_KELUAR'], errors='coerce', format='mixed', dayfirst=True)
    df_servis = df_servis.dropna(subset=['NOPOL', 'TGL_MASUK'])
    
    # Filter TGL_MASUK (Mundur maksimal 2 bulan hingga end_date)
    df_servis = df_servis[(df_servis['TGL_MASUK'] >= fetch_start_date) & (df_servis['TGL_MASUK'] <= end_date)]
    
    df_servis['TGL_KELUAR_FILLED'] = df_servis['TGL_KELUAR'].fillna(end_date)
    
    # ==========================================
    # 4. PEMETAAN STATUS
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
    
    df_servis['STATUS_CODE'] = df_servis.apply(map_status, axis=1)
    
    # ==========================================
    # 5. EXPLODE TANGGAL & FILTER
    # ==========================================
    def get_date_range(row):
        try:
            s_val = pd.to_datetime(row['TGL_MASUK'])
            e_val = pd.to_datetime(row['TGL_KELUAR_FILLED'])
            
            # Kita hanya generate baris untuk TANGGAL DISPLAY TARGET, 
            # menghindari komputasi ribuan kombinasi bulan-bulan sblmnya jika tidak diperlukan
            s = max(s_val, start_date)
            e = min(e_val, end_date)
            
            if s > e: return []
            return pd.date_range(s, e).date.tolist()
        except:
            return []
    
    df_servis['DATE'] = df_servis.apply(get_date_range, axis=1)
    df_servis = df_servis.explode('DATE')
    df_servis = df_servis.dropna(subset=['DATE'])
    df_servis['DATE'] = pd.to_datetime(df_servis['DATE'])
    
    df_servis = df_servis.sort_values(by=['NOPOL', 'DATE', 'TGL_MASUK'])
    df_servis = df_servis.drop_duplicates(subset=['NOPOL', 'DATE'], keep='last')
    
    # ==========================================
    # 6. MENGGUNAKAN CUSTOM ORDER SEBAGAI MASTER JAWABAN
    # ==========================================
    date_range = pd.date_range(start=start_date, end=end_date)
    
    # Kumpulkan SEMUA NOPOL dari custom order secara sekuensial
    master_nopol_list = []
    custom_sheet_map = {}
    
    for sheet_name, nopol_list in CUSTOM_NOPOL_ORDER.items():
        for nopol in nopol_list:
            clean_n = nopol.replace(' ', '').upper()
            master_nopol_list.append(clean_n)
            
            if clean_n not in custom_sheet_map:
                custom_sheet_map[clean_n] = []
            if sheet_name not in custom_sheet_map[clean_n]:
                custom_sheet_map[clean_n].append(sheet_name)
            
    # Buang duplikat jika user tanpa sengaja menginput ganda (pertahankan kemunculan pertama)
    seen = set()
    unique_nopol_list = []
    for n in master_nopol_list:
        if n not in seen:
            unique_nopol_list.append(n)
            seen.add(n)
    
    full_grid = pd.MultiIndex.from_product([unique_nopol_list, date_range], names=['_CLEAN_NOPOL', 'DATE']).to_frame(index=False)
    
    # Persiapkan df_master untuk join dengan _CLEAN_NOPOL
    df_master['_CLEAN_NOPOL'] = df_master['NOPOL'].astype(str).str.replace(' ', '').str.upper()
    
    # Join identitas kendaraan dari Excel Master (Drop duplikat master nopol jika ada)
    df_master_unique = df_master.drop_duplicates(subset=['_CLEAN_NOPOL'], keep='last')
    
    # Gabungkan grid kita dengan detail identitas kendaraan
    df_final = pd.merge(full_grid, df_master_unique, on='_CLEAN_NOPOL', how='left')
    
    # Assign ulang NOPOL asli berdasarkan list order user (prioritas: Order -> Master -> Raw)
    def revert_nopol(row):
        clean_n = row['_CLEAN_NOPOL']
        master_n = row['NOPOL']
        
        # 1. Cek di CUSTOM_NOPOL_ORDER
        for s, lst in CUSTOM_NOPOL_ORDER.items():
            for original in lst:
                if original.replace(' ', '').upper() == clean_n:
                    return original
                    
        # 2. Cek di Master (apabila file master punya format spasi)
        if pd.notna(master_n) and master_n != '-':
            return master_n
            
        # 3. Fallback
        return clean_n
        
    df_final['NOPOL_FIXED'] = df_final.apply(revert_nopol, axis=1)
    df_final['NOPOL'] = df_final['NOPOL_FIXED']
    df_final = df_final.drop(columns=['NOPOL_FIXED'])
    
    # Join Data Servis (Raw STS) -- pastikan data servis punya kolom _CLEAN_NOPOL
    df_servis['_CLEAN_NOPOL'] = df_servis['NOPOL'].astype(str).str.replace(' ', '').str.upper()
    df_final = pd.merge(df_final, df_servis[['_CLEAN_NOPOL', 'DATE', 'STATUS_CODE']], on=['_CLEAN_NOPOL', 'DATE'], how='left')
    df_final = df_final.drop(columns=['_CLEAN_NOPOL'])
    df_final['STATUS_CODE'] = df_final['STATUS_CODE'].fillna('A')
    
    # Kosongkan status di hari Minggu
    df_final.loc[df_final['DATE'].dt.dayofweek == 6, 'STATUS_CODE'] = ''
    
    # ==========================================
    # 7. LOGIKA PENAMAAN SHEET & ISI NAN
    # ==========================================
    # Isi NaN pada kolom identitas (didapat dari Master API yang tak match)
    for col in ['BU', 'DEPT', 'LOKASI', 'JENIS_MOBIL', 'MERK', 'TAHUN_PEMBUATAN', 'USIA']:
        df_final[col] = df_final[col].fillna('-')
        
    def assign_sheet(row):
        nopol_clean = str(row['NOPOL']).replace(' ', '').strip().upper()
        return custom_sheet_map.get(nopol_clean, ['OTHERS'])
    
    df_final['SHEET_NAME'] = df_final.apply(assign_sheet, axis=1)
    df_final = df_final.explode('SHEET_NAME')
    
    # ==========================================
    # 8. EXPORT KE EXCEL MEMORY
    # ==========================================
    date_cols_str = [f"{d.day}/{d.month}" for d in date_range]
    identity_cols = ['BU', 'DEPT', 'LOKASI', 'JENIS_MOBIL', 'MERK', 'TAHUN_PEMBUATAN', 'NOPOL', 'USIA']
    total_work_days = sum(1 for d in date_range if d.dayofweek != 6)
    if total_work_days == 0: total_work_days = 1
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Gunakan list keys dari custom_order untuk nama sheet (agar dinamis & pasti match)
        target_sheets = list(CUSTOM_NOPOL_ORDER.keys())
        # Pastikan tidak ada "OTHERS" di list yang tidak perlu
        if 'OTHERS' not in target_sheets and not df_final[df_final['SHEET_NAME'] == 'OTHERS'].empty:
            target_sheets.append('OTHERS')
            
        for sheet in target_sheets:
            df_sheet = df_final[df_final['SHEET_NAME'] == sheet]
            if df_sheet.empty: continue
                
            pivot_df = df_sheet.pivot_table(
                index=identity_cols, 
                columns='DATE', 
                values='STATUS_CODE',
                aggfunc='first'
            )
            
            # Reset pivot sementara untuk sorting manual yang 100% konsisten
            pivot_df = pivot_df.reset_index()
            
            if sheet in CUSTOM_NOPOL_ORDER:
                custom_list = CUSTOM_NOPOL_ORDER[sheet]
                pivot_df['_CLEAN_NOPOL'] = pivot_df['NOPOL'].astype(str).str.replace(' ', '').str.upper()
                order_dict = {nopol.replace(' ', '').upper(): i for i, nopol in enumerate(custom_list)}
                pivot_df['CUSTOM_ORDER'] = pivot_df['_CLEAN_NOPOL'].map(lambda x: order_dict.get(x, 999999))
                pivot_df = pivot_df.sort_values(['CUSTOM_ORDER'])
                pivot_df = pivot_df.drop(columns=['CUSTOM_ORDER', '_CLEAN_NOPOL'])
            else:
                pivot_df = pivot_df.sort_values(['NOPOL'])
                
            pivot_df = pivot_df.set_index(identity_cols)
            
            pivot_df = pivot_df.reset_index()
            pivot_df.columns = identity_cols + date_cols_str
            
            status_list = ['A', 'AB - INT', 'AB - EXT', 'B - INT', 'B - EXT', 'B - INS', 'R']
            total_days = total_work_days
            for s in status_list:
                pivot_df[f'TOTAL {s}'] = (pivot_df[date_cols_str] == s).sum(axis=1)
                pivot_df[f'% {s}'] = (pivot_df[f'TOTAL {s}'] / total_days).map(lambda x: f"{x:.2%}")
                
            pivot_df.to_excel(writer, sheet_name=sheet, index=False)
            
        # ==========================================
        # 9. SHEET INVESTIGASI (Status Sama Sepanjang Periode)
        # ==========================================
        inv_df = df_final.groupby('NOPOL')['STATUS_CODE'].nunique().reset_index()
        inv_nopols = inv_df[inv_df['STATUS_CODE'] == 1]['NOPOL']
        
        df_investigasi = df_final[df_final['NOPOL'].isin(inv_nopols)]
        if not df_investigasi.empty:
            pivot_inv = df_investigasi.pivot_table(
                index=identity_cols,
                columns='DATE',
                values='STATUS_CODE',
                aggfunc='first'
            )
            pivot_inv = pivot_inv.reset_index()
            
            # Sort menggunakan urutan global dari custom order
            pivot_inv['_CLEAN_NOPOL'] = pivot_inv['NOPOL'].astype(str).str.replace(' ', '').str.upper()
            all_orders = []
            for lst in CUSTOM_NOPOL_ORDER.values():
                all_orders.extend(lst)
            order_dict = {nopol.replace(' ', '').upper(): i for i, nopol in enumerate(all_orders)}
            
            pivot_inv['CUSTOM_ORDER'] = pivot_inv['_CLEAN_NOPOL'].map(lambda x: order_dict.get(x, 999999))
            pivot_inv = pivot_inv.sort_values(['CUSTOM_ORDER', 'NOPOL'])
            pivot_inv = pivot_inv.drop(columns=['CUSTOM_ORDER', '_CLEAN_NOPOL'])
            
            pivot_inv = pivot_inv.set_index(identity_cols)
            pivot_inv = pivot_inv.reset_index()
            pivot_inv.columns = identity_cols + date_cols_str
            
            status_list = ['A', 'AB - INT', 'AB - EXT', 'B - INT', 'B - EXT', 'B - INS', 'R']
            total_days = total_work_days
            for s in status_list:
                pivot_inv[f'TOTAL {s}'] = (pivot_inv[date_cols_str] == s).sum(axis=1)
                pivot_inv[f'% {s}'] = (pivot_inv[f'TOTAL {s}'] / total_days).map(lambda x: f"{x:.2%}")
                
            pivot_inv.to_excel(writer, sheet_name='INVESTIGASI', index=False)
            
            # ==========================================
            # 9.5 SHEET INVESTIGASI_NON_A (Status Sama, Bukan A, dengan Early Start)
            # ==========================================
            first_status = df_investigasi.groupby('NOPOL')['STATUS_CODE'].first()
            non_a_nopols = first_status[first_status != 'A'].index
            df_inv_nona = df_investigasi[df_investigasi['NOPOL'].isin(non_a_nopols)]
            
            if not df_inv_nona.empty:
                pivot_nona = df_inv_nona.pivot_table(
                    index=identity_cols, columns='DATE', values='STATUS_CODE', aggfunc='first'
                ).reset_index()
                
                pivot_nona['_CLEAN_NOPOL'] = pivot_nona['NOPOL'].astype(str).str.replace(' ', '').str.upper()
                
                # Mengambil Early Start dari df_servis
                early_start = df_servis.groupby('_CLEAN_NOPOL')['TGL_MASUK'].min().reset_index()
                early_start.rename(columns={'TGL_MASUK': 'EARLY_START'}, inplace=True)
                
                pivot_nona = pd.merge(pivot_nona, early_start, on='_CLEAN_NOPOL', how='left')
                pivot_nona['EARLY_START'] = pivot_nona['EARLY_START'].dt.strftime('%d/%m/%Y').fillna('-')
                
                pivot_nona['CUSTOM_ORDER'] = pivot_nona['_CLEAN_NOPOL'].map(lambda x: order_dict.get(x, 999999))
                pivot_nona = pivot_nona.sort_values(['CUSTOM_ORDER', 'NOPOL'])
                pivot_nona = pivot_nona.drop(columns=['CUSTOM_ORDER', '_CLEAN_NOPOL'])
                
                # Susun ulang kolom agar EARLY_START berada setelah identity_cols
                cols = list(pivot_nona.columns)
                cols.insert(len(identity_cols), cols.pop(cols.index('EARLY_START')))
                pivot_nona = pivot_nona[cols]
                
                pivot_nona = pivot_nona.set_index(identity_cols + ['EARLY_START'])
                pivot_nona = pivot_nona.reset_index()
                pivot_nona.rename(columns={'EARLY_START': 'TGL MASUK AWAL'}, inplace=True)
                
                # Set ulang kolom date jika ada yang jadi datetime mapping
                # Tidak perlu, date sudah dalam bentuk datetime yang nanti direname
                pivot_nona_dates = pivot_nona.columns[len(identity_cols)+1:]
                
                # Kalkulasi total & persentase
                for s in status_list:
                    pivot_nona[f'TOTAL {s}'] = (pivot_nona[pivot_nona_dates] == s).sum(axis=1)
                    pivot_nona[f'% {s}'] = (pivot_nona[f'TOTAL {s}'] / total_days).map(lambda x: f"{x:.2%}")
                
                # Ubah nama kolom date (datetime -> string)
                date_rename_map = {d: f"{d.day}/{d.month}" for d in date_range if d in pivot_nona.columns}
                pivot_nona.rename(columns=date_rename_map, inplace=True)
                
                pivot_nona.to_excel(writer, sheet_name='INVESTIGASI_NON_A', index=False)
            
        # ==========================================
        # 10. SHEET UNAVAILABLE_AKHIR (Tidak Available di Akhir Periode)
        # ==========================================
        last_date = date_range[-1]
        mask_unavail = (df_final['DATE'] == last_date) & (df_final['STATUS_CODE'] != 'A')
        unavail_nopols = df_final[mask_unavail]['NOPOL'].unique()
        
        df_unavail = df_final[df_final['NOPOL'].isin(unavail_nopols)]
        if not df_unavail.empty:
            pivot_unv = df_unavail.pivot_table(
                index=identity_cols,
                columns='DATE',
                values='STATUS_CODE',
                aggfunc='first'
            )
            pivot_unv = pivot_unv.reset_index()
            
            # Sort menggunakan urutan global dari custom order
            pivot_unv['_CLEAN_NOPOL'] = pivot_unv['NOPOL'].astype(str).str.replace(' ', '').str.upper()
            pivot_unv['CUSTOM_ORDER'] = pivot_unv['_CLEAN_NOPOL'].map(lambda x: order_dict.get(x, 999999))
            pivot_unv = pivot_unv.sort_values(['CUSTOM_ORDER', 'NOPOL'])
            pivot_unv = pivot_unv.drop(columns=['CUSTOM_ORDER', '_CLEAN_NOPOL'])
            
            pivot_unv = pivot_unv.set_index(identity_cols)
            pivot_unv = pivot_unv.reset_index()
            pivot_unv.columns = identity_cols + date_cols_str
            
            for s in status_list:
                pivot_unv[f'TOTAL {s}'] = (pivot_unv[date_cols_str] == s).sum(axis=1)
                pivot_unv[f'% {s}'] = (pivot_unv[f'TOTAL {s}'] / total_days).map(lambda x: f"{x:.2%}")
                
            pivot_unv.to_excel(writer, sheet_name='TDK_AVAILABLE_AKHIR', index=False)
            
    return output.getvalue()

if __name__ == '__main__':
    # Eksekusi testing lokal bila file dijalankan
    input_f = 'data/MEKANIK - STSFORM.csv' 
    master_f = 'data/KPI 95% - Available Car Report - 2026.xlsx'
    ts = '2026-01-01'
    te = '2026-01-31'
    
    print(f"Running ETL KPI Locally for {ts} to {te}...")
    excel_bytes = run_etl_kpi(ts, te, input_f, master_f)
    print("Writing bytes to file...")
    with open(f'output/Output_KPI_Report_{ts}_to_{te}_FINAL.xlsx', 'wb') as f:
        f.write(excel_bytes)
    print("Selesai.")
