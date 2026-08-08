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
            
    # Calculate more precise USIA based on TAHUN_PEMBUATAN
    current_year = datetime.now().year
    def calculate_precise_usia(row):
        try:
            thn = str(row['TAHUN_PEMBUATAN']).strip()
            import re
            nums = re.findall(r"20\d{2}|19\d{2}", thn)
            if nums:
                return max(0, current_year - int(nums[0]))
            
            usia_val = str(row['USIA']).replace(',', '.')
            nums2 = re.findall(r"[-+]?\d*\.\d+|\d+", usia_val)
            if nums2:
                return float(nums2[0])
            return row['USIA']
        except:
            return row['USIA']
            
    df_master['USIA'] = df_master.apply(calculate_precise_usia, axis=1)
    
    # Mapping USIA from Master Usia.xlsx for missing values
    try:
        import os
        import re
        master_usia_path = 'Master Usia.xlsx'
        if not os.path.exists(master_usia_path):
            master_usia_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Master Usia.xlsx')
            
        if os.path.exists(master_usia_path):
            df_mu = pd.read_excel(master_usia_path, sheet_name='Sheet10')
            df_mu['NOPOL'] = df_mu['NOPOL'].astype(str).str.strip().str.upper()
            dict_usia = {}
            for _, r in df_mu.iterrows():
                try:
                    val = str(r['USIA']).replace(',', '.')
                    nums_mu = re.findall(r"[-+]?\d*\.\d+|\d+", val)
                    if nums_mu:
                        nopol_clean = str(r['NOPOL']).replace(' ', '')
                        dict_usia[nopol_clean] = float(nums_mu[0])
                except:
                    pass
            
            def final_usia_mapping(row):
                val = row['USIA']
                nopol_clean = str(row['NOPOL']).replace(' ', '')
                if pd.isna(val) or val == '-' or val == 'Tidak Diketahui' or str(val).strip() == '':
                    return dict_usia.get(nopol_clean, val)
                return val
            df_master['USIA'] = df_master.apply(final_usia_mapping, axis=1)
    except Exception as e:
        pass
    
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
        
        # Build global order_dict for consistent sorting
        all_orders = []
        for lst in CUSTOM_NOPOL_ORDER.values():
            all_orders.extend(lst)
        global_order_dict = {nopol.replace(' ', '').upper(): i for i, nopol in enumerate(all_orders)}
        
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
            
            pivot_inv['CUSTOM_ORDER'] = pivot_inv['_CLEAN_NOPOL'].map(lambda x: global_order_dict.get(x, 999999))
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
                
                pivot_nona['CUSTOM_ORDER'] = pivot_nona['_CLEAN_NOPOL'].map(lambda x: global_order_dict.get(x, 999999))
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
            pivot_unv['CUSTOM_ORDER'] = pivot_unv['_CLEAN_NOPOL'].map(lambda x: global_order_dict.get(x, 999999))
            pivot_unv = pivot_unv.sort_values(['CUSTOM_ORDER', 'NOPOL'])
            pivot_unv = pivot_unv.drop(columns=['CUSTOM_ORDER', '_CLEAN_NOPOL'])
            
            pivot_unv = pivot_unv.set_index(identity_cols)
            pivot_unv = pivot_unv.reset_index()
            pivot_unv.columns = identity_cols + date_cols_str
            
            for s in status_list:
                pivot_unv[f'TOTAL {s}'] = (pivot_unv[date_cols_str] == s).sum(axis=1)
                pivot_unv[f'% {s}'] = (pivot_unv[f'TOTAL {s}'] / total_days).map(lambda x: f"{x:.2%}")
                
            pivot_unv.to_excel(writer, sheet_name='TDK_AVAILABLE_AKHIR', index=False)
            
        # ==========================================
        # 11. SHEET SUMMARY ALL NOPOL
        # ==========================================
        df_all_nopol = df_final.drop_duplicates(subset=['NOPOL', 'DATE'])
        
        if not df_all_nopol.empty:
            pivot_all = df_all_nopol.pivot_table(
                index=identity_cols,
                columns='DATE',
                values='STATUS_CODE',
                aggfunc='first'
            )
            pivot_all = pivot_all.reset_index()
            
            pivot_all['_CLEAN_NOPOL'] = pivot_all['NOPOL'].astype(str).str.replace(' ', '').str.upper()
            pivot_all['CUSTOM_ORDER'] = pivot_all['_CLEAN_NOPOL'].map(lambda x: global_order_dict.get(x, 999999))
            pivot_all = pivot_all.sort_values(['CUSTOM_ORDER', 'NOPOL'])
            pivot_all = pivot_all.drop(columns=['CUSTOM_ORDER', '_CLEAN_NOPOL'])
            
            pivot_all = pivot_all.set_index(identity_cols)
            pivot_all = pivot_all.reset_index()
            pivot_all.columns = identity_cols + date_cols_str
            
            for s in status_list:
                pivot_all[f'TOTAL {s}'] = (pivot_all[date_cols_str] == s).sum(axis=1)
                pivot_all[f'% {s}'] = (pivot_all[f'TOTAL {s}'] / total_days).map(lambda x: f"{x:.2%}")
                
            pivot_all.to_excel(writer, sheet_name='SUMMARY_ALL_NOPOL', index=False)
            
        # ==========================================
        # 12. SHEET SUMMARY USIA
        # ==========================================
        import re
        def get_kategori_usia(val):
            try:
                val_str = str(val).replace(',', '.')
                nums = re.findall(r"[-+]?\d*\.\d+|\d+", val_str)
                if nums:
                    u = float(nums[0])
                    if u <= 5: return '0-5 Tahun'
                    elif u <= 10: return '6-10 Tahun'
                    elif u <= 15: return '11-15 Tahun'
                    else: return '16 Tahun Lebih'
                return 'Tidak Diketahui'
            except:
                return 'Tidak Diketahui'
                
        df_usia = df_final.drop_duplicates(subset=['NOPOL', 'DATE']).copy()
        
        try:
            import os, re
            mu_path = 'Master Usia.xlsx'
            if not os.path.exists(mu_path):
                mu_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Master Usia.xlsx')
            if os.path.exists(mu_path):
                df_mu_tmp = pd.read_excel(mu_path, sheet_name='Sheet10')
                dict_mu = {}
                for _, r in df_mu_tmp.iterrows():
                    val = str(r['USIA']).replace(',', '.')
                    nums = re.findall(r"[-+]?\d*\.\d+|\d+", val)
                    if nums:
                        dict_mu[str(r['NOPOL']).replace(' ', '').upper()] = float(nums[0])
                
                def fill_usia(row):
                    val = row['USIA']
                    if pd.isna(val) or val == '-' or val == 'Tidak Diketahui' or str(val).strip() == '':
                        return dict_mu.get(str(row['NOPOL']).replace(' ', '').upper(), val)
                    return val
                df_usia['USIA'] = df_usia.apply(fill_usia, axis=1)
        except Exception:
            pass
            
        df_usia['KATEGORI_USIA'] = df_usia['USIA'].apply(get_kategori_usia)
        
        def get_jenis_kendaraan(val):
            return 'MOTOR' if str(val).strip().upper() == 'MOTOR' else 'MOBIL'
            
        df_usia['JENIS_KENDARAAN'] = df_usia['JENIS_MOBIL'].apply(get_jenis_kendaraan)
        
        total_cars_df = df_usia.groupby(['JENIS_KENDARAAN', 'KATEGORI_USIA'])['NOPOL'].nunique().reset_index(name='Total Kendaraan')
        
        df_usia_work = df_usia[df_usia['DATE'].dt.dayofweek != 6]
        
        if not df_usia_work.empty:
            status_counts = df_usia_work.groupby(['JENIS_KENDARAAN', 'KATEGORI_USIA', 'STATUS_CODE']).size().unstack(fill_value=0)
        else:
            status_counts = pd.DataFrame(columns=status_list)
            
        for s in status_list:
            if s not in status_counts.columns:
                status_counts[s] = 0
                
        summary_usia = pd.merge(total_cars_df, status_counts.reset_index(), on=['JENIS_KENDARAAN', 'KATEGORI_USIA'], how='left')
        summary_usia.fillna(0, inplace=True)
        
        cat_order = {'0-5 Tahun': 1, '6-10 Tahun': 2, '11-15 Tahun': 3, '16 Tahun Lebih': 4, 'Tidak Diketahui': 5}
        summary_usia['Order'] = summary_usia['KATEGORI_USIA'].map(cat_order)
        summary_usia = summary_usia.sort_values(['JENIS_KENDARAAN', 'Order']).drop(columns=['Order'])
        
        summary_usia['Total Hari Kerja'] = summary_usia['Total Kendaraan'] * total_work_days
        
        out_cols = ['JENIS_KENDARAAN', 'KATEGORI_USIA', 'Total Kendaraan', 'Total Hari Kerja']
        for s in status_list:
            t_col = f'TOTAL {s}'
            p_col = f'% {s}'
            summary_usia[t_col] = summary_usia[s]
            summary_usia[p_col] = summary_usia.apply(
                lambda row: f"{(row[t_col] / row['Total Hari Kerja']):.2%}" if row['Total Hari Kerja'] > 0 else "0.00%", axis=1
            )
            out_cols.extend([t_col, p_col])
            
        summary_usia = summary_usia[out_cols]
        summary_usia.to_excel(writer, sheet_name='SUMMARY_USIA', index=False)
            
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
