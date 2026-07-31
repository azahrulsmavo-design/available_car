import pandas as pd
import io

def run_etl_sparepart(file_bytes, filename, start_date, end_date):
    # 1. Load File
    if filename.lower().endswith('.csv'):
        df = pd.read_csv(io.BytesIO(file_bytes))
    else:
        df = pd.read_excel(io.BytesIO(file_bytes))

    # Clean headers
    df.columns = df.columns.str.strip().str.lower()

    # Pastikan kolom yang dibutuhkan ada
    required_cols = ['tgl_create_rf', 'status_approval_rf', 'kategori', 'nopol', 'jenis_barang']
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Kolom {col} tidak ditemukan dalam file.")

    # Convert tgl_create_rf ke datetime
    df['tgl_create_rf'] = pd.to_datetime(df['tgl_create_rf'], errors='coerce')
    
    # Date filter
    start_dt = pd.to_datetime(start_date)
    # End date include up to 23:59:59
    end_dt = pd.to_datetime(end_date) + pd.Timedelta(days=1, microseconds=-1)
    
    df = df[(df['tgl_create_rf'] >= start_dt) & (df['tgl_create_rf'] <= end_dt)]

    # 2. Filters
    # status_approval_rf = Approved By Head Dept
    df = df[df['status_approval_rf'].astype(str).str.strip().str.lower() == 'approved by head dept']

    # kategori = SPAREPART
    df = df[df['kategori'].astype(str).str.strip().str.lower() == 'sparepart']

    # nopol tidak kosong
    df = df[df['nopol'].notna() & (df['nopol'].astype(str).str.strip() != '') & (df['nopol'].astype(str).str.strip().str.lower() != 'nan')]

    # 4. Aggregate
    # Sheet 1: Total per barang
    df_barang = df.groupby(['jenis_barang']).agg(
        Jumlah_Pemakaian=('jenis_barang', 'count')
    ).reset_index()
    df_barang = df_barang.sort_values(by=['Jumlah_Pemakaian'], ascending=False)

    # Sheet 2: Total per nopol (Hitung SKU sparepart yang dipakai)
    df_nopol = df.groupby(['nopol']).agg(
        Jumlah_Jenis_Sparepart_Terpakai=('jenis_barang', 'nunique')
    ).reset_index()
    df_nopol = df_nopol.sort_values(by=['Jumlah_Jenis_Sparepart_Terpakai'], ascending=False)

    # 5. Export to Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_barang.to_excel(writer, index=False, sheet_name='Total Per Barang')
        df_nopol.to_excel(writer, index=False, sheet_name='Total Per Nopol')
        
        # format workbook
        workbook = writer.book
        
        worksheet1 = writer.sheets['Total Per Barang']
        worksheet1.set_column('A:A', 35)
        worksheet1.set_column('B:B', 20)
        
        worksheet2 = writer.sheets['Total Per Nopol']
        worksheet2.set_column('A:A', 20)
        worksheet2.set_column('B:B', 35)
        
    processed_bytes = output.getvalue()

    return processed_bytes, df_barang
