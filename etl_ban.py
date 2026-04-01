import pandas as pd
import numpy as np
import io

def run_etl_ban(input_file):
    """
    Membaca data BAN SHARE, menggunakan baris 7 sebagai header (skiprows=6),
    mengekstrak Nopol, Serial Number, KM Pemasangan, dan KM Pengajuan.
    Menghitung Jarak Tempuh (KM Pengajuan - KM Pemasangan).
    """
    # Membaca excel, skiprows=6 karena header di baris 7 (indeks 6)
    if isinstance(input_file, bytes):
        input_file = io.BytesIO(input_file)
    df_raw = pd.read_excel(input_file, skiprows=6, dtype=str)
    
    # Normalisasi nama kolom untuk memudahkan pencarian fleksibel
    columns_mapped = {c: str(c).strip().upper() for c in df_raw.columns}
    df_raw.rename(columns=columns_mapped, inplace=True)
    
    col_nopol = None
    col_serial = None
    col_km_pasang = None
    col_km_pengajuan = None

    # Pencarian kolom fleksibel
    for c in df_raw.columns:
        c_str = str(c)
        if 'NOPOL' in c_str or 'NO POL' in c_str or 'POLISI' in c_str:
            if not col_nopol: col_nopol = c
        elif 'SERIAL' in c_str or 'SERI' in c_str or 'S/N' in c_str:
            if not col_serial: col_serial = c
        elif 'PASANG' in c_str or 'AWAL' in c_str or 'PEMASANGAN' in c_str:
            if not col_km_pasang: col_km_pasang = c
        elif 'PENGAJUAN' in c_str or 'BONGKAR' in c_str or 'AKHIR' in c_str:
            if not col_km_pengajuan: col_km_pengajuan = c

    # Fallback ke index kolom jika benar-benar tidak ditemukan (sangat jarang jika format jelas)
    if not col_nopol:
        raise ValueError("Gagal menemukan kolom NOPOL di file Ban Share.")
    if not col_serial:
        raise ValueError("Gagal menemukan kolom Serial Number di file Ban Share.")
    if not col_km_pasang:
        raise ValueError("Gagal menemukan kolom KM Pemasangan di file Ban Share.")
    if not col_km_pengajuan:
        raise ValueError("Gagal menemukan kolom KM Pengajuan di file Ban Share.")

    # Ambil kolom yang dibutuhkan dan hapus duplikat / spasi
    df = df_raw[[col_nopol, col_serial, col_km_pasang, col_km_pengajuan]].copy()
    df.rename(columns={
        col_nopol: 'NOPOL',
        col_serial: 'SERIAL_NUMBER',
        col_km_pasang: 'KM_PEMASANGAN',
        col_km_pengajuan: 'KM_PENGAJUAN'
    }, inplace=True)

    # Bersihkan Data
    df = df.dropna(subset=['NOPOL', 'SERIAL_NUMBER'])
    df['NOPOL'] = df['NOPOL'].astype(str).str.replace(' ', '').str.upper()
    df['SERIAL_NUMBER'] = df['SERIAL_NUMBER'].astype(str).str.strip().str.upper()

    # Ekstrak angka saja dari KM
    df['KM_PEMASANGAN'] = df['KM_PEMASANGAN'].astype(str).str.replace(r'[^\d.]', '', regex=True)
    df['KM_PENGAJUAN']  = df['KM_PENGAJUAN'].astype(str).str.replace(r'[^\d.]', '', regex=True)

    df['KM_PEMASANGAN'] = pd.to_numeric(df['KM_PEMASANGAN'], errors='coerce').fillna(0)
    df['KM_PENGAJUAN']  = pd.to_numeric(df['KM_PENGAJUAN'], errors='coerce').fillna(0)

    # Hitung Jarak Tempuh
    df['JARAK_TEMPUH'] = df['KM_PENGAJUAN'] - df['KM_PEMASANGAN']

    # Urutkan berdasarkan Nopol dan Serial Number
    df = df.sort_values(by=['NOPOL', 'SERIAL_NUMBER', 'JARAK_TEMPUH'], ascending=[True, True, False])
    
    # Export ke format excel di memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Analisis_Ban', index=False)
        
        # Format lebar kolom
        worksheet = writer.sheets['Analisis_Ban']
        worksheet.set_column('A:A', 15)  # Nopol
        worksheet.set_column('B:B', 30)  # Serial Number
        worksheet.set_column('C:D', 20)  # KM
        worksheet.set_column('E:E', 20)  # Jarak Tempuh
        
    return output.getvalue(), df
