'''
APLIKASI PYTHON UNTUK REKONSILIASI
AKUN HARUS MEMILIKI TRANS REFERENCEE YANG SAMA UNTUK TIAP TRANSAKSINYA
'''

import pandas as pd
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename

# Sembunyikan jendela Tkinter
Tk().withdraw()

# Buka file explorer
datapath = askopenfilename(
    title="Pilih file Excel untuk Rekonsiliasi",
    filetypes=[("Excel files", "*.xlsx *.xls")]
)

if datapath:
    # Validasi file exists
    if not os.path.exists(datapath):
        print(f"\n❌ ERROR: File tidak ditemukan: {datapath}")
        exit()
    
    # Validasi file is Excel
    if not (datapath.endswith('.xlsx') or datapath.endswith('.xls')):
        print(f"\n❌ ERROR: File harus berformat Excel (.xlsx atau .xls)")
        exit()
    
    print("File yang akan dicek:", datapath)

    # Baca file
    df = pd.read_excel(datapath)

    # Tampilkan daftar kolom untuk membantu user
    print("\nKolom yang tersedia:")
    for c in df.columns:
        print("-", c)

    # Input nama kolom dari user
    kolom_ket = input("\nMasukkan kolom Transaksi: ")
    kolom_debet = input("Masukkan kolom Debet: ")
    kolom_kredit = input("Masukkan kolom Kredit: ")

    # Validasi kolom
    for col in [kolom_ket, kolom_debet, kolom_kredit]:
        if col not in df.columns:
            print(f"\n❌ ERROR: Kolom '{col}' tidak ditemukan!")
            exit()

    # Konversi kolom debet dan kredit ke tipe numeric
    df[kolom_debet] = pd.to_numeric(df[kolom_debet], errors='coerce')
    df[kolom_kredit] = pd.to_numeric(df[kolom_kredit], errors='coerce')
    
    # Cek ada NaN (data yang tidak bisa dikonversi)
    if df[kolom_debet].isna().any() or df[kolom_kredit].isna().any():
        print(f"\n❌ ERROR: Ada nilai non-numeric di kolom debet atau kredit!")
        exit()

    # Proses rekonsiliasi
    rekap = df.groupby(kolom_ket)[[kolom_debet, kolom_kredit]].sum()
    rekap['Selisih'] = rekap[kolom_debet] - rekap[kolom_kredit]

    # Cari transaksi bermasalah
    keterangan_selisih = rekap[rekap['Selisih'] != 0].index.tolist()
    data_selisih = df[df[kolom_ket].isin(keterangan_selisih)]

    # Hitung total selisih
    total_selisih = rekap['Selisih'].sum()
    total_selisih_format = f"Rp {total_selisih:,.2f}"
    print(f"\n📊 Total Selisih: {total_selisih_format}")
    print(f"📋 Jumlah transaksi bermasalah: {len(keterangan_selisih)}")

    # Output file
    output_file = input("\nMasukkan nama file output (contoh: hasil_rekonsiliasi.xlsx): ")
    
    # Validasi output filename
    if not (output_file.endswith('.xlsx') or output_file.endswith('.xls')):
        output_file += '.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        data_selisih.to_excel(
            writer, sheet_name='Transaksi Bermasalah', index=False)
        rekap.loc[keterangan_selisih].to_excel(
            writer, sheet_name='Rekap Selisih')

    print(f"\nHasil rekonsiliasi disimpan ke: {output_file}")

else:
    print("Tidak ada file yang dipilih!")