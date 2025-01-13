import pandas as pd
import streamlit as st
from datetime import datetime, timedelta

# Fungsi untuk mengelompokkan kategori berdasarkan description
def categorize(description):
    desc = str(description).lower()
    if any(keyword in desc for keyword in ['manajer cabang', 'manager cabang', 'mc', 'manajer', 'manager', 'bm']):
        return 'Manager'
    elif any(keyword in desc for keyword in ['asisten manajer cabang', 'asmen', 'asisten manajer', 'asisten manager cabang', 'asisten manager']):
        return 'Asmen'
    elif any(keyword in desc for keyword in ['admin 1', 'staf admin', 'staf administrasi', 'staff admin', 'staff administrasi', 'fsa', 'admin 2']):
        return 'Admin/FSA'
    elif any(keyword in desc for keyword in ['mis', 'staf mis', 'staff mis', 'msa']):
        return 'MIS/MSA'
    else:
        return 'Lainnya'

# Upload file
uploaded_file = st.file_uploader("Upload file Excel", type=["xlsx"])

if uploaded_file:
    # Baca file Excel
    df = pd.read_excel(uploaded_file)
    
    # Bersihkan spasi di nama kolom
    df.columns = df.columns.str.strip()
    
    # Tampilkan nama kolom untuk verifikasi
    st.write("Nama Kolom:", df.columns.tolist())
    
    # Ubah kolom tanggal menjadi datetime jika kolom tersedia
    if 'TRANS. DATE' in df.columns:
        df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'], format='%d/%m/%Y', errors='coerce')
    else:
        st.error("Kolom 'TRANS. DATE' tidak ditemukan di file.")
    
    # Tambahkan kolom kategori
    df['Kategori'] = df['DESCRIPTION'].apply(categorize)
    
    # Input tanggal awal minggu dari user
    start_date = st.date_input("Pilih tanggal awal minggu (Senin):", value=datetime.today())
    start_date = pd.to_datetime(start_date)
    
    # Hitung minggu berdasarkan input user
    df['Minggu'] = df['TRANS. DATE'].apply(lambda x: (start_date + timedelta(days=((x - start_date).days // 7 * 7))).strftime('%d/%m/%Y') if pd.notnull(x) else 'Tanggal Tidak Valid')
    
    # Rekap pengeluaran per kategori dan per minggu
    summary = df.groupby(['Minggu', 'Kategori'])['DEBIT'].sum().unstack(fill_value=0).reset_index()
    
    # Pastikan semua kategori tampil
    for kategori in ['Manager', 'Asmen', 'Admin/FSA', 'MIS/MSA']:
        if kategori not in summary.columns:
            summary[kategori] = 0
    
    # Urutkan kolom sesuai permintaan
    summary = summary[['Minggu', 'Manager', 'Asmen', 'Admin/FSA', 'MIS/MSA']]
    
    # Tampilkan hasil
    st.dataframe(summary)
    
    # Download hasil ke Excel
    output = pd.ExcelWriter("Rekap_Pengeluaran.xlsx", engine='xlsxwriter')
    summary.to_excel(output, index=False, sheet_name='Rekap')
    output.close()
    st.download_button(
        label="Download Rekap Excel",
        data=open("Rekap_Pengeluaran.xlsx", "rb").read(),
        file_name="Rekap_Pengeluaran.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
