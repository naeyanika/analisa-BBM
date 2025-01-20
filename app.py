import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from datetime import datetime, timedelta
from io import BytesIO


def create_weekly_ranges(start_date, end_date):
    """Membuat list range mingguan dari tanggal awal sampai akhir"""
    current = start_date
    weekly_ranges = []
    while current <= end_date:
        week_end = current + timedelta(days=6)
        weekly_ranges.append((current, week_end))
        current = current + timedelta(days=7)
    return weekly_ranges
    
def is_similar(text, keywords, threshold=95):  # Turunkan threshold ke 80
    """
    Helper function untuk mengecek kemiripan string menggunakan fuzzy matching
    threshold: nilai minimum kemiripan (0-100)
    """
    text = text.lower()
    
    # Direct match first (lebih cepat)
    if any(keyword.lower() in text for keyword in keywords):
        return True
    
    # Fuzzy matching untuk menangani typo
    for keyword in keywords:
        keyword = keyword.lower()
        # Ratio biasa - untuk typo umum
        ratio = fuzz.ratio(text, keyword)
        if ratio >= threshold:
            return True
            
        # Partial ratio - untuk substring matching
        partial_ratio = fuzz.partial_ratio(text, keyword)
        if partial_ratio >= threshold:
            return True
            
        # Token sort ratio - untuk kata yang urutannya berbeda
        token_ratio = fuzz.token_sort_ratio(text, keyword)
        if token_ratio >= threshold:
            return True
    
    return False

def categorize_description(description, custom_keywords):
    description = str(description).lower()
    
    # 1. Cek spesifik untuk "asmen" terlebih dahulu
    asmen_specific = ['asisten', 'assistant', 'asmen', 'assisten']
    for keyword in asmen_specific:
        if is_similar(description, [keyword], threshold=90):
            return 'ASMEN'
    
    # 2. Dictionary untuk kategori prioritas kedua (dengan threshold lebih tinggi)
    categories = {
        'MIS': ['mis', 'msa'],
        'ADMIN': ['admin', 'administrasi', 'fsa'],
        'STAF LAPANG': ['staf', 'staf lapang', 'staff lapang', 'staf lapangan', 'staff', 'orang'],
        'LAINYA': ['genset', 'jenset']
    }
    
    # Cek kategori prioritas kedua dengan threshold 90
    for category, keywords in categories.items():
        if is_similar(description, keywords, threshold=90):  # Naikkan threshold
            return category
    
    # 3. Cek custom keywords (nama-nama yang diinput)
    for category, keywords in custom_keywords.items():
        if keywords:
            if is_similar(description, keywords):
                return category
    
    # 4. Cek kategori MANAGER dengan threshold 85
    manager_keywords = ['manager', 'manajer', 'branch manager', 'kepala cabang', 'mc', 'bm']
    if is_similar(description, manager_keywords, threshold=100):
        return 'MANAGER'
    
    return 'LAINYA'

def process_transactions(df, start_date):
    # Convert start_date to datetime
    start_date = datetime.strptime(start_date, '%d/%m/%Y')
    
    # Convert date column to datetime
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'])
    
    # Add category column
    df['CATEGORY'] = df['DESCRIPTION'].apply(lambda description: categorize_description(description, custom_keywords))
    
    # Tampilkan hasil kategorisasi
    st.write("Detail Kategorisasi:")
    categorized_df = df[['TRANS. DATE', 'DESCRIPTION', 'CATEGORY', 'DEBIT']].copy()
    categorized_df['TRANS. DATE'] = categorized_df['TRANS. DATE'].dt.strftime('%d/%m/%Y')
    st.dataframe(categorized_df)
    
    # Get min and max dates from data
    min_date = df['TRANS. DATE'].min()
    max_date = df['TRANS. DATE'].max()
    
    # Create weekly ranges
    weekly_ranges = create_weekly_ranges(start_date, max_date)
    
    # Initialize results list
    results = []
    
    # Process each week
    for week_start, week_end in weekly_ranges:
        week_mask = (df['TRANS. DATE'].dt.date >= week_start.date()) & (df['TRANS. DATE'].dt.date <= week_end.date())
        week_data = df[week_mask]
        
        # Create pivot for the week
        if not week_data.empty:
            pivot_data = week_data.pivot_table(
                index=None,
                values='DEBIT',
                columns='CATEGORY',
                aggfunc='sum',
                fill_value=0
            )
            
            # Create row for the week
            week_row = {
                'Tanggal ( Per minggu )': f"{week_start.strftime('%d/%m/%Y')} - {week_end.strftime('%d/%m/%Y')}"
            }
            
            # Add columns in specific order
            for col in ['ADMIN', 'ASMEN', 'LAINYA', 'MANAGER', 'MIS', 'STAF LAPANG']:
                week_row[col] = pivot_data.get(col, [0])[0]
                
            results.append(week_row)
    
    return pd.DataFrame(results)

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Format settings
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'bg_color': '#D3D3D3'
        })
        
        # Number format
        number_format = workbook.add_format({'num_format': '#,##0'})
        
        # Apply header format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Apply number format to amount columns
        for col_num in range(1, len(df.columns)):  # Skip the date column
            worksheet.set_column(col_num, col_num, None, number_format)
            
        # Adjust column width
        for i, col in enumerate(df.columns):
            column_width = max(df[col].astype(str).map(len).max(), len(col))
            worksheet.set_column(i, i, column_width + 2)
    
    output.seek(0)
    return output

# Streamlit interface
st.title('Tracking Pengeluaran BBM per Jabatan')
st.write('''Buatlah data baru berisikan kolom | VOUCHER NO. | TRANS. DATE | DESCRIPTION | DEBIT | jika anda copypaste dari data yang didownload di mdis ijo, check terlebih dahulu bagian headernya karena pasti ada karakter spesial, jika ada hapus terlebih dahulu karakter spesial tersebut.
''')
st.write('''Untuk penamaan file jadi BBM.xlsx, untuk kolom debit di ubah ke Numerik bukan Accounting! Karena nilai nol akan terbaca tanda "-" bukan angkan nol "0".''')
st.write('''Input tanggal awal pengecekkan hari senin, misal pengecekkan dari Januari 2025 s.d Desember 2025, maka pilih tanggal awalnya hari senin di minggu itu, jadi diinput tanggal 30 Desember 2024 (karena tanggal 1 hari rabu, dan tanggal 30 Hari Senin di minggu itu.''')
st.write('''Ada beberapa cabang yang tidak menambahkan deskripsi jabatan, mohon di cek terlebih dahulu sebelum di eksekusi dengan tools ini dan diisi manual untuk jabatan nya. Misalnya | Dibayar BBM untuk transport (irfan) |, setelah kata transport tambahkan manual jabatan dengan nama tersebut, sehingga menjadi | Dibayar BBM untuk transport asmen (irfan) |''')

# Form input untuk nama-nama jabatan
with st.form("jabatan_form"):
    st.write("Masukkan nama-nama untuk masing-masing jabatan (opsional):")
    manager_names = st.text_area("Nama Manager (pisahkan dengan koma)", "").split(',')
    asmen_names = st.text_area("Nama Asmen (pisahkan dengan koma)", "").split(',')
    admin_names = st.text_area("Nama Admin/FSA (pisahkan dengan koma)", "").split(',')
    mis_names = st.text_area("Nama MIS/MSA (pisahkan dengan koma)", "").split(',')
    submitted = st.form_submit_button("Simpan")

if submitted:
    st.success("Daftar nama berhasil disimpan!")

custom_keywords = {
    'ASMEN': [name.strip().lower() for name in asmen_names],
    'ADMIN': [name.strip().lower() for name in admin_names],
    'MIS': [name.strip().lower() for name in mis_names],
    'MANAGER': [name.strip().lower() for name in manager_names]
}

# Date input
start_date = st.date_input(
    "Pilih tanggal awal (Senin):",
    datetime.now()
).strftime('%d/%m/%Y')

# File uploader for Excel
uploaded_file = st.file_uploader("Upload file Excel transaksi:", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.write("Preview data:")
        st.write(df.head())
        
        if st.button('Proses Analisa BBM'):
            results_df = process_transactions(df, start_date)
            
            st.write("Hasil perhitungan:")
            st.write(results_df)
            
            # Create Excel download button
            excel_file = to_excel(results_df)
            st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name=f'tracking_bbm_{start_date.replace("/", "-")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
