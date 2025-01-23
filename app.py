import streamlit as st
import pandas as pd
from thefuzz import fuzz, process
from datetime import datetime, timedelta
from io import BytesIO

# Streamlit interface
st.title('Tracking Pengeluaran BBM per Jabatan')
st.write('''Buatlah data baru berisikan kolom | VOUCHER NO. | TRANS. DATE | ENTRY DATE | DESCRIPTION | DEBIT | jika anda copypaste dari data yang didownload di mdis ijo, check terlebih dahulu bagian headernya karena pasti ada karakter spesial, jika ada hapus terlebih dahulu karakter spesial tersebut.
''')
st.write('''Untuk penamaan file jadi BBM.xlsx, untuk kolom debit di ubah ke Numerik bukan Accounting! Karena nilai nol akan terbaca tanda "-" bukan angkan nol "0".''')
st.write('''Input tanggal awal pengecekkan hari senin, misal pengecekkan dari Januari 2025 s.d Desember 2025, maka pilih tanggal awalnya hari senin di minggu itu, jadi diinput tanggal 30 Desember 2024 (karena tanggal 1 hari rabu, dan tanggal 30 Hari Senin di minggu itu.''')
st.write('''Ada beberapa cabang yang tidak menambahkan deskripsi jabatan, mohon di cek terlebih dahulu sebelum di eksekusi dengan tools ini dan diisi manual untuk jabatan nya. Misalnya | Dibayar BBM untuk transport (irfan) |, setelah kata transport tambahkan manual jabatan dengan nama tersebut, sehingga menjadi | Dibayar BBM untuk transport asmen (irfan) |''')
st.write('''Untuk lainya input kata "dan"''')

# Form input untuk nama-nama jabatan
with st.form("jabatan_form"):
    st.write("Masukkan nama-nama untuk masing-masing jabatan (opsional):")
    manager_names = st.text_area("Nama Manager (pisahkan dengan koma)", "").split(',')
    asmen_names = st.text_area("Nama Asmen (pisahkan dengan koma)", "").split(',')
    admin_names = st.text_area("Nama Admin/FSA (pisahkan dengan koma)", "").split(',')
    mis_names = st.text_area("Nama MIS/MSA (pisahkan dengan koma)", "").split(',')
    lainya_names = st.text_area("Nama lainya (pisahkan dengan koma)", "").split(',')
    submitted = st.form_submit_button("Simpan")

# Fungsi-fungsi utama
def is_similar(text, keywords, threshold=85):
    """Helper function untuk mengecek kemiripan string"""
    text = text.lower()
    
    # Direct match first
    if any(keyword.lower() in text for keyword in keywords):
        return True
    
    # Fuzzy matching
    for keyword in keywords:
        keyword = keyword.lower()
        if any(ratio >= threshold for ratio in [
            fuzz.ratio(text, keyword),
            fuzz.partial_ratio(text, keyword),
            fuzz.token_sort_ratio(text, keyword)
        ]):
            return True
    return False

def categorize_description(description, custom_keywords):
    """Kategorisasi dengan prioritas yang tepat"""
    description = str(description).lower()
    
    # 1. Cek custom keywords (nama-nama yang diinput) terlebih dahulu
    for category, keywords in custom_keywords.items():
        if keywords and is_similar(description, keywords, threshold=85):
            return category
    
    # 2. Cek kategori berdasarkan jabatan
    if is_similar(description, ['asisten', 'assistant', 'asmen', 'assisten'], threshold=90):
        return 'ASMEN'
    
    if is_similar(description, ['mis', 'msa'], threshold=85):
        return 'MIS'
    
    if is_similar(description, ['staf', 'staf lapang', 'staff lapang', 'staf lapangan', 'staff', 'orang', 'minggu', 'mingguan'], threshold=85):
        return 'STAF LAPANG'
    
    if is_similar(description, ['admin', 'administrasi', 'fsa', 'adm'], threshold=90):
        return 'ADMIN'
    
    if is_similar(description, ['manager', 'manajer', 'branch manager', 'kepala cabang', 'mc', 'bm'], threshold=85):
        return 'MANAGER'
    
    if is_similar(description, ['genset', 'jenset', 'dan'], threshold=90):
        return 'LAINYA'
    
    return 'LAINYA'

def create_weekly_ranges(start_date, end_date):
    """Create list of weekly date ranges"""
    current = start_date
    ranges = []
    while current <= end_date:
        week_end = current + timedelta(days=6)
        ranges.append((current, min(week_end, end_date)))
        current = week_end + timedelta(days=1)
    return ranges

def detect_date_anomalies(df):
    """
    Deteksi anomali antara tanggal entri dan tanggal transaksi
    """
    # Konversi kolom tanggal ke datetime
    df['ENTRY DATE'] = pd.to_datetime(df['ENTRY DATE'])
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'])
    
    # Hitung selisih hari antara tanggal entri dan tanggal transaksi
    df['DAYS_DIFFERENCE'] = (df['ENTRY DATE'] - df['TRANS. DATE']).dt.days
    
    # Deteksi anomali - misalnya, perbedaan lebih dari 7 hari dianggap anomali
    anomalies = df[abs(df['DAYS_DIFFERENCE']) > 7].copy()
    
    if not anomalies.empty:
        # Format ulang tanggal untuk tampilan yang lebih baik
        anomalies['ENTRY DATE'] = anomalies['ENTRY DATE'].dt.strftime('%d/%m/%Y')
        anomalies['TRANS. DATE'] = anomalies['TRANS. DATE'].dt.strftime('%d/%m/%Y')
        
        return anomalies[['TRANS. DATE', 'ENTRY DATE', 'DAYS_DIFFERENCE', 'DESCRIPTION', 'DEBIT']]
    
    return None
    
def process_transactions(df, start_date, custom_keywords):
    # Convert start_date to datetime
    start_date = datetime.strptime(start_date, '%d/%m/%Y')
    
    # Convert date column to datetime
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'])
    
    # Tambahkan kembali bagian deteksi anomali tanggal
    date_anomalies = detect_date_anomalies(df)
    
    if date_anomalies is not None:
        st.warning("⚠️ Terdeteksi Anomali Tanggal:")
        st.dataframe(date_anomalies)
    
    # Add category column
    df['CATEGORY'] = df['DESCRIPTION'].apply(lambda description: categorize_description(description, custom_keywords))
    
def process_transactions(df, start_date, custom_keywords):
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

def to_excel(categorized_df, weekly_results_df, date_anomalies_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Sheet 1: Kategorisasi
        categorized_df.to_excel(writer, sheet_name='Kategorisasi', index=False)
        
        # Sheet 2: Hasil Perhitungan Mingguan
        weekly_results_df.to_excel(writer, sheet_name='Perhitungan Mingguan', index=False)
        
        # Sheet 3: Anomali Tanggal
        if date_anomalies_df is not None and not date_anomalies_df.empty:
            date_anomalies_df.to_excel(writer, sheet_name='Anomali Tanggal', index=False)
        
        # Format untuk setiap sheet
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True,
            'align': 'center',
            'bg_color': '#D3D3D3'
        })
        number_format = workbook.add_format({'num_format': '#,##0'})
        
        # Apply formatting to each sheet
        sheets = {
            'Kategorisasi': categorized_df,
            'Perhitungan Mingguan': weekly_results_df,
            'Anomali Tanggal': date_anomalies_df
        }
        
        for sheet_name, df in sheets.items():
            if df is not None and not df.empty:
                worksheet = writer.sheets[sheet_name]
                
                # Apply header format
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Apply number format to amount columns
                for col_num in range(1, len(df.columns)):
                    worksheet.set_column(col_num, col_num, None, number_format)
                
                # Adjust column width
                for i, col in enumerate(df.columns):
                    column_width = max(df[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(i, i, column_width + 2)
    
    output.seek(0)
    return output

# Variabel global untuk custom keywords
if 'custom_keywords' not in st.session_state:
    st.session_state.custom_keywords = {
        'MANAGER': [],
        'ASMEN': [],
        'ADMIN': [],
        'MIS': [],
        'LAINYA': []
    }

# Saat form di-submit
if submitted:
    st.session_state.custom_keywords = {
        'MANAGER': [name.strip().lower() for name in manager_names if name.strip()],
        'ASMEN': [name.strip().lower() for name in asmen_names if name.strip()],
        'ADMIN': [name.strip().lower() for name in admin_names if name.strip()],
        'MIS': [name.strip().lower() for name in mis_names if name.strip()],
        'LAINYA': [name.strip().lower() for name in lainya_names if name.strip()],
    }
    st.success("Daftar nama berhasil disimpan!")

# Date input
start_date = st.date_input(
    "Pilih tanggal awal (Senin):",
    datetime.now()
).strftime('%d/%m/%Y')

# File uploader
uploaded_file = st.file_uploader("Upload file Excel transaksi:", type=['xlsx', 'xls'])

# Modifikasi bagian utama
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        if st.button('Proses Analisa BBM'):
            # Simpan hasil kategorisasi
            df['CATEGORY'] = df['DESCRIPTION'].apply(lambda description: categorize_description(description, st.session_state.custom_keywords))
            categorized_df = df[['TRANS. DATE', 'DESCRIPTION', 'CATEGORY', 'DEBIT']].copy()
            categorized_df['TRANS. DATE'] = categorized_df['TRANS. DATE'].dt.strftime('%d/%m/%Y')
            
            # Proses transaksi
            results_df = process_transactions(df, start_date, st.session_state.custom_keywords)
            
            # Deteksi anomali tanggal
            date_anomalies = detect_date_anomalies(df)
            
            st.write("Hasil perhitungan:")
            st.write(results_df)
            
            # Create Excel download button dengan 3 sheet
            excel_file = to_excel(categorized_df, results_df, date_anomalies)
            st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name=f'tracking_bbm_{start_date.replace("/", "-")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
