import streamlit as st
import pandas as pd
from thefuzz import fuzz
from datetime import datetime, timedelta
from io import BytesIO

st.title('Tracking Pengeluaran BBM per Jabatan')

st.write('''
Buatlah data baru berisikan kolom:
- VOUCHER NO.
- TRANS. DATE
- ENTRY DATE
- DESCRIPTION
- DEBIT

Catatan penting:
- Periksa header file untuk menghapus karakter spesial
- Simpan file dengan nama BBM.xlsx
- Ubah kolom debit menjadi numerik
- Pilih tanggal awal minggu untuk perhitungan
''')

def is_similar(text, keywords, threshold=85):
    """Check string similarity with fuzzy matching"""
    text = str(text).lower()
    
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
    """Kategorisasi deskripsi dengan prioritas yang tepat"""
    description = str(description).lower()
    
    # 1. Cek custom keywords
    for category, keywords in custom_keywords.items():
        if keywords and is_similar(description, keywords, threshold=85):
            return category
    
    # 2. Kategori berdasarkan jabatan
    categories = [
        ('ASMEN', ['asisten', 'assistant', 'asmen', 'assisten']),
        ('MIS', ['mis', 'msa']),
        ('STAF LAPANG', ['staf', 'staf lapang', 'staff lapang', 'staf lapangan', 'staff', 'orang', 'minggu', 'mingguan']),
        ('ADMIN', ['admin', 'administrasi', 'fsa', 'adm']),
        ('MANAGER', ['manager', 'manajer', 'branch manager', 'kepala cabang', 'mc', 'bm']),
        ('LAINYA', ['genset', 'jenset', 'dan'])
    ]
    
    for category, keywords in categories:
        if is_similar(description, keywords, threshold=85):
            return category
    
    return 'LAINYA'

def detect_date_anomalies(df):
    """Deteksi anomali antara tanggal entri dan tanggal transaksi"""
    df['ENTRY DATE'] = pd.to_datetime(df['ENTRY DATE'])
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'])
    
    df['DAYS_DIFFERENCE'] = (df['ENTRY DATE'] - df['TRANS. DATE']).dt.days
    
    anomalies = df[abs(df['DAYS_DIFFERENCE']) > 1].copy()
    
    if not anomalies.empty:
        anomalies['ENTRY DATE'] = anomalies['ENTRY DATE'].dt.strftime('%d/%m/%Y')
        anomalies['TRANS. DATE'] = anomalies['TRANS. DATE'].dt.strftime('%d/%m/%Y')
        
        return anomalies[['TRANS. DATE', 'ENTRY DATE', 'DAYS_DIFFERENCE', 'DESCRIPTION', 'DEBIT']]
    
    return None

def process_transactions(df, start_date, custom_keywords):
    """Proses transaksi dan hitung per minggu"""
    start_date = datetime.strptime(start_date, '%d/%m/%Y')
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'])
    
    df['CATEGORY'] = df['DESCRIPTION'].apply(lambda description: categorize_description(description, custom_keywords))
    
    max_date = df['TRANS. DATE'].max()
    weekly_ranges = _create_weekly_ranges(start_date, max_date)
    
    results = []
    for week_start, week_end in weekly_ranges:
        week_mask = (df['TRANS. DATE'].dt.date >= week_start.date()) & (df['TRANS. DATE'].dt.date <= week_end.date())
        week_data = df[week_mask]
        
        if not week_data.empty:
            pivot_data = week_data.pivot_table(
                values='DEBIT',
                columns='CATEGORY',
                aggfunc='sum',
                fill_value=0
            )
            
            week_row = {
                'Tanggal ( Per minggu )': f"{week_start.strftime('%d/%m/%Y')} - {week_end.strftime('%d/%m/%Y')}"
            }
            
            for col in ['ADMIN', 'ASMEN', 'LAINYA', 'MANAGER', 'MIS', 'STAF LAPANG']:
                week_row[col] = pivot_data.get(col, [0])[0]
                
            results.append(week_row)
    
    return pd.DataFrame(results)

def _create_weekly_ranges(start_date, end_date):
    """Create list of weekly date ranges"""
    current = start_date
    ranges = []
    while current <= end_date:
        week_end = current + timedelta(days=6)
        ranges.append((current, min(week_end, end_date)))
        current = week_end + timedelta(days=1)
    return ranges

def to_excel(categorized_df, weekly_results_df, date_anomalies_df):
    """Export data to Excel with formatting"""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        categorized_df.to_excel(writer, sheet_name='Kategorisasi', index=False)
        weekly_results_df.to_excel(writer, sheet_name='Perhitungan Mingguan', index=False)
        
        if date_anomalies_df is not None and not date_anomalies_df.empty:
            date_anomalies_df.to_excel(writer, sheet_name='Anomali Tanggal', index=False)
        
        workbook = writer.book
        header_format = workbook.add_format({'bold': True, 'align': 'center', 'bg_color': '#D3D3D3'})
        number_format = workbook.add_format({'num_format': '#,##0'})
        
        sheets = {
            'Kategorisasi': categorized_df,
            'Perhitungan Mingguan': weekly_results_df,
            'Anomali Tanggal': date_anomalies_df
        }
        
        for sheet_name, df in sheets.items():
            if df is not None and not df.empty:
                worksheet = writer.sheets[sheet_name]
                
                for col_num, value in enumerate(df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                for col_num in range(1, len(df.columns)):
                    worksheet.set_column(col_num, col_num, None, number_format)
                
                for i, col in enumerate(df.columns):
                    column_width = max(df[col].astype(str).map(len).max(), len(col))
                    worksheet.set_column(i, i, column_width + 2)
    
    output.seek(0)
    return output

# Custom keywords input form
with st.form("jabatan_form"):
    st.write("Masukkan nama-nama untuk masing-masing jabatan (opsional):")
    manager_names = st.text_area("Nama Manager (pisahkan dengan koma)", "").split(',')
    asmen_names = st.text_area("Nama Asmen (pisahkan dengan koma)", "").split(',')
    admin_names = st.text_area("Nama Admin/FSA (pisahkan dengan koma)", "").split(',')
    mis_names = st.text_area("Nama MIS/MSA (pisahkan dengan koma)", "").split(',')
    lainya_names = st.text_area("Nama lainya (pisahkan dengan koma)", "").split(',')
    submitted = st.form_submit_button("Simpan")

# Initialize custom keywords
if 'custom_keywords' not in st.session_state:
    st.session_state.custom_keywords = {
        'MANAGER': [], 'ASMEN': [], 'ADMIN': [], 'MIS': [], 'LAINYA': []
    }

# Form submission handling
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

# Main processing
if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        
        if st.button('Proses Analisa BBM'):
            # Kategorisasi
            df['CATEGORY'] = df['DESCRIPTION'].apply(lambda description: categorize_description(description, st.session_state.custom_keywords))
            categorized_df = df[['TRANS. DATE', 'DESCRIPTION', 'CATEGORY', 'DEBIT']].copy()
            categorized_df['TRANS. DATE'] = categorized_df['TRANS. DATE'].dt.strftime('%d/%m/%Y')
            
            # Perhitungan mingguan
            results_df = process_transactions(df, start_date, st.session_state.custom_keywords)
            
            # Anomali tanggal
            date_anomalies = detect_date_anomalies(df)
            
            st.write("Kategorisasi:")
            st.dataframe(categorized_df)
            
            st.write("Perhitungan Mingguan:")
            st.write(results_df)
            
            if date_anomalies is not None:
                st.write("Anomali Tanggal:")
                st.dataframe(date_anomalies)
            
            # Excel download
            excel_file = to_excel(categorized_df, results_df, date_anomalies)
            st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name=f'tracking_bbm_{start_date.replace("/", "-")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
