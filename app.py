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
    
def is_similar(text, keywords, threshold=80):
    """
    Helper function untuk mengecek kemiripan string menggunakan fuzzy matching
    threshold: nilai minimum kemiripan (0-100)
    """
    text = text.lower()
    
    # Direct match first (lebih cepat)
    for keyword in keywords:
        if keyword.lower() in text:
            return True
    
    # Fuzzy matching untuk menangani typo
    for keyword in keywords:
        # Token set ratio - untuk menangani urutan kata dan kata tambahan
        token_ratio = fuzz.token_set_ratio(text, keyword)
        if token_ratio >= threshold:
            return True
            
        # Partial ratio - untuk substring matching yang lebih akurat
        partial_ratio = fuzz.partial_ratio(text, keyword)
        if partial_ratio >= threshold:
            return True
    
    return False

def categorize_description(description):
    """Mengkategorikan description ke dalam jabatan dengan fuzzy matching untuk menangani typo"""
    description = str(description).lower()
    
    # Dictionary untuk kategori
    categories = {
        'MANAGER': [
            'manager',
            'manajer',
            'branch manager',
            'kepala cabang',
            'mc',
            'bm'
        ],
        'ASMEN': [
            'asisten',
            'assistant',
            'asmen'
        ],
        'MIS': [
            'mis',
            'msa',
            'management information system'
        ],
        'ADMIN': [
            'admin',
            'administrasi',
            'fsa',
            'administration'
        ],
        'STAF LAPANG': [
            'staf 16',
            'staff 16',
            'mingguan staf',
            'staf lapang',
            'staff lapang',
            'staf lapangan',
            'staff'
        ]
    }
    
    # Cek kategori berdasarkan prioritas
    # 1. Cek MANAGER dulu karena ini sering muncul sebagai bagian dari kata lain
    if is_similar(description, categories['MANAGER'], threshold=90):
        return 'MANAGER'
        
    # 2. Cek ASMEN dengan threshold yang lebih tinggi
    if is_similar(description, categories['ASMEN'], threshold=85):
        return 'ASMEN'
    
    # 3. Cek kategori lainnya dengan threshold normal
    for category, keywords in categories.items():
        if category not in ['MANAGER', 'ASMEN']:  # Skip yang sudah dicek
            if is_similar(description, keywords):
                return category
    
    return 'LAINYA'

def process_transactions(df, start_date):
    # Convert start_date to datetime
    start_date = datetime.strptime(start_date, '%d/%m/%Y')
    
    # Convert date column to datetime
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE'])
    
    # Add category column
    df['CATEGORY'] = df['DESCRIPTION'].apply(categorize_description)
    
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
        
        if st.button('Proses Data'):
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
