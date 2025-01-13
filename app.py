import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

def process_transactions(df, start_date):
    # Define job title mappings
    job_mappings = {
        'Manager': ['Manajer Cabang', 'Manager Cabang', 'MC', 'Manajer', 'Manager', 'BM'],
        'Asmen': ['Asisten Manajer Cabang', 'Asmen', 'Asisten Manajer', 'Asisten Manager Cabang', 'Asisten Manager'],
        'Admin/FSA': ['Admin 1', 'Staf Admin', 'Staf Administrasi', 'Staff Admin', 'Staff Administrasi', 'FSA', 'Admin 2'],
        'MIS/MSA': ['MIS', 'Staf MIS', 'Staff MIS', 'MSA']
    }
    
    # Convert start_date to datetime
    start_date = datetime.strptime(start_date, '%d/%m/%Y')
    end_date = start_date + timedelta(days=6)
    
    # Convert date column to datetime and remove time component
    df['TRANS. DATE'] = pd.to_datetime(df['TRANS. DATE']).dt.date
    start_date = start_date.date()
    end_date = end_date.date()
    
    # Filter data for the week
    mask = (df['TRANS. DATE'] >= start_date) & (df['TRANS. DATE'] <= end_date)
    weekly_data = df[mask]
    
    # Initialize results dictionary
    results = {
        'Tanggal': f"{start_date.strftime('%d/%m/%Y')} - {end_date.strftime('%d/%m/%Y')}",
        'Manager': 0,
        'Asmen': 0,
        'Admin/FSA': 0,
        'MIS/MSA': 0
    }
    
    # Process each transaction
    for _, row in weekly_data.iterrows():
        desc = str(row['DESCRIPTION']).lower()
        amount = float(row['DEBIT'])
        
        # Proses untuk staf mingguan
        if 'mingguan staf' in desc.lower():
            results['Admin/FSA'] += amount
            continue
            
        # Check each job category
        for category, keywords in job_mappings.items():
            if any(keyword.lower() in desc for keyword in keywords):
                results[category] += amount
                break
    
    return results

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
            results = process_transactions(df, start_date)
            
            # Create DataFrame for display
            df_display = pd.DataFrame([results])
            st.write("Hasil perhitungan:")
            st.write(df_display)
            
            # Create Excel download button
            excel_file = to_excel(df_display)
            st.download_button(
                label="Download Excel",
                data=excel_file,
                file_name=f'tracking_bbm_{start_date.replace("/", "-")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
    except Exception as e:
        st.error(f"Error membaca file: {str(e)}")
