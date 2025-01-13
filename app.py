import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from io import BytesIO

def process_transactions(data_text, start_date):
    # Convert text data to DataFrame
    rows = []
    for line in data_text.strip().split('\n'):
        parts = line.split()
        voucher = parts[0]
        date = parts[1]
        description = ' '.join(parts[2:-1])
        amount = float(parts[-1])
        rows.append([voucher, date, description, amount])
    
    df = pd.DataFrame(rows, columns=['Voucher', 'Date', 'Description', 'Amount'])
    df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')
    
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
    
    # Filter data for the week
    mask = (df['Date'] >= start_date) & (df['Date'] <= end_date)
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
        desc = row['Description'].lower()
        amount = row['Amount']
        
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
        
        # Apply header format
        for col_num, value in enumerate(df.columns.values):
            worksheet.write(0, col_num, value, header_format)
            
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

# Text area for input data
data_input = st.text_area("Masukkan data transaksi:", height=200)

if st.button('Proses Data'):
    if data_input:
        results = process_transactions(data_input, start_date)
        
        # Create DataFrame for display
        df_display = pd.DataFrame([results])
        st.write(df_display)
        
        # Create Excel download button
        excel_file = to_excel(df_display)
        st.download_button(
            label="Download Excel",
            data=excel_file,
            file_name=f'tracking_bbm_{start_date.replace("/", "-")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
