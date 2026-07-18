import os
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def format_tanggal_indonesia(tanggal_obj):
    if pd.isnull(tanggal_obj):
        return ""
    
    bulan_map = {
        1: 'Jan', 2: 'Feb', 3: 'Mar', 4: 'Apr', 5: 'Mei', 6: 'Jun',
        7: 'Jul', 8: 'Agu', 9: 'Sep', 10: 'Okt', 11: 'Nop', 12: 'Des'
    }
    
    tgl = tanggal_obj.day
    bln = bulan_map[tanggal_obj.month]
    thn = tanggal_obj.year
    
    return f"{tgl} {bln} {thn}"

def auto_fit_columns(filename, sheet_name):
    wb = load_workbook(filename)
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.active
    
    for column in ws.columns:
        max_length = 0
        column = [cell for cell in column]
        col_letter = get_column_letter(column[0].column)
        
        for cell in column:
            try:
                if cell.value:
                    length = len(str(cell.value))
                    if length > max_length:
                        max_length = length
            except:
                pass
        
        adjusted_width = (max_length + 2)
        ws.column_dimensions[col_letter].width = adjusted_width
        
    wb.save(filename)
    print(f"--> Format auto-fit selesai untuk: {filename} (Sheet: {sheet_name})")

def baca_filter_txt(nama_file):
    if not os.path.exists(nama_file):
        print(f"--> Peringatan: File {nama_file} tidak ditemukan. Mengembalikan list kosong.")
        return []
    
    with open(nama_file, 'r') as f:
        lines = [line.strip() for line in f if line.strip()]
    return lines

def filter_data(dataframe, keywords, column_name='Nama Penjual'):
    if not keywords:
        return pd.DataFrame(columns=dataframe.columns)
    
    pattern = '|'.join([k for k in keywords]) 
    mask = dataframe[column_name].astype(str).str.contains(pattern, case=False, na=False)
    return dataframe[mask]

def main():
    print("--> Mulai Proses")
    file_sumber = 'Coretaxm.xlsx'
    print(f"--> Membaca {file_sumber}...")

    try:
        df = pd.read_excel(file_sumber, sheet_name='data', 
                           dtype={'NPWP Penjual': str, 'Nomor Faktur Pajak': str})
    except FileNotFoundError:
        print(f"--> Error: File {file_sumber} tidak ditemukan di folder ini.")
        sys.exit()

    df['Tanggal Faktur Pajak'] = pd.to_datetime(df['Tanggal Faktur Pajak'])
    df['Tanggal Faktur Pajak'] = df['Tanggal Faktur Pajak'].apply(format_tanggal_indonesia)

    keywords_jv = baca_filter_txt('hjv.txt')
    keywords_brg = baca_filter_txt('hbrg.txt')

    df_jv = filter_data(df, keywords_jv)
    df_barang = filter_data(df, keywords_brg)

    files_to_save = {
        'CtxJV_temp.xlsx': (df_jv, 'CoretaxJV'),
        'CtxBarang_temp.xlsx': (df_barang, 'CoretaxBarang')
    }

    for filename, (data, sheetname) in files_to_save.items():
        print(f"--> Menyimpan file: {filename} dengan Sheet: {sheetname}...")
        data.to_excel(filename, index=False, sheet_name=sheetname)
        auto_fit_columns(filename, sheetname)

    print("--> Selesai! Semua file berhasil dibuat.")

if __name__ == "__main__":
    main()
