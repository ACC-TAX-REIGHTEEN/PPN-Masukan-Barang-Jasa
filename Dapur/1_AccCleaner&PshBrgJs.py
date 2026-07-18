import os
import sys
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def auto_fit_columns(filename, sheet_name):
    try:
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
                        txt_len = len(str(cell.value))
                        if txt_len > max_length:
                            max_length = txt_len
                except:
                    pass
            
            adjusted_width = (max_length + 2)
            ws.column_dimensions[col_letter].width = adjusted_width
            
        wb.save(filename)
        print(f"--> Auto-fit selesai: {filename}")
    except Exception as e:
        print(f"--> Gagal Auto-fit {filename}: {e}")

def bersihkan_koma_nol(nilai):
    if pd.isna(nilai):
        return ""
    val_str = str(nilai).strip()
    if val_str.endswith(",00"):
        return val_str[:-3]
    if val_str.endswith(".0"):
        return val_str[:-2]
    return val_str

def konversi_ke_angka(nilai):
    if pd.isna(nilai):
        return 0
    
    val_str = str(nilai).strip()
    
    if val_str.endswith(",00"):
        val_str = val_str[:-3]
    elif val_str.endswith(".0"):
        val_str = val_str[:-2]
    
    val_str = val_str.replace(".", "")
    val_str = val_str.replace(",", ".")
    
    try:
        return float(val_str)
    except ValueError:
        return 0

file_sumber = 'Accuratem.xls'
print(f"--> Sedang membaca file {file_sumber}...")

try:
    df = pd.read_excel(file_sumber, header=None)
except Exception as e:
    print(f"--> Error membaca file: {e}")
    sys.exit()

idx_marker_total_ppn = None

print("--> Memproses Tahap A: AccBarang")

marker_start_A = "P : PPN (11.00%)"
marker_end_A = "Total dari P : PPN (11.00%)"

try:
    idx_start_A = df[df[2].astype(str).str.contains(marker_start_A, na=False, regex=False)].index[0]
    idx_marker_total_ppn = df[df[2].astype(str).str.contains(marker_end_A, na=False, regex=False)].index[0]

    df_barang = df.iloc[idx_start_A + 1 : idx_marker_total_ppn].copy()
    df_barang = df_barang[[3, 4, 6, 9, 10, 12]]
    df_barang.columns = ["Tanggal", "Tgl. Pajak", "No. Referensi", "No. Faktur Pajak", "Nama Pemasok", "Jumlah Pajak"]

    for col in ["No. Referensi", "No. Faktur Pajak"]:
        df_barang[col] = df_barang[col].apply(bersihkan_koma_nol)
        
    df_barang["Jumlah Pajak"] = df_barang["Jumlah Pajak"].apply(konversi_ke_angka)

    df_barang.to_excel("AccBarang_temp.xlsx", index=False, sheet_name="AccBarang")
    print(f"--> Sukses: AccBarang.xlsx ({len(df_barang)} baris) - Kolom F sudah Angka")
    auto_fit_columns("AccBarang_temp.xlsx", "AccBarang")

except IndexError:
    print("--> Gagal Tahap A: Marker tidak ditemukan.")

print("--> Memproses Tahap B: AccJV")

marker_end_B = "Total dari Transaksi Lainnya"

if idx_marker_total_ppn is not None:
    try:
        idx_start_B = idx_marker_total_ppn + 2
        df_search_area = df.loc[idx_start_B:] 
        matches_end_B = df_search_area[df_search_area[2].astype(str).str.contains(marker_end_B, na=False, regex=False)]
        
        if not matches_end_B.empty:
            idx_end_B = matches_end_B.index[0]

            df_jv = df.loc[idx_start_B : idx_end_B - 1].copy()
            df_jv = df_jv[[3, 4, 6, 12]]
            df_jv.columns = ["Tanggal", "Tgl. Pajak", "No. Referensi", "Jumlah Pajak"]
            df_jv["Jumlah Pajak"] = df_jv["Jumlah Pajak"].apply(konversi_ke_angka)

            df_jv.to_excel("AccJV_temp.xlsx", index=False, sheet_name="AccJV")
            print(f"--> Sukses: AccJV.xlsx ({len(df_jv)} baris) - Kolom D sudah Angka")
            auto_fit_columns("AccJV_temp.xlsx", "AccJV")
        else:
             print("--> Gagal Tahap B: Penutup 'Total dari Transaksi Lainnya' tidak ditemukan.")

    except Exception as e:
        print(f"--> Gagal Tahap B: Terjadi kesalahan logika ({e})")
else:
    print("--> Gagal Tahap B: Acuan dari Tahap A tidak ditemukan.")

print("--> Selesai")