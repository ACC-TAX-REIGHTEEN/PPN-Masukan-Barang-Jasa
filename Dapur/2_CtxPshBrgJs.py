import os
import re
import sys
import warnings
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

    try:
        tgl = tanggal_obj.day
        bln = bulan_map[tanggal_obj.month]
        thn = tanggal_obj.year
        return f"{tgl} {bln} {thn}"
    except AttributeError:
        return str(tanggal_obj)


def auto_fit_columns(filename, sheet_name):
    try:
        wb = load_workbook(filename)
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            ws = wb.active

        for column in ws.columns:
            max_length = 0
            col_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    if cell.value is not None:
                        length = len(str(cell.value))
                        if length > max_length:
                            max_length = length
                except Exception:
                    pass

            adjusted_width = max(max_length + 2, 10)
            ws.column_dimensions[col_letter].width = adjusted_width

        wb.save(filename)
        print(f"--> Format auto-fit selesai untuk: {filename} (Sheet: {sheet_name})")
    except Exception as e:
        print(f"--> Gagal Auto-fit {filename}: {e}")


def parse_nominal(val):
    if pd.isnull(val) or val is None:
        return None
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip()
    if not val_str:
        return None
    if ',' in val_str and '.' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
            val_str = val_str.replace('.', '').replace(',', '.')
        else:
            val_str = val_str.replace(',', '')
    elif ',' in val_str:
        if val_str.count(',') > 1:
            val_str = val_str.replace(',', '')
        else:
            val_str = val_str.replace(',', '.')
    elif '.' in val_str:
        parts = val_str.split('.')
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3):
            val_str = val_str.replace('.', '')
    try:
        return float(val_str)
    except ValueError:
        return None


def baca_filter_txt_kondisional(nama_file):
    if not os.path.exists(nama_file):
        print(f"--> Peringatan: File {nama_file} tidak ditemukan.")
        return []

    rules = []
    with open(nama_file, 'r', encoding='utf-8') as f:
        for line in f:
            line_str = line.strip()
            if not line_str:
                continue

            if '|' in line_str:
                parts = line_str.split('|', 1)
                vendor = parts[0].strip()
                nom = parse_nominal(parts[1])
            else:
                vendor = line_str
                nom = None

            if not vendor:
                continue

            pattern = re.compile(rf'(?<![A-Z0-9]){re.escape(vendor.upper())}(?![A-Z0-9])')
            rules.append({'vendor': vendor, 'pattern': pattern, 'nominal': nom})

    return rules


def cari_nama_kolom(df, daftar_kemungkinan):
    for nama in daftar_kemungkinan:
        target = nama.lower()
        for col in df.columns:
            if str(col).strip().lower() == target:
                return col

    for nama in daftar_kemungkinan:
        target = nama.lower()
        for col in df.columns:
            if target in str(col).strip().lower():
                return col
    return None


def deteksi_kolom_string(file_sumber):
    kata_kunci_str = ('npwp', 'nomor', 'kode', 'nik', 'nitku', 'referensi')
    try:
        header_df = pd.read_excel(file_sumber, sheet_name='data', nrows=0, engine='calamine')
    except Exception:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                header_df = pd.read_excel(file_sumber, sheet_name='data', nrows=0, engine='openpyxl')
        except Exception:
            return {}

    dtype_map = {}
    for col in header_df.columns:
        col_lower = str(col).lower()
        if any(k in col_lower for k in kata_kunci_str):
            dtype_map[col] = str
    return dtype_map


def cocok_aturan(penjual_val, ppn_val, rules):
    for rule in rules:
        if rule['pattern'].search(penjual_val):
            rule_nom = rule['nominal']
            if rule_nom is not None:
                if ppn_val is not None and abs(ppn_val - rule_nom) < 1.0:
                    return True
            else:
                return True
    return False


def pisahkan_jv_dan_barang(df, rules_jv, rules_brg):
    col_penjual = cari_nama_kolom(df, [
        "Nama Penjual Barang Kena Pajak/Barang Kena Pajak Tidak Berwujud/Jasa Kena Pajak",
        "Nama Penjual",
        "Nama Pemasok"
    ])

    col_ppn = cari_nama_kolom(df, ["Jumlah PPN", "Nilai PPN", "PPN"])

    if not col_penjual:
        print("--> Error: Kolom nama penjual tidak ditemukan di Coretax.")
        return pd.DataFrame(columns=df.columns), pd.DataFrame(columns=df.columns)

    penjual_list = df[col_penjual].fillna('').astype(str).str.upper().tolist()
    ppn_list = [parse_nominal(v) for v in df[col_ppn]] if col_ppn else [None] * len(df)

    indices_jv = []
    indices_brg = []
    konflik_vendors = set()
    diabaikan = 0

    for idx, penjual_val, ppn_val in zip(df.index, penjual_list, ppn_list):
        is_jv = cocok_aturan(penjual_val, ppn_val, rules_jv)
        is_brg = cocok_aturan(penjual_val, ppn_val, rules_brg)

        if is_jv and is_brg:
            konflik_vendors.add(penjual_val)

        if is_jv:
            indices_jv.append(idx)
        elif is_brg:
            indices_brg.append(idx)
        else:
            diabaikan += 1

    if konflik_vendors:
        daftar_konflik = ', '.join(sorted(konflik_vendors))
        print(f"--> Peringatan: Vendor berikut cocok di hjv.txt dan hbrg.txt sekaligus (diprioritaskan ke JV): {daftar_konflik}")

    if diabaikan > 0:
        print(f"--> Info: {diabaikan} baris tidak terdaftar di hjv.txt maupun hbrg.txt (dilewati).")

    df_jv = df.loc[indices_jv].copy()
    df_barang = df.loc[indices_brg].copy()

    return df_jv, df_barang


def main():
    print("--> Mulai Proses Pemisahan Data Coretax")
    file_sumber = 'Coretaxm.xlsx'
    print(f"--> Membaca {file_sumber}...")

    if not os.path.exists(file_sumber):
        print(f"--> Error: File {file_sumber} tidak ditemukan di folder ini.")
        sys.exit()

    dtype_map = deteksi_kolom_string(file_sumber)

    df = None
    try:
        df = pd.read_excel(file_sumber, sheet_name='data', dtype=dtype_map, engine='calamine')
    except Exception:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(file_sumber, sheet_name='data', dtype=dtype_map, engine='openpyxl')
        except Exception as e:
            print(f"--> Error membaca file {file_sumber}: {e}")
            sys.exit()

    col_tanggal_v1 = 'Tanggal Faktur Pajak'
    col_tanggal_v2 = 'Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Tanggal'

    if col_tanggal_v1 in df.columns:
        col_tanggal = col_tanggal_v1
    elif col_tanggal_v2 in df.columns:
        col_tanggal = col_tanggal_v2
    else:
        col_tanggal = None

    if col_tanggal:
        df[col_tanggal] = pd.to_datetime(df[col_tanggal])
        df[col_tanggal] = df[col_tanggal].apply(format_tanggal_indonesia)

    rules_jv = baca_filter_txt_kondisional('hjv.txt')
    rules_brg = baca_filter_txt_kondisional('hbrg.txt')

    if not rules_brg:
        print("--> Error: File 'hbrg.txt' wajib diisi (minimal 1 aturan vendor). Proses dibatalkan.")
        sys.exit()

    df_jv, df_barang = pisahkan_jv_dan_barang(df, rules_jv, rules_brg)

    files_to_save = {
        'CtxJV_temp.xlsx': (df_jv, 'CoretaxJV'),
        'CtxBarang_temp.xlsx': (df_barang, 'CoretaxBarang')
    }

    for filename, (data, sheetname) in files_to_save.items():
        print(f"--> Menyimpan file: {filename} dengan Sheet: {sheetname}...")
        data.to_excel(filename, index=False, sheet_name=sheetname)
        auto_fit_columns(filename, sheetname)

    print("--> Selesai! Pemisahan data berhasil diproses.")


if __name__ == "__main__":
    main()