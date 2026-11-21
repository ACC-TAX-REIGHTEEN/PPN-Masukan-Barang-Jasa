import os
import sys
import warnings
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def auto_fit_columns(filename, sheet_name):
    try:
        wb = load_workbook(filename)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.active
        for column in ws.columns:
            max_length = 0
            col_letter = get_column_letter(column[0].column)
            for cell in column:
                try:
                    if cell.value is not None:
                        txt_len = len(str(cell.value))
                        if txt_len > max_length:
                            max_length = txt_len
                except Exception:
                    pass
            ws.column_dimensions[col_letter].width = max(max_length + 2, 10)
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
    if val_str.startswith("(") and val_str.endswith(")"):
        val_str = "-" + val_str[1:-1]
    val_str = val_str.replace(".", "").replace(",", ".")
    try:
        return float(val_str)
    except ValueError:
        return 0


def ekstrak_ppn_1(df_source):
    marker_start = "E : PPN 1 (1.10%)"
    marker_end = "Total dari E : PPN 1 (1.10%)"
    try:
        match_start = df_source[df_source[2].astype(str).str.contains(marker_start, na=False, regex=False)]
        match_end = df_source[df_source[2].astype(str).str.contains(marker_end, na=False, regex=False)]
        if not match_start.empty and not match_end.empty:
            idx_start = match_start.index[0]
            idx_end = match_end.index[0]
            df_sub = df_source.iloc[idx_start + 1:idx_end, [3, 5, 7, 16]].copy()
            df_sub.columns = ["Tanggal", "Tgl. Pajak", "No. Referensi", "Jumlah Pajak"]
            df_sub = df_sub[df_sub["Tanggal"].notna() & (df_sub["Tanggal"].astype(str).str.strip() != "")]
            df_sub["No. Referensi"] = df_sub["No. Referensi"].apply(bersihkan_koma_nol)
            df_sub["Jumlah Pajak"] = df_sub["Jumlah Pajak"].apply(konversi_ke_angka)
            return df_sub
    except Exception:
        pass
    return pd.DataFrame(columns=["Tanggal", "Tgl. Pajak", "No. Referensi", "Jumlah Pajak"])


def baca_file_referensi_txt(nama_file):
    if not os.path.exists(nama_file):
        print(f"--> Peringatan: File referensi '{nama_file}' tidak ditemukan.")
        return []
    for enc in ['utf-8', 'latin-1', 'cp1252']:
        try:
            with open(nama_file, 'r', encoding=enc) as f:
                return [line.strip() for line in f if line.strip()]
        except Exception:
            continue
    return []


def baca_hjv_acc_2_bagian(nama_file):
    list_pemasok = []
    dict_referensi = {}

    if not os.path.exists(nama_file):
        print(f"--> Peringatan: File referensi '{nama_file}' tidak ditemukan.")
        return list_pemasok, dict_referensi

    lines = []
    for enc in ['utf-8', 'latin-1', 'cp1252']:
        try:
            with open(nama_file, 'r', encoding=enc) as f:
                lines = [line.strip() for line in f]
            break
        except Exception:
            continue

    current_mode = 'pemasok'

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue

        line_upper = line.upper()
        if line_upper in ['[PEMASOK]', '[NAMA PEMASOK]', '[BARANG_KE_JV]', '[BARANG TO JV]']:
            current_mode = 'pemasok'
            continue
        elif line_upper in ['[REFERENSI]', '[NO REFERENSI]', '[NO. REFERENSI]', '[JV_KE_BARANG]', '[JV TO BARANG]']:
            current_mode = 'referensi'
            continue
        elif line.startswith('---') or line.startswith('==='):
            current_mode = 'referensi'
            continue

        if current_mode == 'pemasok':
            list_pemasok.append(line)
        else:
            if '|' in line:
                parts = line.split('|', 1)
                ref_no = parts[0].strip()
                vendor_opt = parts[1].strip()
            else:
                ref_no = line.strip()
                vendor_opt = ""
            if ref_no:
                dict_referensi[ref_no] = vendor_opt

    return list_pemasok, dict_referensi


def proses_langkah_kedua():
    file_barang = 'AccBarang_temp.xlsx'
    file_jv = 'AccJV_temp.xlsx'
    file_hbrg = 'hbrg_acc.txt'
    file_hjv = 'hjv_acc.txt'

    print("--> Memproses Langkah Kedua: Filter Referensi hbrg_acc.txt & hjv_acc.txt")

    if not os.path.exists(file_barang) or not os.path.exists(file_jv):
        print("--> Langkah Kedua dilewati: AccBarang_temp.xlsx atau AccJV_temp.xlsx tidak ditemukan.")
        return

    list_hbrg = baca_file_referensi_txt(file_hbrg)
    set_hbrg_upper = {p.upper() for p in list_hbrg}

    list_hjv_pemasok, dict_hjv_ref = baca_hjv_acc_2_bagian(file_hjv)
    set_hjv_pemasok_upper = {p.upper() for p in list_hjv_pemasok}
    map_hjv_ref_upper = {k.upper(): (k, v) for k, v in dict_hjv_ref.items()}

    df_barang = None
    df_jv = None
    for eng in ['calamine', 'openpyxl', 'xlrd']:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                if df_barang is None:
                    df_barang = pd.read_excel(
                        file_barang,
                        sheet_name='AccBarang',
                        engine=eng,
                        dtype={'No. Referensi': str, 'No. Faktur Pajak': str, 'Tanggal': str, 'Tgl. Pajak': str, 'Nama Pemasok': str}
                    )
                if df_jv is None:
                    df_jv = pd.read_excel(
                        file_jv,
                        sheet_name='AccJV',
                        engine=eng,
                        dtype={'No. Referensi': str, 'Tanggal': str, 'Tgl. Pajak': str}
                    )
            if df_barang is not None and df_jv is not None:
                break
        except Exception:
            continue

    if df_barang is None or df_jv is None:
        print("--> Gagal Langkah Kedua: Tidak dapat membaca file Excel hasil tahap 1.")
        return

    for col in ['No. Referensi', 'No. Faktur Pajak']:
        if col in df_barang.columns:
            df_barang[col] = df_barang[col].apply(bersihkan_koma_nol)
    if 'No. Referensi' in df_jv.columns:
        df_jv['No. Referensi'] = df_jv['No. Referensi'].apply(bersihkan_koma_nol)

    df_barang['Jumlah Pajak'] = pd.to_numeric(df_barang['Jumlah Pajak'], errors='coerce').fillna(0)
    df_jv['Jumlah Pajak'] = pd.to_numeric(df_jv['Jumlah Pajak'], errors='coerce').fillna(0)

    if 'Nama Pemasok' in df_barang.columns:
        pemasok_barang_unik = [p for p in df_barang['Nama Pemasok'].dropna().astype(str).str.strip().unique() if p != '']
        pemasok_tidak_di_hbrg = [p for p in pemasok_barang_unik if p.upper() not in set_hbrg_upper]

        if pemasok_tidak_di_hbrg:
            print("--> Informasi: Ditemukan 'Nama Pemasok' di AccBarang_temp.xlsx yang tidak ada di hbrg_acc.txt:")
            for p in pemasok_tidak_di_hbrg:
                print(f"    * {p}")
        else:
            print("--> Informasi: Seluruh 'Nama Pemasok' di AccBarang_temp.xlsx sudah cocok dengan hbrg_acc.txt.")

    ada_perubahan = False

    if 'Nama Pemasok' in df_barang.columns and set_hjv_pemasok_upper:
        mask_ke_jv = df_barang['Nama Pemasok'].astype(str).str.strip().str.upper().isin(set_hjv_pemasok_upper)
        if mask_ke_jv.any():
            baris_pindah_ke_jv = df_barang[mask_ke_jv].copy()
            kolom_jv = ['Tanggal', 'Tgl. Pajak', 'No. Referensi', 'Jumlah Pajak']
            baris_untuk_jv = baris_pindah_ke_jv[kolom_jv].copy()

            df_jv = pd.concat([df_jv, baris_untuk_jv], ignore_index=True)
            df_barang = df_barang[~mask_ke_jv].reset_index(drop=True)
            ada_perubahan = True
            print(f"--> Pindah Data: {len(baris_untuk_jv)} baris dipindahkan dari AccBarang -> AccJV (Kunci: Nama Pemasok di hjv_acc.txt)")

    if 'No. Referensi' in df_jv.columns and map_hjv_ref_upper:
        mask_ke_barang = df_jv['No. Referensi'].astype(str).str.strip().str.upper().isin(map_hjv_ref_upper.keys())
        if mask_ke_barang.any():
            baris_pindah_ke_brg = df_jv[mask_ke_barang].copy()
            
            baris_untuk_barang = pd.DataFrame()
            baris_untuk_barang['Tanggal'] = baris_pindah_ke_brg['Tanggal']
            baris_untuk_barang['Tgl. Pajak'] = baris_pindah_ke_brg['Tgl. Pajak']
            baris_untuk_barang['No. Referensi'] = baris_pindah_ke_brg['No. Referensi']
            baris_untuk_barang['No. Faktur Pajak'] = ""
            
            baris_untuk_barang['Nama Pemasok'] = baris_pindah_ke_brg['No. Referensi'].apply(
                lambda x: map_hjv_ref_upper.get(str(x).strip().upper(), ("", ""))[1]
            )
            baris_untuk_barang['Jumlah Pajak'] = baris_pindah_ke_brg['Jumlah Pajak']

            df_barang = pd.concat([df_barang, baris_untuk_barang], ignore_index=True)
            df_jv = df_jv[~mask_ke_barang].reset_index(drop=True)
            ada_perubahan = True
            print(f"--> Pindah Data: {len(baris_untuk_barang)} baris dipindahkan dari AccJV -> AccBarang (Kunci: No. Referensi di hjv_acc.txt)")

    if ada_perubahan:
        kolom_barang_resmi = ["Tanggal", "Tgl. Pajak", "No. Referensi", "No. Faktur Pajak", "Nama Pemasok", "Jumlah Pajak"]
        kolom_jv_resmi = ["Tanggal", "Tgl. Pajak", "No. Referensi", "Jumlah Pajak"]

        df_barang = df_barang[kolom_barang_resmi]
        df_jv = df_jv[kolom_jv_resmi]

        df_barang.to_excel(file_barang, index=False, sheet_name="AccBarang")
        print(f"--> Update Selesai: {file_barang} ({len(df_barang)} baris)")
        auto_fit_columns(file_barang, "AccBarang")

        df_jv.to_excel(file_jv, index=False, sheet_name="AccJV")
        print(f"--> Update Selesai: {file_jv} ({len(df_jv)} baris)")
        auto_fit_columns(file_jv, "AccJV")
    else:
        print("--> Tidak ada perpindahan baris data antara AccBarang dan AccJV.")


def main():
    file_sumber = 'Accuratem.xls'
    print(f"--> Sedang membaca file {file_sumber}...")

    df = None
    for engine in ['calamine', 'openpyxl', 'xlrd']:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                df = pd.read_excel(file_sumber, header=None, engine=engine)
            break
        except Exception:
            continue

    if df is None:
        print("--> Error membaca file: Semua engine gagal membaca file sumber.")
        sys.exit()

    idx_marker_total_ppn = None
    print("--> Memproses Tahap A: AccBarang")
    marker_start_A = "P : PPN (11.00%)"
    marker_end_A = "Total dari P : PPN (11.00%)"

    try:
        idx_start_A = df[df[2].astype(str).str.contains(marker_start_A, na=False, regex=False)].index[0]
        idx_marker_total_ppn = df[df[2].astype(str).str.contains(marker_end_A, na=False, regex=False)].index[0]

        df_barang = df.iloc[idx_start_A + 1:idx_marker_total_ppn, [3, 5, 7, 11, 13, 16]].copy()
        df_barang.columns = ["Tanggal", "Tgl. Pajak", "No. Referensi", "No. Faktur Pajak", "Nama Pemasok", "Jumlah Pajak"]
        df_barang = df_barang[df_barang["Tanggal"].notna() & (df_barang["Tanggal"].astype(str).str.strip() != "")]

        for col in ["No. Referensi", "No. Faktur Pajak"]:
            df_barang[col] = df_barang[col].apply(bersihkan_koma_nol)

        df_barang["Jumlah Pajak"] = df_barang["Jumlah Pajak"].apply(konversi_ke_angka)
        df_barang.to_excel("AccBarang_temp.xlsx", index=False, sheet_name="AccBarang")
        print(f"--> Sukses: AccBarang_temp.xlsx ({len(df_barang)} baris) - Kolom F sudah Angka")
        auto_fit_columns("AccBarang_temp.xlsx", "AccBarang")
    except IndexError:
        print("--> Gagal Tahap A: Marker tidak ditemukan.")

    print("--> Memproses Tahap B: AccJV")
    marker_end_B = "Total dari Transaksi Lainnya"

    if idx_marker_total_ppn is not None:
        try:
            idx_start_B = idx_marker_total_ppn + 1
            df_search_area = df.loc[idx_start_B:]
            matches_end_B = df_search_area[df_search_area[2].astype(str).str.contains(marker_end_B, na=False, regex=False)]

            if not matches_end_B.empty:
                idx_end_B = matches_end_B.index[0]
                df_jv = df.loc[idx_start_B:idx_end_B - 1, [3, 5, 7, 16]].copy()
                df_jv.columns = ["Tanggal", "Tgl. Pajak", "No. Referensi", "Jumlah Pajak"]
                df_jv = df_jv[df_jv["Tanggal"].notna() & (df_jv["Tanggal"].astype(str).str.strip() != "")]
                df_jv["Jumlah Pajak"] = df_jv["Jumlah Pajak"].apply(konversi_ke_angka)

                df_ppn1 = ekstrak_ppn_1(df)
                if not df_ppn1.empty:
                    df_jv = pd.concat([df_ppn1, df_jv], ignore_index=True)

                df_jv.to_excel("AccJV_temp.xlsx", index=False, sheet_name="AccJV")
                print(f"--> Sukses: AccJV_temp.xlsx ({len(df_jv)} baris) - Kolom D sudah Angka")
                auto_fit_columns("AccJV_temp.xlsx", "AccJV")
            else:
                print("--> Gagal Tahap B: Penutup 'Total dari Transaksi Lainnya' tidak ditemukan.")
        except Exception as e:
            print(f"--> Gagal Tahap B: Terjadi kesalahan logika ({e})")
    else:
        print("--> Gagal Tahap B: Acuan dari Tahap A tidak ditemukan.")

    proses_langkah_kedua()

    print("--> Selesai")


if __name__ == "__main__":
    main()