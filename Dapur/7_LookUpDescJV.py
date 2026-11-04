import configparser
import os
import re
import openpyxl
from openpyxl.utils import column_index_from_string, get_column_letter


def baca_konfigurasi(config_path="config.conf"):
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"File konfigurasi '{config_path}' tidak ditemukan.")

    config = configparser.ConfigParser()
    config.read(config_path, encoding="utf-8")

    return {
        "dest_key": config.get("LOOKUP", "lo_dest_key").strip().upper(),
        "dest_ret": config.get("LOOKUP", "lo_dest_ret").strip().upper(),
        "fr_key": config.get("LOOKUP", "lo_fr_key").strip().upper(),
        "fr_ret": config.get("LOOKUP", "lo_fr_ret").strip().upper(),
    }


def ekstraksi_kunci_jv(teks):
    if not teks:
        return None

    str_teks = str(teks).strip()

    pola_ref = re.search(r"Ref:\s*([^)]+)", str_teks, re.IGNORECASE)
    if pola_ref:
        return pola_ref.group(1).strip()

    pola_jv = re.search(r"(JV\s*\d+)", str_teks, re.IGNORECASE)
    if pola_jv:
        return pola_jv.group(1).strip()

    return None


def buat_kamus_lookup_accuratem(file_path, col_key_letter, col_val_letter):
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File data sumber '{file_path}' tidak ditemukan.")

    lookup_map = {}
    col_k_idx = column_index_from_string(col_key_letter) - 1
    col_v_idx = column_index_from_string(col_val_letter) - 1

    try:
        import xlrd

        wb = xlrd.open_workbook(file_path)
        sheet = wb.sheet_by_index(0)

        for r in range(sheet.nrows):
            if sheet.ncols > max(col_k_idx, col_v_idx):
                k = sheet.cell_value(r, col_k_idx)
                v = sheet.cell_value(r, col_v_idx)
                k_clean = str(k).strip()

                if k_clean and k_clean not in lookup_map:
                    lookup_map[k_clean] = str(v).strip()

    except Exception as e:
        import pandas as pd

        df = pd.read_excel(file_path, header=None)
        for _, row in df.iterrows():
            if len(row) > max(col_k_idx, col_v_idx):
                k_val = row[col_k_idx]
                v_val = row[col_v_idx]
                k_clean = str(k_val).strip() if pd.notna(k_val) else ""
                v_clean = str(v_val).strip() if pd.notna(v_val) else ""
                if k_clean and k_clean not in lookup_map:
                    lookup_map[k_clean] = v_clean

    return lookup_map


def autofit_kolom(ws):
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value is not None:
                lines = str(cell.value).split("\n")
                pjg_baris = max(len(l) for l in lines)
                if pjg_baris > max_len:
                    max_len = pjg_baris
        if max_len > 0:
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)


def jalankan_proses_lookup():
    file_config = "config.conf"
    file_dest = "Accuratem.xls"
    file_target = "Hasil_Analisis_Barang_Dan_Jasa.xlsx"

    cfg = baca_konfigurasi(file_config)
    print(f"--> Konfigurasi aktif: {cfg}")

    print(f"--> Membaca data referensi dari '{file_dest}'...")
    lookup_map = buat_kamus_lookup_accuratem(
        file_dest, cfg["dest_key"], cfg["dest_ret"]
    )
    print(f"--> Ditemukan {len(lookup_map)} data unik untuk referensi.")

    if not os.path.exists(file_target):
        raise FileNotFoundError(f"File target '{file_target}' tidak ditemukan.")

    wb_target = openpyxl.load_workbook(file_target)

    if "Analisis JV" in wb_target.sheetnames:
        ws = wb_target["Analisis JV"]
    else:
        ws = wb_target.active

    col_fr_ret = cfg["fr_ret"]
    col_fr_key = cfg["fr_key"]
    ws[f"{col_fr_ret}1"] = "Deskripsi JV ACC"

    baris_terisi = 0
    for r in range(2, ws.max_row + 1):
        nilai_asal = ws[f"{col_fr_key}{r}"].value
        kunci = ekstraksi_kunci_jv(nilai_asal)

        if kunci and kunci in lookup_map:
            ws[f"{col_fr_ret}{r}"] = lookup_map[kunci]
            baris_terisi += 1

    print(f"--> Berhasil memetakan {baris_terisi} baris deskripsi.")

    print("--> Menyesuaikan ukuran kolom (Auto-fit)...")
    autofit_kolom(ws)

    wb_target.save(file_target)
    wb_target.close()
    print(f"--> Proses selesai. File '{file_target}' berhasil diperbarui.")


if __name__ == "__main__":
    jalankan_proses_lookup()