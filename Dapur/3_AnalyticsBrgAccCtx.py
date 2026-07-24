import configparser
from datetime import datetime
import os
import re
import sys
import numpy as np
import pandas as pd

file_coretax = "CtxBarang_temp.xlsx"
file_accurate = "AccBarang_temp.xlsx"
output_filename = "Hasil_Analisis_Barang_temp.xlsx"


def clean_invoice_number(val):
    if pd.isna(val):
        return ""
    val = str(val).strip()
    if "e+" in val.lower() or "e-" in val.lower():
        try:
            f_val = float(val)
            return str(int(f_val))
        except Exception:
            pass
    val = re.sub(r"[^0-9a-zA-Z]", "", val)
    return val


def normalize_name(name, alias_dict=None):
    if pd.isna(name):
        return "UNKNOWN"

    name = str(name).upper().strip()

    if alias_dict:
        for key_alias, target_value in alias_dict.items():
            if key_alias in name:
                return target_value

    name = re.sub(r"^(PT\.|CV\.|UD\.|FA\.)\s*", "", name)
    return name.strip()


def parse_indonesian_date(date_val):
    if pd.isna(date_val):
        return None
    if isinstance(date_val, datetime):
        return date_val
    date_str = str(date_val).strip()
    indo_months = {
        "Jan": "01",
        "Feb": "02",
        "Mar": "03",
        "Apr": "04",
        "Mei": "05",
        "Jun": "06",
        "Jul": "07",
        "Agt": "08",
        "Agu": "08",
        "Sep": "09",
        "Okt": "10",
        "Nov": "11",
        "Nop": "11",
        "Des": "12",
    }
    for k, v in indo_months.items():
        if k in date_str:
            date_str = date_str.replace(k, v)
            break
    try:
        parts = date_str.split()
        if len(parts) == 3:
            day, month, year = parts
            return datetime(int(year), int(month), int(day))
    except Exception:
        pass
    try:
        return pd.to_datetime(date_val, dayfirst=True)
    except Exception:
        return None


def format_final_date_text(dt):
    if dt is None or pd.isna(dt):
        return ""
    try:
        months_map = {
            1: "Jan",
            2: "Feb",
            3: "Mar",
            4: "Apr",
            5: "Mei",
            6: "Jun",
            7: "Jul",
            8: "Agt",
            9: "Sep",
            10: "Okt",
            11: "Nov",
            12: "Des",
        }
        d = dt.day
        m = months_map[dt.month]
        y = dt.year
        return f"{d:02d} {m} {y}"
    except Exception:
        return ""


def add_total_row(df, label_col, numeric_cols):
    if df.empty:
        return df
    sums = df[numeric_cols].sum()
    total_row = {col: "" for col in df.columns}
    total_row[label_col] = "TOTAL"
    for col in numeric_cols:
        total_row[col] = sums[col]
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)


def find_column(df, candidates, search_keyword):
    for col in candidates:
        if col in df.columns:
            return col

    for col in df.columns:
        if search_keyword.lower() in str(col).lower():
            return col

    return None


def main():
    print("--> Memulai proses analisis...")

    alias_map = {}
    config_file = "config.conf"
    if os.path.exists(config_file):
        print(f"--> Memuat konfigurasi dari {config_file}...")
        config = configparser.ConfigParser()
        config.read(config_file)
        if "ALIAS" in config:
            for key, val in config["ALIAS"].items():
                alias_map[key.upper().strip()] = val.strip()
    else:
        print(
            f"--> Peringatan: File '{config_file}' tidak ditemukan. Berjalan tanpa alias eksternal."
        )

    if not os.path.exists(file_coretax) or not os.path.exists(file_accurate):
        print(
            f"--> ERROR: Pastikan file '{file_coretax}' dan '{file_accurate}' ada di folder ini."
        )
        sys.exit()

    try:
        print(f"--> Membaca {file_coretax}...")
        df_coretax = pd.read_excel(file_coretax, dtype=str)
        print(f"--> Membaca {file_accurate}...")
        df_accurate = pd.read_excel(file_accurate, dtype=str)
    except Exception as e:
        print(f"--> Error membaca file Excel: {e}")
        sys.exit()

    c_nama = find_column(
        df_coretax,
        candidates=[
            "Nama Penjual",
            "Nama Penjual Barang Kena Pajak/Barang Kena Pajak Tidak Berwujud/Jasa Kena Pajak",
        ],
        search_keyword="nama",
    )

    c_faktur = find_column(
        df_coretax,
        candidates=[
            "Nomor Faktur Pajak",
            "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Nomor",
        ],
        search_keyword="faktur",
    )

    c_tgl = find_column(
        df_coretax,
        candidates=[
            "Tanggal Faktur Pajak",
            "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Tanggal",
        ],
        search_keyword="tgl",
    )

    c_ppn = find_column(
        df_coretax,
        candidates=["PPN", "PPN (Rupiah)"],
        search_keyword="ppn",
    )

    missing_core = [
        label
        for label, col in [
            ("Nama Penjual", c_nama),
            ("Nomor Faktur", c_faktur),
            ("Tanggal", c_tgl),
            ("PPN", c_ppn),
        ]
        if col is None
    ]
    if missing_core:
        print(f"--> ERROR: Kolom Coretax berikut tidak ditemukan: {missing_core}")
        sys.exit()

    a_faktur = find_column(df_accurate, ["No. Faktur Pajak", "Nomor Faktur"], "faktur")
    a_nama = find_column(df_accurate, ["Nama Pemasok", "Nama Penjual"], "nama")
    a_tgl = find_column(df_accurate, ["Tgl. Pajak", "Tanggal Pajak"], "tgl")
    a_ppn = find_column(df_accurate, ["Jumlah Pajak", "Nilai PPN"], "ppn")

    missing_acc = [
        label
        for label, col in [
            ("No. Faktur Pajak", a_faktur),
            ("Nama Pemasok", a_nama),
            ("Tgl. Pajak", a_tgl),
            ("Jumlah Pajak", a_ppn),
        ]
        if col is None
    ]
    if missing_acc:
        print(f"--> ERROR: Kolom Accurate berikut tidak ditemukan: {missing_acc}")
        sys.exit()

    df_coretax["clean_invoice"] = df_coretax[c_faktur].apply(clean_invoice_number)
    df_coretax["clean_name"] = df_coretax[c_nama].apply(
        normalize_name, alias_dict=alias_map
    )
    df_coretax["dt_obj"] = df_coretax[c_tgl].apply(parse_indonesian_date)
    df_coretax["formatted_date"] = df_coretax["dt_obj"].apply(format_final_date_text)
    df_coretax["PPN_Clean"] = pd.to_numeric(
        df_coretax[c_ppn], errors="coerce"
    ).fillna(0)

    coretax_agg = (
        df_coretax.groupby(["clean_invoice", "clean_name"])
        .agg({"PPN_Clean": "sum", "formatted_date": "first"})
        .reset_index()
    )
    coretax_agg.rename(columns={"PPN_Clean": "PPN_Coretax"}, inplace=True)

    df_accurate["clean_invoice"] = df_accurate[a_faktur].apply(clean_invoice_number)
    df_accurate["clean_name"] = df_accurate[a_nama].apply(
        normalize_name, alias_dict=alias_map
    )
    df_accurate["dt_obj"] = df_accurate[a_tgl].apply(parse_indonesian_date)
    df_accurate["formatted_date"] = df_accurate["dt_obj"].apply(format_final_date_text)
    df_accurate["PPN_Clean"] = pd.to_numeric(
        df_accurate[a_ppn], errors="coerce"
    ).fillna(0)

    accurate_agg = (
        df_accurate.groupby(["clean_invoice", "clean_name"])
        .agg({"PPN_Clean": "sum", "formatted_date": "first"})
        .reset_index()
    )
    accurate_agg.rename(columns={"PPN_Clean": "PPN_Accurate"}, inplace=True)

    print("--> Menggabungkan data...")
    merged = pd.merge(
        coretax_agg,
        accurate_agg,
        on="clean_invoice",
        how="outer",
        suffixes=("_core", "_acc"),
    )
    merged["Final_Name"] = merged["clean_name_core"].combine_first(
        merged["clean_name_acc"]
    )
    merged["Date_Coretax"] = merged["formatted_date_core"]
    merged["Date_Accurate"] = merged["formatted_date_acc"]
    merged["PPN_Coretax"] = merged["PPN_Coretax"].fillna(0)
    merged["PPN_Accurate"] = merged["PPN_Accurate"].fillna(0)
    merged["Selisih"] = merged["PPN_Coretax"] - merged["PPN_Accurate"]

    detail_view = merged[
        [
            "Final_Name",
            "clean_invoice",
            "Date_Coretax",
            "Date_Accurate",
            "PPN_Coretax",
            "PPN_Accurate",
            "Selisih",
        ]
    ].copy()
    detail_view.columns = [
        "Nama Penjual",
        "Nomor Faktur",
        "Tanggal Coretax",
        "Tanggal Accurate",
        "PPN Coretax",
        "PPN Accurate",
        "Selisih",
    ]
    detail_view = detail_view.sort_values(by=["Nama Penjual", "Tanggal Coretax"])
    summary_view = (
        detail_view.groupby("Nama Penjual")[["PPN Coretax", "PPN Accurate", "Selisih"]]
        .sum()
        .reset_index()
    )

    print("--> Menjalankan analisis Daily Matching...")
    detail_view["Grouping_Date"] = (
        detail_view["Tanggal Coretax"]
        .replace("", np.nan)
        .combine_first(detail_view["Tanggal Accurate"])
    )
    daily_check = (
        detail_view.groupby(["Nama Penjual", "Grouping_Date"])[
            ["PPN Coretax", "PPN Accurate"]
        ]
        .sum()
        .reset_index()
    )
    daily_check["Daily_Diff"] = (
        daily_check["PPN Coretax"] - daily_check["PPN Accurate"]
    ).abs()
    matched_groups = daily_check[daily_check["Daily_Diff"] <= 50]
    matched_keys = set(
        zip(matched_groups["Nama Penjual"], matched_groups["Grouping_Date"])
    )

    def is_unmatched(row):
        if abs(row["Selisih"]) <= 50:
            return False
        key = (row["Nama Penjual"], row["Grouping_Date"])
        if key in matched_keys:
            return False
        return True

    unmatched_data = detail_view[detail_view.apply(is_unmatched, axis=1)].copy()
    print(f"--> Total Transaksi: {len(detail_view)} baris.")
    print(f"--> Transaksi yang Match: {len(detail_view) - len(unmatched_data)} baris.")
    print(f"--> Sisa Selisih yang harus dicek: {len(unmatched_data)} baris.")

    rincian_final_raw = unmatched_data[unmatched_data["Selisih"].abs() > 50].copy()
    rincian_final_raw = rincian_final_raw.sort_values(
        by=["Nama Penjual", "Tanggal Coretax"]
    )

    def beri_keterangan(row):
        acc = row["PPN Accurate"]
        core = row["PPN Coretax"]
        selisih = row["Selisih"]
        if acc == 0:
            return "Belum Input di Accurate / Beda Masa"
        elif core == 0:
            return "Input di Accurate Tidak Dikenal Coretax"
        elif abs(selisih) <= 100:
            return "Selisih Pembulatan"
        else:
            return "Selisih Nominal (Input Tidak Sesuai)"

    rincian_final_raw["Keterangan"] = rincian_final_raw.apply(beri_keterangan, axis=1)
    cols = [
        c
        for c in rincian_final_raw.columns
        if c != "Keterangan" and c != "Grouping_Date"
    ] + ["Keterangan"]
    rincian_final_raw = rincian_final_raw[cols]

    numeric_cols = ["PPN Coretax", "PPN Accurate", "Selisih"]
    summary_view_final = add_total_row(summary_view, "Nama Penjual", numeric_cols)
    rincian_final_view = add_total_row(rincian_final_raw, "Nama Penjual", numeric_cols)
    detail_view_final = add_total_row(
        detail_view.drop(columns=["Grouping_Date"]),
        "Nama Penjual",
        numeric_cols,
    )

    print(f"--> Menyimpan ke {output_filename}...")
    try:
        with pd.ExcelWriter(output_filename, engine="xlsxwriter") as writer:
            workbook = writer.book
            fmt_accounting = workbook.add_format(
                {"num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)'}
            )
            fmt_header = workbook.add_format(
                {
                    "bold": True,
                    "bg_color": "#D3D3D3",
                    "border": 1,
                    "align": "center",
                    "valign": "vcenter",
                }
            )
            fmt_total_row = workbook.add_format(
                {
                    "bold": True,
                    "bg_color": "#FFFFCC",
                    "num_format": '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
                    "top": 1,
                }
            )
            fmt_total_label = workbook.add_format(
                {"bold": True, "bg_color": "#FFFFCC", "top": 1}
            )

            def write_sheet(sheet_name, df):
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                worksheet = writer.sheets[sheet_name]
                (max_row, max_col) = df.shape
                money_cols = ["PPN Coretax", "PPN Accurate", "Selisih"]
                money_indices = [
                    df.columns.get_loc(c) for c in money_cols if c in df.columns
                ]

                for idx, col in enumerate(df.columns):
                    col_str = str(col)
                    if not df.empty:
                        max_val_len = df[col].fillna("").astype(str).str.len().max()
                        max_val_len = int(max_val_len) if pd.notna(max_val_len) else 0
                    else:
                        max_val_len = 0

                    final_len = max(len(col_str), max_val_len) + 3
                    if final_len > 50:
                        final_len = 50

                    if idx in money_indices:
                        worksheet.set_column(idx, idx, final_len + 5, fmt_accounting)
                    else:
                        worksheet.set_column(idx, idx, final_len)
                    worksheet.write(0, idx, col_str, fmt_header)

                if not df.empty:
                    label_col_idx = 0
                    worksheet.write(max_row, label_col_idx, "TOTAL", fmt_total_label)
                    for col_idx in money_indices:
                        val = df.iloc[-1, col_idx]
                        if pd.isna(val) or val == "":
                            val = 0
                        worksheet.write(max_row, col_idx, val, fmt_total_row)

            write_sheet("Summary Total", summary_view_final)
            write_sheet("Rincian Selisih", rincian_final_view)
            write_sheet("Detail Data", detail_view_final)

        print(f"--> Selesai! File '{output_filename}' berhasil dibuat.")
    except Exception as e:
        print(f"--> Gagal menyimpan Excel: {e}")
        print("--> Pastikan file Excel output tidak sedang dibuka.")


if __name__ == "__main__":
    main()