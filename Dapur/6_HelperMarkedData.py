import os
import warnings
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
import pandas as pd

FILE_HASIL = "Hasil_Analisis_Barang_Dan_Jasa.xlsx"
MODE_WARNA = "BARIS_DATA"

REKONSILIASI_CONFIG = [
    {
        "sheet_name": "Analisis JV",
        "file_ctx": "CtxJV_temp.xlsx",
    },
    {
        "sheet_name": "Detail Data",
        "file_ctx": "CtxBarang_temp.xlsx",
    },
]


def baca_excel_fleksibel(file_path, **kwargs):
    try:
        return pd.read_excel(file_path, engine="calamine", **kwargs)
    except Exception:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            return pd.read_excel(file_path, engine="openpyxl", **kwargs)


def clean_str(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip().upper()


def clean_faktur(val):
    if pd.isna(val) or val is None:
        return ""
    return str(val).strip().replace(".", "").replace("-", "")


def parse_date(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, (pd.Timestamp, pd.DatetimeIndex)) or hasattr(
        val, "strftime"
    ):
        try:
            return val.strftime("%Y-%m-%d")
        except Exception:
            pass

    val_str = str(val).strip()
    if not val_str:
        return ""

    if len(val_str) >= 10 and val_str[4] == "-" and val_str[7] == "-":
        try:
            return pd.to_datetime(val_str[:10], format="%Y-%m-%d").strftime(
                "%Y-%m-%d"
            )
        except Exception:
            pass

    bulan_id_en = {
        "JAN": "JAN",
        "PEB": "FEB",
        "FEB": "FEB",
        "MAR": "MAR",
        "APR": "APR",
        "MEI": "MAY",
        "JUN": "JUN",
        "JUL": "JUL",
        "AGU": "AUG",
        "AGS": "AUG",
        "SEP": "SEP",
        "OKT": "OCT",
        "NOP": "NOV",
        "NOV": "NOV",
        "DES": "DEC",
    }
    val_conv = val_str.upper()
    for id_m, en_m in bulan_id_en.items():
        if id_m in val_conv:
            val_conv = val_conv.replace(id_m, en_m)
            break

    try:
        dt = pd.to_datetime(val_conv, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    except Exception:
        pass
    return val_str


def clean_ppn(val):
    if pd.isna(val) or val is None:
        return None
    if isinstance(val, (int, float)):
        return round(float(val), 2)
    val_str = str(val).strip()
    if not val_str or val_str == "-":
        return None
    if "." in val_str and "," in val_str:
        if val_str.find(".") < val_str.find(","):
            val_str = val_str.replace(".", "").replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    elif "." in val_str:
        parts = val_str.split(".")
        if len(parts) > 1 and len(parts[-1]) == 3:
            val_str = val_str.replace(".", "")
    elif "," in val_str:
        parts = val_str.split(",")
        if len(parts) > 1 and len(parts[-1]) == 2:
            val_str = val_str.replace(",", ".")
        else:
            val_str = val_str.replace(",", "")
    try:
        return round(float(val_str), 2)
    except Exception:
        return None


def get_cell_value(ws, row, col):
    if col is None:
        return None
    cell = ws.cell(row=row, column=col)
    if cell.value is not None:
        return cell.value
    for rng in ws.merged_cells.ranges:
        if cell.coordinate in rng:
            return ws.cell(row=rng.min_row, column=rng.min_col).value
    return None


def temukan_kolom_ws(ws, daftar_nama, baris_header=1):
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=baris_header, column=col).value
        if val is not None:
            val_clean = str(val).strip().upper()
            for nama in daftar_nama:
                if val_clean == nama.strip().upper():
                    return col
    return None


def cari_kolom_df(df, daftar_nama):
    for col in df.columns:
        col_clean = str(col).strip().upper()
        for nama in daftar_nama:
            if col_clean == nama.strip().upper():
                return col
    return None


def muat_dict_coretax(file_ctx):
    print(f"--> Membaca referensi Coretax: {file_ctx}...")
    df_ctx = baca_excel_fleksibel(file_ctx)

    col_tgl = cari_kolom_df(
        df_ctx, ["Tanggal Faktur Pajak", "Tanggal Coretax", "Tanggal"]
    )
    col_nama = cari_kolom_df(
        df_ctx, ["Nama Penjual", "Nama Pemasok", "Nama Lawan Transaksi"]
    )
    col_ppn = cari_kolom_df(
        df_ctx, ["PPN", "Jumlah PPN", "Nilai PPN", "PPN Coretax"]
    )
    col_faktur = cari_kolom_df(
        df_ctx, ["Nomor Faktur Pajak", "Nomor Faktur", "No Faktur"]
    )
    col_status = cari_kolom_df(df_ctx, ["Status Faktur", "Status"])

    ctx_dict = {}
    ctx_faktur_dict = {}

    for _, row in df_ctx.iterrows():
        tgl = parse_date(row.get(col_tgl)) if col_tgl else ""
        nama = clean_str(row.get(col_nama)) if col_nama else ""
        ppn = clean_ppn(row.get(col_ppn)) if col_ppn else None

        status = ""
        if col_status and pd.notna(row.get(col_status)):
            status = str(row.get(col_status)).strip()

        if tgl and nama and ppn is not None:
            ctx_dict[(tgl, nama, ppn)] = status

        if col_faktur and pd.notna(row.get(col_faktur)):
            no_faktur = clean_faktur(row.get(col_faktur))
            if no_faktur:
                ctx_faktur_dict[no_faktur] = status

    print(
        f"    Dimuat: {len(ctx_dict)} data tuple, {len(ctx_faktur_dict)} nomor faktur."
    )
    return ctx_dict, ctx_faktur_dict


def proses_sheet(ws, ctx_dict, ctx_faktur_dict, sheet_name):
    print(f"--> Memproses pencocokan pada sheet: {sheet_name}...")

    col_tgl = temukan_kolom_ws(
        ws, ["Tanggal Coretax", "Tanggal Faktur Pajak", "Tanggal"]
    )
    col_nama = temukan_kolom_ws(
        ws, ["Nama Pemasok", "Nama Penjual", "Nama Lawan Transaksi"]
    )
    col_ppn = temukan_kolom_ws(ws, ["PPN Coretax", "PPN"])
    col_faktur = temukan_kolom_ws(
        ws, ["Nomor Faktur", "Nomor Faktur Pajak", "No Faktur"]
    )

    col_status = temukan_kolom_ws(ws, ["Status Faktur", "Status"])
    if not col_status:
        col_status = 8  

    header_cell = ws.cell(row=1, column=col_status, value="Status Faktur")
    header_cell.font = Font(bold=False, size=11)
    header_cell.alignment = Alignment(horizontal="center", vertical="center")

    pink_fill = PatternFill(
        start_color="F8CECC", end_color="F8CECC", fill_type="solid"
    )
    yellow_fill = PatternFill(
        start_color="FFF2CC", end_color="FFF2CC", fill_type="solid"
    )

    cocok_count = 0
    for r in range(2, ws.max_row + 1):
        val_check = (
            get_cell_value(ws, r, col_tgl)
            or get_cell_value(ws, r, col_nama)
            or get_cell_value(ws, r, 1)
        )
        if val_check and "TOTAL" in str(val_check).upper():
            continue

        tgl_raw = get_cell_value(ws, r, col_tgl)
        nama_raw = get_cell_value(ws, r, col_nama)
        ppn_raw = get_cell_value(ws, r, col_ppn)
        faktur_raw = get_cell_value(ws, r, col_faktur)

        if tgl_raw is None and r - 1 >= 2 and col_tgl:
            tgl_raw = get_cell_value(ws, r - 1, col_tgl)
        if nama_raw is None and r - 1 >= 2 and col_nama:
            nama_raw = get_cell_value(ws, r - 1, col_nama)

        if ppn_raw is None and col_ppn:
            if (
                r + 1 <= ws.max_row
                and ws.cell(row=r + 1, column=col_ppn).value is not None
            ):
                ppn_raw = ws.cell(row=r + 1, column=col_ppn).value
            elif r - 1 >= 2 and ws.cell(row=r - 1, column=col_ppn).value is not None:
                ppn_raw = ws.cell(row=r - 1, column=col_ppn).value

        tgl = parse_date(tgl_raw)
        nama = clean_str(nama_raw)
        ppn = clean_ppn(ppn_raw)
        no_faktur = clean_faktur(faktur_raw)

        status_val = None
        key = (tgl, nama, ppn)
        if key in ctx_dict:
            status_val = ctx_dict[key]
        elif no_faktur and no_faktur in ctx_faktur_dict:
            status_val = ctx_faktur_dict[no_faktur]

        if status_val:
            cocok_count += 1
            cell_status = ws.cell(row=r, column=col_status, value=status_val)
            cell_status.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            status_upper = status_val.upper()

            fill_to_apply = None
            if "CANCELED" in status_upper or "BATAL" in status_upper:
                fill_to_apply = pink_fill
            elif "PENDING" in status_upper:
                fill_to_apply = yellow_fill

            if fill_to_apply:
                if MODE_WARNA == "HANYA_STATUS":
                    cell_status.fill = fill_to_apply
                elif MODE_WARNA == "BARIS_DATA":
                    for c in range(1, col_status + 1):
                        ws.cell(row=r, column=c).fill = fill_to_apply

    print(
        f"    Selesai sheet '{sheet_name}': {cocok_count} baris berhasil dicocokkan."
    )


def jalankan_rekonsiliasi():
    if not os.path.exists(FILE_HASIL):
        raise FileNotFoundError(f"File {FILE_HASIL} tidak ditemukan!")

    wb = openpyxl.load_workbook(FILE_HASIL)

    for item in REKONSILIASI_CONFIG:
        sheet_target = item["sheet_name"]
        file_patokan = item["file_ctx"]

        if sheet_target not in wb.sheetnames:
            print(
                f"[Lewati] Sheet '{sheet_target}' tidak ditemukan di {FILE_HASIL}."
            )
            continue

        if not os.path.exists(file_patokan):
            print(
                f"[Lewati] File referensi '{file_patokan}' tidak ditemukan di folder."
            )
            continue

        ctx_dict, ctx_faktur_dict = muat_dict_coretax(file_patokan)
        ws = wb[sheet_target]
        proses_sheet(ws, ctx_dict, ctx_faktur_dict, sheet_target)

    wb.save(FILE_HASIL)
    print(f"--> Sukses! Seluruh pembaruan disimpan ke {FILE_HASIL}.")


if __name__ == "__main__":
    jalankan_rekonsiliasi()