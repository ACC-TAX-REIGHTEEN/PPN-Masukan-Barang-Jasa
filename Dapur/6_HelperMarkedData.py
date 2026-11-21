from copy import copy
import os
import warnings
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
import pandas as pd

FILE_HASIL = "Hasil_Analisis_Barang_Dan_Jasa.xlsx"
FILE_ALIAS = "alias.txt"
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
    return " ".join(str(val).strip().upper().split())


def clean_faktur(val):
    if pd.isna(val) or val is None:
        return ""
    if isinstance(val, float) and val.is_integer():
        val = int(val)
    val_str = (
        str(val).strip().replace(".", "").replace("-", "").replace("'", "")
    )
    if val_str.endswith(".0"):
        val_str = val_str[:-2]
    return val_str.upper()


def muat_alias(file_path=FILE_ALIAS):
    alias_dict = {}
    if not os.path.exists(file_path):
        print(
            f"[Peringatan] File alias '{file_path}' tidak ditemukan. Pencocokan menggunakan nama asli."
        )
        return alias_dict

    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        for baris, line in enumerate(f, 1):
            line = line.strip()
            if not line or line.startswith("#"):
                continue
            if "=" in line:
                kiri, kanan = line.split("=", 1)
                k_clean = clean_str(kiri)
                v_clean = clean_str(kanan)
                if k_clean and v_clean:
                    alias_dict[k_clean] = v_clean
                    k_no_dot = (
                        k_clean.replace("PT.", "PT")
                        .replace("CV.", "CV")
                        .replace("TBK.", "TBK")
                    )
                    k_no_dot = " ".join(k_no_dot.split())
                    alias_dict[k_no_dot] = v_clean

    print(f"--> Dimuat {len(alias_dict)} pemetaan alias dari '{file_path}'.")
    return alias_dict


def standarkan_nama(nama, alias_dict):
    c = clean_str(nama)
    if not c or not alias_dict:
        return c
    if c in alias_dict:
        return alias_dict[c]
    c_no_dot = (
        c.replace("PT.", "PT").replace("CV.", "CV").replace("TBK.", "TBK")
    )
    c_no_dot = " ".join(c_no_dot.split())
    if c_no_dot in alias_dict:
        return alias_dict[c_no_dot]
    return c


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


def muat_dict_coretax(file_ctx, alias_dict=None):
    if alias_dict is None:
        alias_dict = {}
    print(f"--> Membaca referensi Coretax: {file_ctx}...")
    df_ctx = baca_excel_fleksibel(file_ctx)

    col_tgl = cari_kolom_df(
        df_ctx,
        [
            "Tanggal Faktur Pajak",
            "Tanggal Coretax",
            "Tanggal",
            "Tgl Faktur Pajak",
            "Tgl Faktur",
        ],
    )
    col_nama = cari_kolom_df(
        df_ctx,
        [
            "Nama Penjual",
            "Nama Pemasok",
            "Nama Lawan Transaksi",
            "Nama PKP Penjual",
        ],
    )
    col_ppn = cari_kolom_df(
        df_ctx,
        ["PPN", "Jumlah PPN", "Nilai PPN", "PPN Coretax", "Jumlah PPN (Rp)"],
    )
    col_faktur = cari_kolom_df(
        df_ctx,
        [
            "Nomor Faktur Pajak",
            "Nomor Faktur",
            "No Faktur",
            "No. Faktur Pajak",
            "No Faktur Pajak",
        ],
    )
    col_status = cari_kolom_df(
        df_ctx, ["Status Faktur", "Status", "Status Persetujuan", "Status Approval"]
    )

    ctx_dict = {}
    ctx_faktur_dict = {}

    for _, row in df_ctx.iterrows():
        tgl = parse_date(row.get(col_tgl)) if col_tgl else ""
        nama_raw = clean_str(row.get(col_nama)) if col_nama else ""
        nama = standarkan_nama(nama_raw, alias_dict)
        ppn = clean_ppn(row.get(col_ppn)) if col_ppn else None

        status = ""
        if col_status and pd.notna(row.get(col_status)):
            status = str(row.get(col_status)).strip()

        if tgl and nama and ppn is not None:
            key = (tgl, nama, ppn)
            if (
                key not in ctx_dict
                or "CANCEL" in status.upper()
                or "BATAL" in status.upper()
            ):
                ctx_dict[key] = status

        if col_faktur and pd.notna(row.get(col_faktur)):
            no_faktur = clean_faktur(row.get(col_faktur))
            if no_faktur:
                ctx_faktur_dict[no_faktur] = status
                ctx_faktur_dict[no_faktur.lstrip("0")] = status

    print(
        f"--> Dimuat: {len(ctx_dict)} data tuple, {len(ctx_faktur_dict)} nomor faktur."
    )
    return ctx_dict, ctx_faktur_dict


def proses_sheet(
    ws, ctx_dict, ctx_faktur_dict, sheet_name, alias_dict, cancel_per_supplier
):
    print(f"--> Memproses pencocokan pada sheet: {sheet_name}...")

    col_tgl = temukan_kolom_ws(
        ws,
        [
            "Tanggal Coretax",
            "Tanggal Faktur Pajak",
            "Tanggal",
            "Tgl Faktur Pajak",
        ],
    )
    col_nama = temukan_kolom_ws(
        ws,
        [
            "Nama Pemasok",
            "Nama Penjual",
            "Nama Lawan Transaksi",
            "Nama PKP Penjual",
        ],
    )
    col_ppn = temukan_kolom_ws(ws, ["PPN Coretax", "PPN", "Jumlah PPN"])
    col_faktur = temukan_kolom_ws(
        ws, ["Nomor Faktur", "Nomor Faktur Pajak", "No Faktur", "No. Faktur"]
    )

    col_status = temukan_kolom_ws(ws, ["Status Faktur", "Status"])
    if not col_status:
        col_status = ws.max_column + 1

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
    last_nama = ""

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

        if nama_raw:
            last_nama = nama_raw
        elif r - 1 >= 2 and last_nama:
            nama_raw = last_nama

        if tgl_raw is None and r - 1 >= 2 and col_tgl:
            tgl_raw = get_cell_value(ws, r - 1, col_tgl)

        if ppn_raw is None and col_ppn:
            if (
                r + 1 <= ws.max_row
                and ws.cell(row=r + 1, column=col_ppn).value is not None
            ):
                ppn_raw = ws.cell(row=r + 1, column=col_ppn).value
            elif (
                r - 1 >= 2
                and ws.cell(row=r - 1, column=col_ppn).value is not None
            ):
                ppn_raw = ws.cell(row=r - 1, column=col_ppn).value

        tgl = parse_date(tgl_raw)
        nama = standarkan_nama(nama_raw, alias_dict)
        ppn = clean_ppn(ppn_raw)
        no_faktur = clean_faktur(faktur_raw)

        status_val = None
        key = (tgl, nama, ppn)
        if key in ctx_dict:
            status_val = ctx_dict[key]
        elif no_faktur:
            if no_faktur in ctx_faktur_dict:
                status_val = ctx_faktur_dict[no_faktur]
            elif no_faktur.lstrip("0") in ctx_faktur_dict:
                status_val = ctx_faktur_dict[no_faktur.lstrip("0")]

        if not status_val:
            status_existing = get_cell_value(ws, r, col_status)
            if status_existing and any(
                k in str(status_existing).upper()
                for k in ["CANCEL", "BATAL", "APPROVED", "PENDING"]
            ):
                status_val = str(status_existing).strip()

        if status_val:
            cocok_count += 1
            cell_status = ws.cell(row=r, column=col_status, value=status_val)
            cell_status.alignment = Alignment(
                horizontal="center", vertical="center"
            )
            status_upper = status_val.upper()

            fill_to_apply = None
            if "CANCEL" in status_upper or "BATAL" in status_upper:
                fill_to_apply = pink_fill

                nilai_ppn = clean_ppn(get_cell_value(ws, r, col_ppn))
                if nilai_ppn is None:
                    col_selisih_ws = temukan_kolom_ws(
                        ws, ["Selisih", "Selisih PPN"]
                    )
                    col_acc_ws = temukan_kolom_ws(
                        ws, ["PPN Accurate", "PPN Acc"]
                    )
                    nilai_ppn = (
                        clean_ppn(get_cell_value(ws, r, col_selisih_ws))
                        or clean_ppn(get_cell_value(ws, r, col_acc_ws))
                        or 0.0
                    )

                if sheet_name == "Analisis JV":
                    cancel_per_supplier["JV"] = (
                        cancel_per_supplier.get("JV", 0.0) + nilai_ppn
                    )
                if nama:
                    cancel_per_supplier[nama] = (
                        cancel_per_supplier.get(nama, 0.0) + nilai_ppn
                    )

            elif "PENDING" in status_upper:
                fill_to_apply = yellow_fill

            if fill_to_apply:
                if MODE_WARNA == "HANYA_STATUS":
                    cell_status.fill = fill_to_apply
                elif MODE_WARNA == "BARIS_DATA":
                    for c in range(1, col_status + 1):
                        ws.cell(row=r, column=c).fill = fill_to_apply

    print(
        f"--> Selesai sheet '{sheet_name}': {cocok_count} baris berhasil dicocokkan."
    )


def proses_summary_total(wb, cancel_per_supplier, alias_dict):
    sheet_name = "Summary Total"
    if sheet_name not in wb.sheetnames:
        print(f"[Lewati] Sheet '{sheet_name}' tidak ditemukan di {FILE_HASIL}.")
        return

    ws = wb[sheet_name]
    print(f"--> Memproses pembaruan pada sheet: {sheet_name}...")

    col_nama = (
        temukan_kolom_ws(
            ws, ["Nama Penjual", "Nama Pemasok", "Nama Lawan Transaksi"]
        )
        or 1
    )
    col_selisih = temukan_kolom_ws(ws, ["Selisih", "Selisih PPN"]) or 4

    col_cancel = temukan_kolom_ws(ws, ["CANCEL"])
    col_selisih_cancel = temukan_kolom_ws(
        ws, ["SELISIH CANCEL", "SELISIH_CANCEL"]
    )

    if not col_cancel:
        col_cancel = col_selisih + 1
    if not col_selisih_cancel:
        col_selisih_cancel = col_cancel + 1

    letter_selisih = get_column_letter(col_selisih)
    letter_cancel = get_column_letter(col_cancel)
    letter_selisih_cancel = get_column_letter(col_selisih_cancel)

    ref_header = ws.cell(row=1, column=col_selisih)
    for col_idx, title in [
        (col_cancel, "CANCEL"),
        (col_selisih_cancel, "SELISIH CANCEL"),
    ]:
        c = ws.cell(row=1, column=col_idx, value=title)
        if ref_header.font:
            c.font = copy(ref_header.font)
        c.alignment = Alignment(horizontal="center", vertical="center")
        if ref_header.border:
            c.border = copy(ref_header.border)
        if ref_header.fill and ref_header.fill.fill_type:
            c.fill = copy(ref_header.fill)

    pink_fill = PatternFill(
        start_color="F8CECC", end_color="F8CECC", fill_type="solid"
    )
    no_fill = PatternFill(fill_type=None)

    ref_data = ws.cell(row=2, column=col_selisih)
    default_num_format = (
        ref_data.number_format
        if (
            ref_data
            and ref_data.number_format
            and ref_data.number_format != "General"
        )
        else '#,##0;(#,##0);"-"'
    )

    baris_cancel_count = 0

    for r in range(2, ws.max_row + 1):
        nama_raw = get_cell_value(ws, r, col_nama)
        if not nama_raw:
            continue

        nama_str = str(nama_raw).strip().upper()
        if "TOTAL" in nama_str:
            ref_cell = ws.cell(row=r, column=col_selisih)
            if ref_cell.border:
                ws.cell(row=r, column=col_cancel).border = copy(ref_cell.border)
                ws.cell(row=r, column=col_selisih_cancel).border = copy(
                    ref_cell.border
                )
            continue

        nama_standar = standarkan_nama(nama_raw, alias_dict)

        cell_cancel = ws.cell(row=r, column=col_cancel)
        cell_sel_cancel = ws.cell(row=r, column=col_selisih_cancel)

        ref_cell = ws.cell(row=r, column=col_selisih)
        if ref_cell.border:
            cell_cancel.border = copy(ref_cell.border)
            cell_sel_cancel.border = copy(ref_cell.border)

        num_fmt = (
            ref_cell.number_format
            if (
                ref_cell.number_format and ref_cell.number_format != "General"
            )
            else default_num_format
        )
        cell_cancel.number_format = num_fmt
        cell_sel_cancel.number_format = num_fmt

        if ref_cell.font:
            cell_cancel.font = copy(ref_cell.font)
            cell_sel_cancel.font = copy(ref_cell.font)

        cell_cancel.alignment = Alignment(horizontal="right", vertical="center")
        cell_sel_cancel.alignment = Alignment(
            horizontal="right", vertical="center"
        )

        nilai_cancel = cancel_per_supplier.get(nama_standar, 0.0)
        if nilai_cancel == 0.0 and clean_str(nama_raw) in cancel_per_supplier:
            nilai_cancel = cancel_per_supplier[clean_str(nama_raw)]

        if nilai_cancel > 0:
            baris_cancel_count += 1
            cell_cancel.value = round(nilai_cancel, 2)
            cell_sel_cancel.value = f"={letter_selisih}{r}-{letter_cancel}{r}"

            for c in range(1, col_selisih + 1):
                ws.cell(row=r, column=c).fill = pink_fill
        else:
            if cell_cancel.value is not None and not str(
                cell_cancel.value
            ).startswith("="):
                cell_cancel.value = None
            if cell_sel_cancel.value is not None:
                cell_sel_cancel.value = None

            first_cell = ws.cell(row=r, column=1)
            if (
                first_cell.fill
                and first_cell.fill.fill_type == "solid"
                and first_cell.fill.start_color.rgb
                and "F8CECC" in str(first_cell.fill.start_color.rgb).upper()
            ):
                for c in range(1, col_selisih + 1):
                    ws.cell(row=r, column=c).fill = no_fill

    ws.column_dimensions[letter_cancel].width = 16
    ws.column_dimensions[letter_selisih_cancel].width = 18

    print(
        f"--> Selesai '{sheet_name}': {baris_cancel_count} baris berstatus CANCEL diisi dan diwarnai."
    )


def jalankan_rekonsiliasi():
    if not os.path.exists(FILE_HASIL):
        raise FileNotFoundError(f"File {FILE_HASIL} tidak ditemukan!")

    alias_dict = muat_alias(FILE_ALIAS)

    wb = openpyxl.load_workbook(FILE_HASIL)
    cancel_per_supplier = {}

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

        ctx_dict, ctx_faktur_dict = muat_dict_coretax(file_patokan, alias_dict)
        ws = wb[sheet_target]
        proses_sheet(
            ws,
            ctx_dict,
            ctx_faktur_dict,
            sheet_target,
            alias_dict,
            cancel_per_supplier,
        )

    proses_summary_total(wb, cancel_per_supplier, alias_dict)

    wb.save(FILE_HASIL)
    print(f"--> Sukses! Seluruh pembaruan disimpan ke {FILE_HASIL}.")


if __name__ == "__main__":
    jalankan_rekonsiliasi()