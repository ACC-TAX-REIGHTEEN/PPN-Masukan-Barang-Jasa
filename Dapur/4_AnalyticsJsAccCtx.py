import itertools
from datetime import datetime
import os
import sys
import pandas as pd

INPUT_FILE_CORETAX = "CtxJV_temp.xlsx"
INPUT_FILE_ACCURATE = "AccJV_temp.xlsx"
OUTPUT_FILE = "Hasil_Analisis_JV_temp.xlsx"
TOLERANCE = 50

BULAN_MAP = {
    "Jan": "01",
    "Feb": "02",
    "Mar": "03",
    "Apr": "04",
    "Mei": "05",
    "Jun": "06",
    "Jul": "07",
    "Agu": "08",
    "Sep": "09",
    "Okt": "10",
    "Nov": "11",
    "Des": "12",
    "Oktober": "10",
    "Agustus": "08",
    "Nop": "11",
}


def parse_indo_date(date_val):
    if pd.isna(date_val):
        return pd.NaT

    date_str = str(date_val)
    if isinstance(date_val, datetime):
        return date_val

    try:
        parts = date_str.split()
        if len(parts) >= 3:
            day = parts[0].zfill(2)
            month_str = parts[1][:3]
            month = BULAN_MAP.get(month_str, "00")
            year = parts[2]
            return pd.to_datetime(f"{year}-{month}-{day}")
    except Exception:
        pass
    return pd.to_datetime(date_val, errors="coerce")


def fmt_date(d):
    if pd.isna(d):
        return ""
    try:
        return d.strftime("%d %b %Y")
    except Exception:
        return str(d)


def extract_column(df, name_candidates, index_candidates):
    for name in name_candidates:
        if name in df.columns:
            return df[name]

    for col in df.columns:
        for name in name_candidates:
            if name.lower() in str(col).lower():
                return df[col]

    for idx in index_candidates:
        if 0 <= idx < len(df.columns):
            return df.iloc[:, idx]

    return pd.Series([None] * len(df))


def main():
    print("--> Membaca file Excel...")
    try:
        df_coretax = pd.read_excel(INPUT_FILE_CORETAX)
        df_accurate = pd.read_excel(INPUT_FILE_ACCURATE)
    except Exception as e:
        print(f"--> Gagal membaca file: {e}")
        return

    print("--> Membersihkan data dengan Fallback...")

    s_nama_ct = extract_column(
        df_coretax,
        name_candidates=[
            "Nama Penjual",
            "Nama Penjual Barang Kena Pajak/Barang Kena Pajak Tidak Berwujud/Jasa Kena Pajak",
            "Nama",
            "Pemasok",
        ],
        index_candidates=[1, 0], 
    )

    s_tgl_ct = extract_column(
        df_coretax,
        name_candidates=[
            "Tanggal Faktur Pajak",
            "Faktur Pajak/Dokumen Tertentu/Nota Retur/Nota Pembatalan - Tanggal",
            "Tanggal",
            "Tgl",
        ],
        index_candidates=[3],
    )

    s_ppn_ct = extract_column(
        df_coretax,
        name_candidates=["PPN", "PPN (Rupiah)", "Jumlah PPN", "Nilai PPN"],
        index_candidates=[11, 6],  
    )

    ct_clean = pd.DataFrame()
    ct_clean["Nama"] = s_nama_ct.fillna("UNKNOWN")
    ct_clean["Tanggal"] = s_tgl_ct.apply(parse_indo_date)
    ct_clean["PPN"] = pd.to_numeric(s_ppn_ct, errors="coerce").fillna(0)
    ct_clean["ID"] = ct_clean.index

    s_tgl_ac = extract_column(
        df_accurate,
        name_candidates=["Tgl. Pajak", "Tanggal Pajak", "Tanggal", "Tgl"],
        index_candidates=[1],
    )

    s_ref_ac = extract_column(
        df_accurate,
        name_candidates=["No. Faktur Pajak", "Ref", "Nomor Ref", "Keterangan"],
        index_candidates=[2],
    )

    s_ppn_ac = extract_column(
        df_accurate,
        name_candidates=["Jumlah Pajak", "PPN", "Nilai PPN", "Jumlah"],
        index_candidates=[3],
    )

    ac_clean = pd.DataFrame()
    ac_clean["Tanggal_Raw"] = s_tgl_ac
    ac_clean["Tanggal"] = ac_clean["Tanggal_Raw"].apply(parse_indo_date)
    ac_clean["Ref"] = s_ref_ac.fillna("")
    ac_clean["PPN"] = pd.to_numeric(s_ppn_ac, errors="coerce").fillna(0)
    ac_clean["ID"] = ac_clean.index

    print("--> Mencari angka saling hapus (offset)...")
    ac_pool = ac_clean.to_dict("records")
    used_ac = set()
    offset_pairs = []

    val_map = {}
    for item in ac_pool:
        v = item["PPN"]
        if v == 0:
            continue
        if v not in val_map:
            val_map[v] = []
        val_map[v].append(item["ID"])

    for val in list(val_map.keys()):
        if val > 0:
            neg_val = -val
            if neg_val in val_map:
                pos_ids = val_map[val]
                neg_ids = val_map[neg_val]
                while pos_ids and neg_ids:
                    p_id = pos_ids.pop(0)
                    n_id = neg_ids.pop(0)
                    offset_pairs.append((p_id, n_id))
                    used_ac.add(p_id)
                    used_ac.add(n_id)

    print("--> Melakukan pencocokan data (Smart Matching)...")
    results = []
    matched_ct = set()

    curr_ac_pool = [
        x for x in ac_pool if x["ID"] not in used_ac and x["PPN"] != 0
    ]

    targets = ct_clean.to_dict("records")
    vendor_groups = {}

    for t in targets:
        if t["PPN"] == 0:
            continue
        v = t["Nama"]
        if v not in vendor_groups:
            vendor_groups[v] = []
        vendor_groups[v].append(t)

    for vendor, items in vendor_groups.items():
        unmatched_items = [x for x in items if x["ID"] not in matched_ct]
        if not unmatched_items:
            continue

        total_vendor = sum(x["PPN"] for x in unmatched_items)

        match_ac = None
        for ac in curr_ac_pool:
            if abs(ac["PPN"] - total_vendor) <= TOLERANCE:
                match_ac = ac
                break

        if match_ac:
            results.append(
                {
                    "CT_Items": unmatched_items,
                    "AC_Items": [match_ac],
                    "Selisih": total_vendor - match_ac["PPN"],
                    "Ket": "Matched (Vendor Group)",
                }
            )
            for x in unmatched_items:
                matched_ct.add(x["ID"])
            used_ac.add(match_ac["ID"])
            curr_ac_pool.remove(match_ac)

    unmatched_targets = [
        t for t in targets if t["ID"] not in matched_ct and t["PPN"] != 0
    ]
    for t in unmatched_targets:
        match_ac = None
        best_diff = TOLERANCE + 1
        for ac in curr_ac_pool:
            diff = abs(ac["PPN"] - t["PPN"])
            if diff <= TOLERANCE and diff < best_diff:
                best_diff = diff
                match_ac = ac
                if diff == 0:
                    break
        if match_ac:
            results.append(
                {
                    "CT_Items": [t],
                    "AC_Items": [match_ac],
                    "Selisih": t["PPN"] - match_ac["PPN"],
                    "Ket": "Matched (Single)",
                }
            )
            matched_ct.add(t["ID"])
            used_ac.add(match_ac["ID"])
            curr_ac_pool.remove(match_ac)

    unmatched_targets = [
        t for t in targets if t["ID"] not in matched_ct and t["PPN"] != 0
    ]
    for t in unmatched_targets:
        t_val = t["PPN"]
        t_date = t["Tanggal"]
        filtered_pool = []
        for p in curr_ac_pool:
            if (
                pd.isna(p["Tanggal"])
                or pd.isna(t_date)
                or abs((p["Tanggal"] - t_date).days) <= 60
            ):
                filtered_pool.append(p)
        if len(filtered_pool) > 40:
            filtered_pool = filtered_pool[:40]

        found_combo = None
        for r in range(2, 5):
            for combo in itertools.combinations(filtered_pool, r):
                s = sum(c["PPN"] for c in combo)
                if abs(s - t_val) <= TOLERANCE:
                    found_combo = list(combo)
                    break
            if found_combo:
                break
        if found_combo:
            s_combo = sum(c["PPN"] for c in found_combo)
            results.append(
                {
                    "CT_Items": [t],
                    "AC_Items": found_combo,
                    "Selisih": t_val - s_combo,
                    "Ket": "Matched (Combination)",
                }
            )
            matched_ct.add(t["ID"])
            for item in found_combo:
                used_ac.add(item["ID"])
                curr_ac_pool = [
                    x for x in curr_ac_pool if x["ID"] != item["ID"]
                ]

    print("--> Menyusun laporan akhir...")
    final_rows = []

    for res in results:
        ct_items = res["CT_Items"]
        ac_items = res["AC_Items"]

        dates_c = sorted(list(set([fmt_date(x["Tanggal"]) for x in ct_items])))
        names_c = sorted(list(set([x["Nama"] for x in ct_items])))
        total_c = sum(x["PPN"] for x in ct_items)

        ac_desc = []
        total_ac = 0
        for x in ac_items:
            ref_str = f", Ref: {x['Ref']}" if str(x["Ref"]).strip() != "" else ""
            ac_desc.append(f"{x['PPN']:,.0f} ({fmt_date(x['Tanggal'])}{ref_str})")
            total_ac += x["PPN"]

        final_rows.append(
            {
                "Tanggal Coretax": "\n".join(dates_c),
                "Nama Pemasok": "\n".join(names_c),
                "PPN Coretax": total_c,
                "Rincian PPN Accurate": "\n".join(ac_desc),
                "Total PPN Accurate": total_ac,
                "Selisih": res["Selisih"],
                "Keterangan": res["Ket"],
            }
        )

    unmatched_final_ct = [
        t for t in targets if t["ID"] not in matched_ct and t["PPN"] != 0
    ]
    for t in unmatched_final_ct:
        final_rows.append(
            {
                "Tanggal Coretax": fmt_date(t["Tanggal"]),
                "Nama Pemasok": t["Nama"],
                "PPN Coretax": t["PPN"],
                "Rincian PPN Accurate": "",
                "Total PPN Accurate": 0,
                "Selisih": t["PPN"],
                "Keterangan": "Unmatched Coretax",
            }
        )

    unmatched_final_ac = [
        x
        for x in ac_pool
        if x["ID"] not in used_ac
        and x["ID"] not in [o[0] for o in offset_pairs]
        and x["ID"] not in [o[1] for o in offset_pairs]
    ]
    for x in unmatched_final_ac:
        ref_str = f", Ref: {x['Ref']}" if str(x["Ref"]).strip() != "" else ""
        final_rows.append(
            {
                "Tanggal Coretax": "",
                "Nama Pemasok": "",
                "PPN Coretax": 0,
                "Rincian PPN Accurate": f"{x['PPN']:,.0f} ({fmt_date(x['Tanggal'])}{ref_str})",
                "Total PPN Accurate": x["PPN"],
                "Selisih": -x["PPN"],
                "Keterangan": "Unmatched Accurate",
            }
        )

    for p_id, n_id in offset_pairs:
        item1 = ac_clean.loc[ac_clean["ID"] == p_id].iloc[0]
        item2 = ac_clean.loc[ac_clean["ID"] == n_id].iloc[0]

        ref1 = f", Ref: {item1['Ref']}" if str(item1["Ref"]).strip() != "" else ""
        ref2 = f", Ref: {item2['Ref']}" if str(item2["Ref"]).strip() != "" else ""

        desc = f"{item1['PPN']:,.0f} ({fmt_date(item1['Tanggal'])}{ref1})\n{item2['PPN']:,.0f} ({fmt_date(item2['Tanggal'])}{ref2})"

        final_rows.append(
            {
                "Tanggal Coretax": "",
                "Nama Pemasok": "",
                "PPN Coretax": 0,
                "Rincian PPN Accurate": desc,
                "Total PPN Accurate": item1["PPN"] + item2["PPN"],
                "Selisih": 0,
                "Keterangan": "Offset Pair (Zeroed)",
            }
        )

    df_final = pd.DataFrame(final_rows)
    df_final["PPN Coretax"] = pd.to_numeric(df_final["PPN Coretax"])
    df_final["Total PPN Accurate"] = pd.to_numeric(df_final["Total PPN Accurate"])
    df_final["Selisih"] = pd.to_numeric(df_final["Selisih"])

    grand_total_row = pd.DataFrame(
        [
            {
                "Tanggal Coretax": "GRAND TOTAL",
                "Nama Pemasok": "",
                "PPN Coretax": df_final["PPN Coretax"].sum(),
                "Rincian PPN Accurate": "",
                "Total PPN Accurate": df_final["Total PPN Accurate"].sum(),
                "Selisih": df_final["Selisih"].sum(),
                "Keterangan": "",
            }
        ]
    )

    df_final = pd.concat([df_final, grand_total_row], ignore_index=True)

    print(f"--> Menyimpan ke {OUTPUT_FILE}...")
    with pd.ExcelWriter(OUTPUT_FILE, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Analisis")
        wb = writer.book
        ws = writer.sheets["Analisis"]

        fmt_num = wb.add_format({"num_format": "#,##0"})
        fmt_wrap = wb.add_format({"text_wrap": True, "valign": "top"})

        ws.set_column("A:A", 15, fmt_wrap)
        ws.set_column("B:B", 30, fmt_wrap)
        ws.set_column("C:C", 15, fmt_num)
        ws.set_column("D:D", 40, fmt_wrap)
        ws.set_column("E:F", 15, fmt_num)
        ws.set_column("G:G", 20, fmt_wrap)

    print("--> Selesai!")


if __name__ == "__main__":
    main()