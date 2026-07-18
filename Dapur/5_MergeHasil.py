from copy import copy
import sys
import openpyxl


def main():
    try:
        wb_jv = openpyxl.load_workbook("Hasil_Analisis_JV_temp.xlsx")
        wb_barang = openpyxl.load_workbook("Hasil_Analisis_Barang_temp.xlsx")
    except Exception as e:
        print(f"--> Gagal membuka file: {e}")
        sys.exit()

    ws_src = wb_jv["Analisis"]
    ws_dest = wb_barang.create_sheet(title="Analisis JV")

    for row in ws_src.rows:
        for cell in row:
            new_cell = ws_dest.cell(
                row=cell.row, column=cell.column, value=cell.value
            )
            if cell.has_style:
                new_cell.font = copy(cell.font)
                new_cell.border = copy(cell.border)
                new_cell.fill = copy(cell.fill)
                new_cell.number_format = copy(cell.number_format)
                new_cell.protection = copy(cell.protection)
                new_cell.alignment = copy(cell.alignment)

        for merged_cell in ws_src.merged_cells.ranges:
            ws_dest.merge_cells(str(merged_cell))

    for col_name, col_dim in ws_src.column_dimensions.items():
        ws_dest.column_dimensions[col_name].width = col_dim.width

    for row_name, row_dim in ws_src.row_dimensions.items():
        ws_dest.row_dimensions[row_name].height = row_dim.height

    try:
        wb_barang.save("Hasil_Analisis_Barang_Dan_Jasa.xlsx")
        print("--> Berhasil menggabungkan file dan menyimpannya...")
    except Exception as e:
        print(f"--> Gagal menyimpan file: {e}")
    finally:
        wb_jv.close()
        wb_barang.close()


if __name__ == "__main__":
    main()
