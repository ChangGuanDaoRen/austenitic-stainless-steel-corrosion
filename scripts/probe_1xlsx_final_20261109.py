"""只读复核最终 1.xlsx 的维度与关键列，避免 pwsh 内联 Python 引号被吞。"""
import openpyxl

wb = openpyxl.load_workbook("1.xlsx", read_only=True)
ws = wb["Sheet1"]
print("sheets", wb.sheetnames)
print("dims", ws.max_row, ws.max_column)
rows = list(ws.iter_rows(min_row=1, max_row=3, values_only=True))
print("header1_A", rows[0][0], "header1_Y", rows[0][24], "header1_Z", rows[0][25])
print("row3_A", rows[2][0], "row3_RE", rows[2][19], "row3_Epit", rows[2][22])
print("row3_Y_nonempty", rows[2][24] is not None, "row3_Z_nonempty", rows[2][25] is not None)
