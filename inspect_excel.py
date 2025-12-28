
import openpyxl
import os

files = [
    "kisi_bilgileri.xlsx",
    "d:/YAPAY ZEKALILAR/rapor_sistemi/kisi_bilgileri.xlsx"
]

target = None
for f in files:
    if os.path.exists(f):
        target = f
        break

if not target:
    print("Excel file not found!")
    exit(1)

print(f"Reading: {target}")
wb = openpyxl.load_workbook(target, data_only=True)
sheet = wb.active

for row in sheet.iter_rows(min_row=1, max_row=50, values_only=True):
    print(row)
