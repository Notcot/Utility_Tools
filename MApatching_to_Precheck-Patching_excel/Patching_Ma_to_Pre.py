import openpyxl


MApatching = input("Enter name of Master Patching List: ")
precheck = input("Enter name of Pre-check File: ")

output_file = "Precheck_Patching_List.xlsx"



wb_ma = openpyxl.load_workbook(MApatching)
wb_precheck = openpyxl.load_workbook(precheck)

sheet_ma = wb_ma.worksheets[0]
sheet_pre = wb_precheck.worksheets[0]

print(f"Sheet from Master Patching: {sheet_ma.title}")
print(f"Sheet from Pre-check File: {sheet_pre.title}")


ma_data = {}
print("Reading Master Patching..")
'''
for row in range(3, sheet_ma.max_row + 1):
    ref_num = sheet_ma.cell(row = row, column = 1).value
    if ref_num is None:
        continue

    row_data = []
    for col in range(2, sheet_ma.max_column + 1):
        row_data.append(sheet_ma.cell(row=row, column = col).value)

    ma_data[ref_num] = row_data
'''

track = 0
for row in sheet_ma.iter_rows(min_row = 3, values_only=True):
    ref_num = row[0]
    if ref_num is None:
        continue

    row_data = list(row[1:])
    ma_data[ref_num] = row_data

    track += 1
    if track % 2000 == 0:
        print(f"Processed {track} rows..")


wb_result = openpyxl.Workbook()
result_sheet = wb_result.active
result_sheet.title = "Patching List"


header1 = []
for col in range(1, sheet_ma.max_column + 1):
    header1.append(sheet_ma.cell(row=1, column = col).value)

result_sheet.append(header1)

header2 = []
for col in range(1,sheet_ma.max_column +1):
    header2.append(sheet_ma.cell(row=2, column = col).value)

result_sheet.append(header2)
print("Header copied from Master Patching")


for row in range(3, sheet_pre.max_row + 1):
    ref_num = sheet_pre.cell(row=row, column = 1).value
    if ref_num is None:
        continue

    if ref_num in ma_data:
        entry = [ref_num] + ma_data[ref_num]
        result_sheet.append(entry)


wb_result.save(output_file)
print(f"Done! Result saved as: {output_file}")

wb_ma.close()
wb_precheck.close()