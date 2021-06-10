import xlrd
import xlsxwriter
import requests
import pandas
file = ("D:\\Infogistic\\Python worksheets\\Test.xls")
workbook = xlrd.open_workbook(file)
sheet = workbook.sheet_by_index(0)
rows = sheet.nrows
cols = sheet.ncols
print(rows)
print(cols)

all_rows = []
for row in range(sheet.nrows):
    curr_row = []
    for col in range(sheet.ncols):
        curr_row.append(sheet.cellvalue(row, col))
    all_rows.append(curr_row)


for row in range(1, rows):
    url = sheet.cell_value(row, 0)
print(url)
response = requests.get(url)
if response.status_code == 200:
    all_rows[1][5] == 'Pass'
new_path = ("D:\\Infogistic\\Python worksheets\\Result.xlsx")
new_workbook = xlsxwriter.workbook(new_path)
new_worksheet = new_workbook.add_worksheet()
for row in range(len(all_rows)):
    for col in range(len(all_rows[0])):
        new_worksheet.write(row, col, all_rows[rows][cols])