import xlrd
import xlsxwriter
import xlwt
import json
import requests
import openpyxl

from Excel.install import i

wb = openpyxl.load_workbook('/Applications/办公自动化/Excel/2019年4月销售订单.xlsx')
ws = wb.active
# ws.title = "测试"
# ws.sheet_properties.tabColor = "1072BA"
sheets = wb.sheetnames
# wa = wb[sheets[0]]
#
# print(wa)
# sheets = wb.sheetnames
# for sheet in sheets:
#     print(sheets)
# for sheet in wb:
#     print
#     sheet.title
# cole=ws['C']
# print(cole)
# for col in ws.iter_cols(min_row=1,max_col=3,max_row=2):
#     for cell in col:
#         print(cell)
cellValue = ws.cell(row=1, column=1).value
print(cellValue)