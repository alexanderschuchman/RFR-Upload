import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook

wb1 = load_workbook("samples/RFR Sample.xlsx")
ws1 = wb1.worksheets[0]

print(ws1.cell(row = 1, column = 1).value)