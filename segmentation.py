import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
try:
    wb = excel.Workbooks("C:/Users/bp_bhagyashree phadn/Downloads/RFR-Upload/samples/RFRSample.xlsx")
except Exception as e:
    try:
        wb = excel.Workbooks.Open("C:/Users/bp_bhagyashree phadn/Downloads/RFR-Upload/samples/RFRSample.xlsx")
    except Exception as e:
        print(e)
try:
    ws = wb.Worksheets('Sheet1')
    
    LastRow = ws.UsedRange.Rows.Count
    
    ws.Range('R2:R'+str(LastRow)).Sort(Key1=ws.Range('R1'), Order1=1, Orientation=1)
    
    wb.SaveAs(Filename="C:\\Users\\bp_bhagyashree phadn\\Downloads\\RFR-Upload\\samples\\sorted.xlsx")
    excel.Application.Quit()
except Exception as e:
    print(e)