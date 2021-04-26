import openpyxl
from openpyxl import Workbook
from openpyxl import load_workbook
import win32com.client as win32

def sortFile():
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
        
        ws.Range('A2:R'+str(LastRow)).Sort(Key1=ws.Range('R1'), Order1=1, Orientation=1)
        
        wb.SaveAs(Filename="C:\\Users\\bp_bhagyashree phadn\\Downloads\\RFR-Upload\\samples\\sorted.xlsx")
        # wb.Save()
        excel.Application.Quit()
    except Exception as e:
        print(e)

def generateGroups():
    wb1 = load_workbook("samples/sorted.xlsx")
    ws1 = wb1.worksheets[0]
    groups = {}
    i = 2
    val = ws1.cell(row = 2, column = 18).value
    print(ws1.max_row)
    while i <= ws1.max_row:
        while ws1.cell(row = i, column = 18).value == val:
            if val not in groups.keys():
                groups[val] = [ws1.cell(row = i, column = 3).value]
                # print(groups)
                # print(i)
            elif ws1.cell(row = i, column = 3).value not in groups[val]:
                groups[val].append(ws1.cell(row = i, column = 3).value)
                # print(groups)
                # print(i)
            i += 1
        if ws1.cell(row = i, column = 18).value != val:
            val = ws1.cell(row = i, column = 18).value
            i += 1
    # print(groups)
    # groups['Discontinued/ Obselete'] += groups['Discontinued/ Obselete']
    # groups['Discontinued/ Obselete'] += groups['Discontinued/ Obselete']
    # groups['Discontinued/ Obselete'] += groups['Discontinued/ Obselete']
    for k in groups.keys():
        if len(groups[k])>200:
            groups[k] = [groups[k][:200], groups[k][200:]]
            while len(groups[k][-1])>200:
                excess = groups[k][-1][200:]
                groups[k][-1] = groups[k][-1][:200]
                groups[k].append(excess)
    print(groups)
    return groups

sortFile()
groups = generateGroups()