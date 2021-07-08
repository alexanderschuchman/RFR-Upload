from celery import shared_task
import time
from celery_progress.backend import ProgressRecorder
from .scripts.segmentation import sortFile, generateGroups
from .scripts.input import generateInput, updateReason

@shared_task(bind=True, name="rfrlogic", ignore_result=False)
def logic(self):
    progress_recorder = ProgressRecorder(self)
    print('Started')
    sortFile()
    progress_recorder.set_progress(1, 3, description="Input File Sorted")
    groups = generateGroups()
    # print(groups['Discontinued/ Obselete'])
    # print(groups['Master data issue'])
    # print(groups)
    # reason dict
    reasons = {
        'Unreasonable request':'10',
        'Not valid for credit/debit memo':'33',
        'Order qty< Min order quantity': '35',
        'Discontinued/ Obsolete Items': '37',
        'Discontinued/ Obsolete': '37',
        'Item has been rounded off': '40',
        'Customer related': 'C0',
        'Customer cancel Line/Order': 'C1',
        'Pricing does not match PO': 'C4',
        'Customer Rejected CP Price (SPI)': 'C5',
        # 'No Stocks Available': '34',
        'Customers cannot receive': 'C2',
        'Expired PO (CRRD)': 'C6',
        'RDC Replenishment Delay': 'D3',
        'Systems Hardware relat': 'H0',
        'Manufacturing - related': 'M0',
        'Production Capacity Cons': 'M3',
        'Copacking Delay': 'M4',
        'Orders processing rela': 'O0',
        'Picking error': 'O1',
        'Order entry/processing': 'O3',
        'Sales - related': 'S0',
        'Sales Related': 'S0',
        'Duplicate order entry': 'S2',
        'Sales Cancel-min deliver': 'S6',
        'Sales Cancel line-clear': 'S7',
        'Insufficient lead time-c': 'S8',
        'Stocks > 6 months': 'ST',
        'Stocks on QA status': 'W1',
        'Wrong inventory': 'W2',
        'Distribution - wrong loc': 'W3',
        'Backorders not accepted': 'Z1',
        'Credit Held Purge': 'Z2',
        'Customer order Qty < 1 c': 'Z7',
        'PGI Qty Less than Order Qty (Quota)': 'EX',
        'Manufacturing defect/poor quality': 'M1',
        'Production/Transfer delays': 'M2',
        'Wrong material code used': 'O2',
        'Wrong SKU code': 'O2',
        'Exceeds Quota Quantity': 'O4',
        'No Stocks Available': 'O5',
        'Change the Old Product Hier. in SKU': 'PH',
        'Over Customer allocation': 'S1',
        'Sales Related': 'S3',
        'Under Forecast': 'S4',
        'Advanced Selling': 'S5',
        'Advance Selling': 'S5',
        'Warehouse - related': 'W0',
        'In excess of truckload': 'W4',
        'EDI - Dummy SKU': 'Z3',
        'EDI - Price Check': 'Z4',
        'EDI - Wrong UOM': 'Z5',
        'APO-CMI Dummy Sales Order': 'Z6'
    }
    progress_recorder.set_progress(2, 3, description="Input File Fragmented")
    count = 0
    failedcount = 0
    string = ""
    for k in groups.keys():
        res = generateInput(groups[k])
        salesorg = res[0]
        sales = res[1]
        material = res[2]
        # reason = reasons[k]
        reason = k
        print(reason)
        ret = updateReason(salesorg, sales, material, reason)
        print(ret)
        if isinstance(ret, list):
            print(len(ret))
            for x in range(len(ret)):
                if ret[x]['TYPE']!='E':
                    count += 1
                else:
                    failedcount += 1
                    # print(ret[x])
                    string += ret[x]['MESSAGE_V1'][1:] + " - " + ret[x]['MESSAGE_V2'] + ", "
        else:
            if ret['TYPE']!='E':
                count += 1
            else:
                failedcount += 1
                string += ret['MESSAGE_V1'][1:] + " - " + ret['MESSAGE_V2'] + ", "
    print(string)
    if string!="":
        string = "Updated " + str(count) + " records. Failed " + str(failedcount) +" records: " + string[:-2]
    else:
        string = "Updated " + str(count) + " records."
    progress_recorder.set_progress(3, 3, description="")
    return string