from pyrfc import Connection, ABAPApplicationError, ABAPRuntimeError, LogonError, CommunicationError


def generateInput(groups):
    sales = []
    material = []
    if isinstance(groups[0][0], list):
        for i in range(len(groups)):
            for j in range(len(groups[i])):
                sales.append({'LOW':str(groups[i][j][0])})
                material.append({'LOW':str(groups[i][j][1])})
    else:
        for i in range(len(groups)):
            sales.append({'LOW':str(groups[i][0])})
            material.append({'LOW':str(groups[i][1])})
    # print(sales)
    # print(material)
    salesorg = [{'LOW':'IN01'} for i in range(len(sales))]
    return [salesorg, sales, material]

def updateReason(salesorg, sales, material, reason):
    try:
        conn = Connection(user='INBHP002', ashost='CADapp05.esc.win.colpal.com', sysnr='05', client='321', passwd='Bdp@251299')
        print(conn.alive)
        salesm = [
        {'LOW':'101204797'},
        {'LOW':'101204797'}, 
        # {'LOW':'101074101'}
        ]
        materialm = [
        {'LOW':'1601032'}, 
        {'LOW':'1608876'},
        # {'LOW':'1107710'}
        ]
        salestype = [
            {'LOW': 'ZIOR'},
            {'LOW': 'ZIOR'}, 
            # {'LOW': 'ZNOR'}
        ]
        orderdate = [
            {'LOW':'01/01/2000'},
            {'LOW':'01/01/2000'}, 
            # {'LOW':'06/02/2021'}
        ]
        crdddate = [
            {'LOW':'01/01/2000', 'HIGH': '12/31/9999'},
            {'LOW':'01/01/2000', 'HIGH': '12/31/9999'}, 
            # {'LOW':'02/01/2004', 'HIGH': '02/29/2004'}
        ]
        fm_dict = {
        'function_name':'ZFM_BWRFR02',
        'connection': conn,
        'import_args': {
            'IT_SALESORG': salesorg,
            # 'IT_SALESTYP': salestype,
            # 'IT_ORDR_CRE_DATE': orderdate,
            # 'IT_CRDDDATE': crdddate,
            # 'IT_SALESDOC': salesm,
            # 'IT_MATERIAL': materialm,
            'IT_SALESDOC': sales,
            'IT_MATERIAL': material,
            # 'ACTUAL_REASON': '33',
            'ACTUAL_REASON': reason
        }
    }
        result = fm_dict['connection'].call(fm_dict['function_name'], **fm_dict['import_args'])
        # print(result)
        conn.close()

    except CommunicationError:
        print(u"Could not connect to server.")
        raise
    except LogonError:
        print(u"Could not log in. Wrong credentials?")
        raise
    except (ABAPApplicationError, ABAPRuntimeError):
        print(u"An error occurred.")
        raise
    finally:
        conn.close()
    return result['IT_RETURN']