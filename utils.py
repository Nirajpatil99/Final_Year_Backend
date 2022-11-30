from openpyxl.workbook import Workbook as openpyxlWorkbook
import pandas as pd
from openpyxl import load_workbook
import xlrd
import json
import os


def convertToXLSX(orgFileName, newFileName):
    xlsBook = xlrd.open_workbook(orgFileName)
    workbook = openpyxlWorkbook()
    for i in range(0, xlsBook.nsheets):
        xlsSheet = xlsBook.sheet_by_index(i)
        sheet = workbook.active if i == 0 else workbook.create_sheet()
        sheet.title = xlsSheet.name

        for row in range(4, xlsSheet.nrows):
            for col in range(0, xlsSheet.ncols):
                sheet.cell(row=row - 3, column=col +
                           1).value = xlsSheet.cell_value(row, col)
    workbook.save(newFileName)
    workbook = load_workbook(newFileName)
    return workbook


def generateJsonForDailyProd(workbook):
    res = {}
    for ws in workbook.worksheets:
        prev = ""
        jobOrder = []
        partName = []
        operation = []
        cy_time = []
        part_running = []
        operator_name = []
        actual_prod_time = []
        prod_per_hr = []
        qty_req = []
        qty_achieved = []
        loss_qty = []
        loss_mc_hr = []
        st = []
        programmer_busy = []

        for idx, i in enumerate(ws.rows):
            if (idx == 0):
                continue
            if (i[0].value != None):
                prev = i[0].value
            else:
                i[0].value = prev
            if (i[1].value == None):
                break
            jobOrder.append(i[1].value)
            partName.append(i[2].value)
            operation.append(i[3].value)
            cy_time.append(i[4].value)
            part_running.append(i[5].value)
            operator_name.append(i[6].value)
            actual_prod_time.append(i[7].value)
            prod_per_hr.append(i[8].value)
            qty_req.append(i[9].value)
            qty_achieved.append(i[10].value)
            loss_qty.append(i[11].value)
            loss_mc_hr.append(i[12].value)
            st.append(i[13].value)
            programmer_busy.append(i[14].value)

        combined = {"joborders": jobOrder, "partname": partName,
                    "operation": operation, "cy_time": cy_time, "part_running": part_running, "operator_name": operator_name,
                    "actual_prod_time": actual_prod_time, "prod_per_hr": prod_per_hr, "qty_req": qty_req, "qty_achieved": qty_achieved,
                    "loss_qty": loss_qty, "loss_mc_hr": loss_mc_hr, "st": st, "programmer_busy": programmer_busy}

        df = pd.DataFrame.from_dict(combined)

        # print(df[df.qty_achieved == df.qty_achieved.max()]["qty_achieved"])
        # print(df.dtypes)

        df['actual_prod_time'] = df['actual_prod_time'].astype(float).round(2)
        df['prod_per_hr'] = df['prod_per_hr'].astype(float).round(2)
        df['qty_req'] = df['qty_req'].astype(float).round(2)
        df['loss_qty'] = df['loss_qty'].astype(float).round(2)
        df['loss_mc_hr'] = df['loss_mc_hr'].astype(float).round(2)
        df['st'] = df['st'].astype(float).round(2)
        df['programmer_busy'] = df['programmer_busy'].astype(float).round(2)
        df['qty_achieved'] = df['qty_achieved'].astype(float).round(2)

        # df.round(2)

        # df.to_json("test.json", orient="records")
        partGrp = df.groupby(['partname'])
        res[ws.title] = {}
        rows = []
        for key in list(partGrp.groups.keys()):
            row = partGrp.get_group(
                key).groupby(['partname']).sum().to_dict(orient="records")[0]
            row['partname'] = key
            rows.append(row)

        majors = {}

        majors["max_lossqty"] = df[["partname", "loss_qty"]].sort_values(
            by="loss_qty", ascending=False).head(5).to_dict(orient="records")

        majors["max_actual_prod_time"] = df[["partname", "actual_prod_time"]].sort_values(
            by="actual_prod_time", ascending=False).head(5).to_dict(orient="records")

        majors["max_prod_per_hr"] = df[["partname", "prod_per_hr"]].sort_values(
            by="prod_per_hr", ascending=False).head(5).to_dict(orient="records")

        majors["least_prod_per_hr"] = df[["partname", "prod_per_hr"]].sort_values(
            by="prod_per_hr", ascending=True).head(5).to_dict(orient="records")

        majors["max_loss_mc_hr"] = df[["partname", "loss_mc_hr"]].sort_values(
            by="loss_mc_hr", ascending=False).head(5).to_dict(orient="records")

        majors["least_loss_mc_hr"] = df[["partname", "loss_mc_hr"]].sort_values(
            by="loss_mc_hr", ascending=True).head(5).to_dict(orient="records")

        majors["max_qty_req"] = df[["partname", "qty_req"]].sort_values(
            by="qty_req", ascending=False).head(5).to_dict(orient="records")

        majors["max_qty_achieved"] = df[["partname", "qty_achieved"]].sort_values(
            by="qty_achieved", ascending=False).head(5).to_dict(orient="records")

        res[ws.title].update({"rows": rows, "majors": majors})

        # print(res)
        # res = res

    # print(json.dumps(res))
    return res


def handleDailyProd(filename):
    workBook = convertToXLSX(filename, filename.replace("xls", "xlsx"))
    # print(workBook.worksheets)
    # return {}
    return generateJsonForDailyProd(workBook)


def handleMisProd(filename):
    workBook = load_workbook(filename, data_only=True)
    return {'FC - Operations - In House Prod': handleMisProdInHouseProd(
        workBook['FC - Operations - In House Prod']), "FC - Machine Shop": handleMisProdMachineShop(workBook["FC - Machine Shop"])}


def handleMisProdMachineShop(ws):
    dic = {}
    mainH = ws.cell(10, 2).value
    dic[mainH] = {}
    # row = Row , col = Col
    for row in range(6, 1000, 5):
        datee = ws.cell(4, row)
        if (datee.value == None):
            break
        datee = datee.value.strftime("%b-%Y")
        dic[mainH][datee] = {}
        for col in range(11, 16):
            key = ws.cell(col, 2).value.strip()
            dic[mainH][datee][key] = []
            for k in range(0, 4):
                valToAppend = ws.cell(col, row+k).value
                dic[mainH][datee][key].append(
                    valToAppend if valToAppend != None else 0)

    mainH = ws.cell(17, 2).value
    dic[mainH] = {}
    # row = Row , col = Col
    for row in range(6, 1000, 5):
        datee = ws.cell(4, row)
        if (datee.value == None):
            break
        datee = datee.value.strftime("%b-%Y")
        dic[mainH][datee] = {}
        for col in range(18, 23):
            key = ws.cell(col, 2).value.strip()
            dic[mainH][datee][key] = []
            for k in range(0, 4):
                valToAppend = ws.cell(col, row+k).value
                dic[mainH][datee][key].append(
                    valToAppend if valToAppend != None else 0)

    lol = []
    for key, value in dic.items():
        # temp = dict()
        # temp['detail'] = key
        for key1, value1 in value.items():
            # temp['month'] = key1
            for key2, value2 in value1.items():
                # print(value2[0])
                # temp['subdetail'] = key2
                # temp['fitting'] = value2[0]
                # temp['valves'] = value2[1]
                # temp['clamp'] = value2[2]
                lol.append({'detail': key, 'month': key1,
                            'subdetail': key2, 'fitting': value2[0], 'valves': value2[1], 'clamp': value2[2]})
        temp = dict()

    # print(pd.DataFrame.from_records(lol).to_json(orient="records"))
    return dic


def handleMisProdInHouseProd(ws):
    dic = {}
    mainH = ws.cell(8, 2).value
    dic[mainH] = {}
    # row = Row , col = Col
    for row in range(6, 1000, 5):
        datee = ws.cell(4, row)
        if (datee.value == None):
            break
        datee = datee.value.strftime("%b-%Y")
        dic[mainH][datee] = {"total": []}
        total = [0, 0, 0, 0]

        for col in range(9, 14):
            dic[mainH][datee][ws.cell(col, 2).value] = []
            for k in range(0, 4):
                valToAppend = ws.cell(col, row+k).value
                dic[mainH][datee][ws.cell(col, 2).value].append(
                    valToAppend if valToAppend != None else 0)
                total[k] = total[k] + \
                    (valToAppend if valToAppend != None else 0)
        dic[mainH][datee]["total"] = total

    mainH = ws.cell(15, 2).value
    dic[mainH] = {}
    # row = Row , col = Col
    for row in range(6, 1000, 5):
        datee = ws.cell(4, row)
        if (datee.value == None):
            break
        datee = datee.value.strftime("%b-%Y")
        dic[mainH][datee] = {"total": []}
        total = [0, 0, 0, 0]
        for col in range(17, 20):
            dic[mainH][datee][ws.cell(col, 2).value] = []
            for k in range(0, 4):
                valToAppend = ws.cell(col, row+k).value
                dic[mainH][datee][ws.cell(col, 2).value].append(
                    valToAppend if valToAppend != None else 0)
                total[k] = total[k] + \
                    (valToAppend if valToAppend != None else 0)
        dic[mainH][datee]["total"] = total

    mainH = ws.cell(21, 2).value
    dic[mainH] = {}
    # row = Row , col = Col
    for row in range(6, 1000, 5):
        datee = ws.cell(4, row)
        if (datee.value == None):
            break
        datee = datee.value.strftime("%b-%Y")
        dic[mainH][datee] = {"total": []}
        total = [0, 0, 0, 0]
        for col in range(23, 29):
            dic[mainH][datee][ws.cell(col, 2).value] = []
            for k in range(0, 4):
                valToAppend = ws.cell(col, row+k).value
                dic[mainH][datee][ws.cell(col, 2).value].append(
                    valToAppend if valToAppend != None else 0)
                total[k] = total[k] + \
                    (valToAppend if valToAppend != None else 0)
        dic[mainH][datee]["total"] = total

    mainH = ws.cell(30, 2).value
    dic[mainH] = {}
    # row = Row , col = Col
    for row in range(6, 1000, 5):
        datee = ws.cell(4, row)
        if (datee.value == None):
            break
        datee = datee.value.strftime("%b-%Y")
        dic[mainH][datee] = {"total": []}
        total = [0, 0, 0, 0]
        for col in range(33, 36):
            dic[mainH][datee][ws.cell(col, 3).value] = []
            for k in range(0, 4):
                valToAppend = ws.cell(col, row+k).value
                dic[mainH][datee][ws.cell(col, 3).value].append(
                    valToAppend if valToAppend != None else 0)
                total[k] = total[k] + \
                    (valToAppend if valToAppend != None else 0)
        dic[mainH][datee]["total"] = total

    lol = []
    for key, value in dic.items():
        # temp = dict()
        # temp['detail'] = key
        for key1, value1 in value.items():
            # temp['month'] = key1
            for key2, value2 in value1.items():
                # print(value2[0])
                # temp['subdetail'] = key2
                # temp['fitting'] = value2[0]
                # temp['valves'] = value2[1]
                # temp['clamp'] = value2[2]
                lol.append({'detail': key, 'month': key1,
                            'subdetail': key2, 'fitting': value2[0], 'valves': value2[1], 'clamp': value2[2]})
    return dic


def deleteConvertedXLS(filename):
    if (os.path.exists(filename)):
        os.remove(filename)
    filename = filename.replace("xls", "xlsx")
    if (os.path.exists(filename)):
        os.remove(filename)
    pass