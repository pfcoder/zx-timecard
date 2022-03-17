# -*- coding: UTF-8 -*-
from openpyxl import load_workbook


def init():
    targetWb = load_workbook("./cost_input.xlsx")
    referWb = load_workbook("./cost_refer.xlsx")

    # print(targetWb.sheetnames)

    targetSheet = targetWb[targetWb.sheetnames[0]]
    departPool = targetSheet.cell(row=1, column=1).value
    departPool = set(departPool.split("、"))
    #print("PBU111" in departPool)

    # go through target sheet
    targetColumn = 6
    typeColumn = 4
    for i in range(5, targetSheet.max_row + 1):
        typeCell = targetSheet.cell(row=i, column=typeColumn)
        if isEmptyCell(typeCell):
            continue
        types = typeConvert(typeCell.value.split("、"))
        # print(types)
        converts = []
        for item in types:
            if isinstance(item, list):
                converts.extend(item)
            else:
                converts.append(item)
        #print("converted:", converts)
        count = computeTypesCost(converts, referWb, departPool)
        #print("process row:", i, count)
        targetSheet.cell(row=i, column=targetColumn).value = round(
            count / 10000, 2)

    # write out
    targetWb.save("cost_output.xlsx")


def typeConvert(types):
    return list(map(checkRange, types))


def checkRange(item):
    items = item.split("…")
    #print("checkrange:", items)
    if len(items) > 1:
        list = []
        for i in range(0, int(items[1]) - int(items[0]) + 1):
            list.append(str(int(items[0]) + i))
        #print("convert list:", item, list)
        return list
    else:
        return item


def computeTypesCost(types, referWb, departs):
    config = {
        "收入": {
            "typeIndex": 2,
            "departIndex": 7,
            "amountIndex": 10,
        },
        "生产成本-分项": {
            "typeIndex": 1,
            "departIndex": 4,
            "amountIndex": 5,
        },
        "销售费用分项": {
            "typeIndex": 2,
            "departIndex": 6,
            "amountIndex": 7,
        },
        "管理费用分项": {
            "typeIndex": 2,
            "departIndex": 5,
            "amountIndex": 6,
        },
        "研发费用分项": {
            "typeIndex": 2,
            "departIndex": 5,
            "amountIndex": 6,
        }
    }

    typeSet = set(types)
    count = 0.0
    # gothrough refer sheer
    # check sheet 1:
    # print(referWb.sheetnames)
    for i in range(0, len(referWb.sheetnames)):
        sheetName = referWb.sheetnames[i]
        #print("sheet name:", sheetName)
        if sheetName in config:
            c = config[sheetName]
            print("process sheet:", sheetName)
            count += processSheet(referWb[sheetName], c["typeIndex"], c["departIndex"],
                                  c["amountIndex"], typeSet, departs)

    return count


def processSheet(sheet, typeIndex, departIndex, amountIndex, types, departs):
    count = 0.0
    for i in range(2, sheet.max_row + 1):
        #print(sheet.cell(row=i, column=1))
        departCell = sheet.cell(row=i, column=departIndex)
        if isEmptyCell(departCell):
            #print("depart empty")
            continue
        typeCell = sheet.cell(row=i, column=typeIndex)
        if isEmptyCell(typeCell):
            #print("type empty")
            continue
        amountCell = sheet.cell(row=i, column=amountIndex)
        if isEmptyCell(amountCell):
            #print("amount empty")
            continue

        typeValue = typeCell.value.replace("'", "")
        departValue = departCell.value.replace("'", "")
        #print("process type:", typeValue, departValue, departs, types)
        if departValue in departs and typeValue in types:
            print("found match:", departValue, typeValue, amountCell.value)
            count += float(amountCell.value)

    print("sheet process count:", count)
    return count


def isEmptyCell(cell):
    return cell.value is None or cell.value == ""


init()
