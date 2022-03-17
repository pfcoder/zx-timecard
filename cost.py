# -*- coding: UTF-8 -*-
from openpyxl import load_workbook


def init():
    targetWb = load_workbook("./cost_input.xlsx")
    referWb = load_workbook("./cost_refer.xlsx")
    targetSheet = targetWb[targetWb.sheetnames[0]]
    departPool = set(targetSheet.cell(row=1, column=1).value.split("、"))
    targetColumn = 6
    typeColumn = 4
    for i in range(5, targetSheet.max_row + 1):
        typeCell = targetSheet.cell(row=i, column=typeColumn)
        if isEmptyCell(typeCell):
            continue
        types = typeConvert(typeCell.value.split("、"))
        converts = []
        for item in types:
            if isinstance(item, list):
                converts.extend(item)
            else:
                converts.append(item)
        print("converted:", converts)
        count = computeTypesCost(converts, referWb, departPool)
        #print("process row:", i, count)
        targetSheet.cell(row=i, column=targetColumn).value = round(
            count, 2)

    # write out
    targetWb.save("cost_output.xlsx")


def typeConvert(types):
    return list(map(checkRange, types))


def checkRange(item):
    items = item.split("…")
    if len(items) > 1:
        list = []
        for i in range(0, int(items[1]) - int(items[0]) + 1):
            list.append(str(int(items[0]) + i))
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
    for i in range(0, len(referWb.sheetnames)):
        sheetName = referWb.sheetnames[i]
        if sheetName in config:
            c = config[sheetName]
            print("process sheet:", sheetName)
            count += processSheet(referWb[sheetName], c["typeIndex"], c["departIndex"],
                                  c["amountIndex"], typeSet, departs)

    return count


def processSheet(sheet, typeIndex, departIndex, amountIndex, types, departs):
    count = 0.0
    for i in range(2, sheet.max_row + 1):
        departCell = sheet.cell(row=i, column=departIndex)
        if isEmptyCell(departCell):
            continue
        typeCell = sheet.cell(row=i, column=typeIndex)
        if isEmptyCell(typeCell):
            continue
        amountCell = sheet.cell(row=i, column=amountIndex)
        if isEmptyCell(amountCell):
            continue

        typeValue = typeCell.value.replace("'", "")
        departValue = departCell.value.replace("'", "")
        if departValue in departs and typeValue in types:
            print("found match:", departValue, typeValue, amountCell.value)
            count += float(amountCell.value)

    print("sheet process count:", count)
    return count


def isEmptyCell(cell):
    return cell.value is None or cell.value == ""


init()
