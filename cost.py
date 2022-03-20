# -*- coding: UTF-8 -*-
from openpyxl import load_workbook


def init():
    targetWb = load_workbook("./cost_total.xlsx")
    # Jan
    referWb1 = load_workbook("./cost_refer_1.xlsx")
    # Feb
    referWb2 = load_workbook("./cost_refer_2.xlsx")

    targetColumn1 = 6
    targetColumn2 = 8

    processDepart(targetWb["汇总-DBU"], targetColumn1, referWb1)
    processDepart(targetWb["汇总-DBU"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表MBU"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表MBU"], targetColumn2, referWb2)

    processDepart(targetWb["汇总表-PBU"], targetColumn1, referWb1)
    processDepart(targetWb["汇总表-PBU"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-业务拓展部"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-业务拓展部"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-商务拓展部"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-商务拓展部"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-研究院"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-研究院"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-总裁办合并"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-总裁办合并"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-市场部"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-市场部"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-人事行政中心"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-人事行政中心"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-董办"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-董办"], targetColumn2, referWb2)

    processDepart(targetWb["预算汇总表-财务中心+法务"], targetColumn1, referWb1)
    processDepart(targetWb["预算汇总表-财务中心+法务"], targetColumn2, referWb2)

    # write out
    targetWb.save("cost_output.xlsx")


def processDepart(sheet, targetColumn, referWb):
    typeColumn = 4
    departPool = set(sheet.cell(row=1, column=1).value.split("、"))
    for i in range(5, sheet.max_row + 1):
        typeCell = sheet.cell(row=i, column=typeColumn)
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
        sheet.cell(row=i, column=targetColumn).value = round(
            count, 2)


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

        typeValue = str(typeCell.value).replace("'", "")
        departValue = str(departCell.value).replace("'", "")
        if departValue in departs and typeValue in types:
            print("found match:", departValue, typeValue, amountCell.value)
            count += float(amountCell.value)

    print("sheet process count:", count)
    return count


def isEmptyCell(cell):
    return cell.value is None or cell.value == ""


init()
