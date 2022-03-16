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
        types = typeCell.value.split("、")
        print(types)


def computeTypesCost(types)


def isEmptyCell(cell):
    return cell.value is None or cell.value == ""


init()
