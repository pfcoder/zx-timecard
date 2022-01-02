# coding=utf-8
from openpyxl import load_workbook
import sys


def init():
    if len(sys.argv) < 2:
        print("need input working days!!!")
        return
    working_days = int(sys.argv[1])
    print(working_days)

    initial = {
        "应发工资": 0.0,
        "养老保险_个人": 0.0,
        "医疗保险_个人": 0.0,
        "失业保险_个人": 0.0,
        "大病医疗保险_个人": 0.0,
        "住房公积金_个人": 0.0,
        "养老保险_公司": 0.0,
        "医疗保险_公司": 0.0,
        "失业保险_公司": 0.0,
        "工伤保险_公司": 0.0,
        "生育保险_公司": 0.0,
        "大病医疗保险_公司": 0.0,
        "住房公积金_公司": 0.0,
        "所得税": 0.0,
        "实发工资": 0.0
    }

    print("working days:", working_days)
    time_sheet = load_workbook("./项目工时统计表.xlsx")
    salary_sheet = load_workbook("./s.xlsx")
    salary_records = setupSalary(salary_sheet)
    # go through process time
    time_info, projects = processAllTime(time_sheet)
    # print(projects)
    # print(time_info)
    result = processSalary(
        time_info, salary_records, working_days, initial, projects)
    # print(projects)
    update(time_sheet, result, initial)
    summaryTime(projects, time_info, time_sheet)
    print("process done!")


def isEmptyCell(cell):
    return cell.value is None or cell.value == ""


def processAllTime(wb):
    timeRecordsSheet = wb[wb.sheetnames[0]]
    # name as key
    result = {}
    # project id as key
    projects = {}

    # print(timeRecordsSheet.cell(row=2, column=2).value)
    # go through timeRecordsSheet row by row
    for i in range(2, timeRecordsSheet.max_row + 1):
        nameCell = timeRecordsSheet.cell(row=i, column=2)
        # print("process:", nameCell.value)
        if isEmptyCell(nameCell):
            print("empty name cell:", i)
            continue

        name = nameCell.value
        # print(name)
        if name not in result:
            result[name] = {
                "total_hours": 0,
                "projects": {}
            }

        hoursCell = timeRecordsSheet.cell(row=i, column=12)
        if isEmptyCell(hoursCell):
            print("empty hours cell:", i)
            continue
        result[name]["total_hours"] += int(hoursCell.value)

        projectIdCell = timeRecordsSheet.cell(row=i, column=8)
        if isEmptyCell(projectIdCell):
            print("empty project cell:", i)
            continue

        if projectIdCell.value not in result[name]["projects"]:
            result[name]["projects"][projectIdCell.value] = 0

        result[name]["projects"][projectIdCell.value] += int(hoursCell.value)

        if projectIdCell.value not in projects:
            # employee name as key (should use employee id, TODO)
            projects[projectIdCell.value] = {}

        if name not in projects[projectIdCell.value]:
            projects[projectIdCell.value][name] = 0
        projects[projectIdCell.value][name] += int(hoursCell.value)

    return (result, projects)


def processSalary(timeInfo, salaryInfo, workingDays, initial, projects):
    # for name in timeInfo:
    # project id as key
    result = {}
    standardHours = workingDays * 8
    # go through time info
    for name in timeInfo:
        # if name == "赵海林":
        #    print("process:", name, timeInfo[name])
        if name not in salaryInfo:
            print("can not find salary:", name)
            continue
        salary = salaryInfo[name]
        totalHours = timeInfo[name]["total_hours"]
        departHours = 0

        departCost = initial.copy()
        if totalHours < standardHours:
            departHours = standardHours - totalHours
            depart = salary["depart"]
            if depart not in result:
                result[depart] = initial.copy()

            for key in result[depart]:
                departCost[key] = (salary[key] /
                                   standardHours) * departHours
                result[depart][key] += departCost[key]

            if depart not in projects:
                projects[depart] = {}
            if name not in projects[depart]:
                projects[depart][name] = 0
            projects[depart][name] += departHours

            # if depart == "理论组":
            #    print(departCost)
            #    print(name, timeInfo[name])

        for project in timeInfo[name]["projects"]:
            if project not in result:
                result[project] = initial.copy()
            for key in result[project]:
                result[project][key] += ((salary[key] - departCost[key]) /
                                         totalHours) * timeInfo[name]["projects"][project]
            # if project == "PBUXS-A02-27-2021-001":
            #    print(result[project], totalHours, depart)
    return result
    # print(result)


def update(wb, records, initial):
    # write to excel
    sheetName = "time_cost"
    if sheetName in wb.sheetnames:
        del wb[sheetName]
    targetSheet = wb.create_sheet(sheetName)
    titles = ["项目代号"] + list(initial.keys())
    # print(titles)
    targetSheet.append(titles)

    for project in records:
        row = [project]
        for i in range(1, len(titles)):
            row.append(records[project][titles[i]])
        targetSheet.append(row)

    wb.save("output.xlsx")


def setupSalary(wb):
    salary = {}
    salarySheet = wb[wb.sheetnames[1]]
    for i in range(2, salarySheet.max_row):
        nameCell = salarySheet.cell(row=i, column=1)
        if isEmptyCell(nameCell):
            print("salary empty name cell:", i)
            continue
        # print("process:", nameCell.value)
        departCell = salarySheet.cell(row=i, column=9)
        if isEmptyCell(departCell):
            print("salary empty depart cell:", i)
            continue
        salary[nameCell.value] = {
            "depart": departCell.value,
            "应发工资": float(salarySheet.cell(row=i, column=12).value),
            "养老保险_个人": float(salarySheet.cell(row=i, column=13).value),
            "医疗保险_个人": float(salarySheet.cell(row=i, column=14).value),
            "失业保险_个人": float(salarySheet.cell(row=i, column=15).value),
            "大病医疗保险_个人": float(salarySheet.cell(row=i, column=16).value),
            "住房公积金_个人": float(salarySheet.cell(row=i, column=17).value),
            "养老保险_公司": float(salarySheet.cell(row=i, column=19).value),
            "医疗保险_公司": float(salarySheet.cell(row=i, column=20).value),
            "失业保险_公司": float(salarySheet.cell(row=i, column=21).value),
            "工伤保险_公司": float(salarySheet.cell(row=i, column=22).value),
            "生育保险_公司": float(salarySheet.cell(row=i, column=23).value),
            "大病医疗保险_公司": float(salarySheet.cell(row=i, column=24).value),
            "住房公积金_公司": float(salarySheet.cell(row=i, column=25).value),
            "所得税": float(salarySheet.cell(row=i, column=27).value),
            "实发工资": float(salarySheet.cell(row=i, column=29).value),
        }

    # print(salary)
    return salary


def summaryTime(prjInfo, timeInfo, wb):
    sheetName = "time_summary"
    if sheetName in wb.sheetnames:
        del wb[sheetName]
    targetSheet = wb.create_sheet(sheetName)
    # x is employee names, y is project nos
    # build first title row
    titles = ['项目'] + list(timeInfo.keys()) + ['汇总']
    targetSheet.append(titles)
    total = [0] * len(titles)
    total[0] = '工时汇总'
    for project in prjInfo:
        row = [project]
        prj = prjInfo[project]
        prjTotal = 0
        for i in range(1, len(titles) - 1):
            name = titles[i]
            if name in prj:
                row += [prj[name]]
                total[i] += prj[name]
                prjTotal += prj[name]
            else:
                row += [0]
        row += [prjTotal]
        targetSheet.append(row)
    targetSheet.append(total)

    wb.save("output.xlsx")


init()
