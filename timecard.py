# coding=utf-8
from openpyxl import load_workbook
import sys

valid_depart = {
    "产品组": 1,
    "电力四川战略网省": 1,
    "电力线损分析部": 1,
    "电力项目管理与实施部": 1,
    "电力研发部": 1,
    "电力终端产品部": 1,
    "平台架构部": 1,
    "设计部": 1,
    "研究院": 1
}

projects_info = {}


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
    time_info, projects = processAllTime(time_sheet, salary_records)
    # print(projects)
    # print(time_info)
    result, prjDepart, departTimes, departCost = processSalary(
        time_info, salary_records, working_days, initial, projects)
    # print(departCost)
    update(time_sheet, result, initial, prjDepart, departTimes, departCost)
    summaryTime(projects, time_info, time_sheet, departTimes, departCost)
    verify(time_info, salary_records, result, working_days, departCost)
    print("process done!")


def isEmptyCell(cell):
    return cell.value is None or cell.value == ""


def processAllTime(wb, salary):
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

        if name not in salary:
            #print("name not in salary", name)
            continue

        depart = salary[name]["depart"]
        if depart not in valid_depart:
            print("name not in valid depart", name, depart)
            continue

        projectIdCell = timeRecordsSheet.cell(row=i, column=8)
        if isEmptyCell(projectIdCell):
            print("empty project cell:", i, name)
            continue

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

        projectNameCell = timeRecordsSheet.cell(row=i, column=9)
        if projectIdCell.value not in projects_info:
            projects_info[projectIdCell.value] = projectNameCell.value

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
    prjDepart = {}
    departTimes = {}
    departCostRecords = {}
    # go through time info
    for name in timeInfo:
        if name not in salaryInfo:
            #print("can not find salary:", name)
            continue
        salary = salaryInfo[name]
        totalHours = timeInfo[name]["total_hours"]
        departHours = 0
        depart = salary["depart"]

        departCost = initial.copy()
        if totalHours < standardHours:
            departHours = standardHours - totalHours

            if depart not in departCostRecords:
                departCostRecords[depart] = initial.copy()

            for key in departCostRecords[depart]:
                departCost[key] = salary[key] * (departHours / standardHours)
                departCostRecords[depart][key] += departCost[key]

            if depart not in departTimes:
                departTimes[depart] = {}
            if name not in departTimes[depart]:
                departTimes[depart][name] = 0
            departTimes[depart][name] += departHours

        for project in timeInfo[name]["projects"]:
            if project not in result:
                result[project] = initial.copy()
            for key in result[project]:
                prjCost = (salary[key] - departCost[key]) * \
                    (timeInfo[name]["projects"][project] / totalHours)
                result[project][key] += prjCost
                if project not in prjDepart:
                    prjDepart[project] = {}
                if depart not in prjDepart[project]:
                    prjDepart[project][depart] = {}
                if key not in prjDepart[project][depart]:
                    prjDepart[project][depart][key] = 0.0
                prjDepart[project][depart][key] += prjCost
    # print(departCostRecords)
    return (result, prjDepart, departTimes, departCostRecords)
    # print(result)


def update(wb, records, initial, prjDepart, departTimes, departCostRecords):
    # write to excel
    sheetName = "time_cost"
    if sheetName in wb.sheetnames:
        del wb[sheetName]
    targetSheet = wb.create_sheet(sheetName)
    titles = ["项目代号", "项目名称", "部门"] + list(initial.keys())
    # print(titles)
    targetSheet.append(titles)

    for project in records:
        row = [project]
        if project in projects_info:
            row.append(projects_info[project])
        else:
            row.append(None)
        row.append(None)
        # for i in range(3, len(titles)):
        #    row.append(round(records[project][titles[i]], 2))
        targetSheet.append(row)
        # append depart detail
        if project in prjDepart:
            departCost = prjDepart[project]
            for depart in departCost:
                detailRow = [None, None, depart]
                for j in range(3, len(titles)):
                    detailRow.append(round(departCost[depart][titles[j]], 2))
                targetSheet.append(detailRow)

    for prj in departCostRecords:
        print("project:", prj)
        row = [prj]
        if prj in projects_info:
            row.append(projects_info[prj])
        else:
            row.append(None)
        row.append(None)
        for i in range(3, len(titles)):
            row.append(round(departCostRecords[prj][titles[i]], 2))
        targetSheet.append(row)

    wb.save("output.xlsx")


def setupSalary(wb):
    salary = {}
    salarySheet = wb[wb.sheetnames[0]]
    print("Salary info length:", salarySheet.max_row)
    for i in range(2, salarySheet.max_row + 1):
        nameCell = salarySheet.cell(row=i, column=1)
        if isEmptyCell(nameCell):
            print("salary empty name cell:", i)
            continue
        print("process:", nameCell.value)
        departCell = salarySheet.cell(row=i, column=6)
        if isEmptyCell(departCell):
            print("salary empty depart cell:", i)
            continue
        salary[nameCell.value] = {
            "depart": departCell.value,
            "应发工资": float(salarySheet.cell(row=i, column=9).value),
            "养老保险_个人": float(salarySheet.cell(row=i, column=10).value),
            "医疗保险_个人": float(salarySheet.cell(row=i, column=11).value),
            "失业保险_个人": float(salarySheet.cell(row=i, column=12).value),
            "大病医疗保险_个人": float(salarySheet.cell(row=i, column=13).value),
            "住房公积金_个人": float(salarySheet.cell(row=i, column=14).value),
            "养老保险_公司": float(salarySheet.cell(row=i, column=16).value),
            "医疗保险_公司": float(salarySheet.cell(row=i, column=17).value),
            "失业保险_公司": float(salarySheet.cell(row=i, column=18).value),
            "工伤保险_公司": float(salarySheet.cell(row=i, column=19).value),
            "生育保险_公司": float(salarySheet.cell(row=i, column=20).value),
            "大病医疗保险_公司": float(salarySheet.cell(row=i, column=21).value),
            "住房公积金_公司": float(salarySheet.cell(row=i, column=22).value),
            "所得税": float(salarySheet.cell(row=i, column=24).value),
            "实发工资": float(salarySheet.cell(row=i, column=25).value),
        }

    # print(salary)
    return salary


def summaryTime(prjInfo, timeInfo, wb, departTimes, departCostRecords):
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

    for project in departTimes:
        row = [project]
        prj = departTimes[project]
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


def verify(timeInfo, salary, processResult, workDays, departCost):
    print("start verify....")
    sTotal = 0.0
    eCount = 0
    for name in timeInfo:
        if name in salary:
            sTotal += salary[name]["应发工资"]
            eCount += 1
    print("included employee salary total:", sTotal, eCount)

    pTotal = 0
    for prj in processResult:
        pTotal += processResult[prj]["应发工资"]

    pDepartTotal = 0
    for prj in departCost:
        pDepartTotal += departCost[prj]["应发工资"]

    print("total project spend:", pTotal + pDepartTotal, pTotal, pDepartTotal)

    # cross verify project cost compute

    verifyResult = {}
    standardHours = workDays * 8
    for project in processResult:
        pTotal = 0.0
        for name in timeInfo:
            if project in timeInfo[name]["projects"]:
                if name not in verifyResult:
                    verifyResult[name] = {
                        "salary": salary[name]["应发工资"],
                        "total_hours": timeInfo[name]["total_hours"],
                        "project_cost": 0.0,
                        "depart_cost": 0.0,
                        "project_detail": {}
                    }
                prjHours = timeInfo[name]["projects"][project]
                totalHours = timeInfo[name]["total_hours"]
                if totalHours < standardHours:
                    totalHours = standardHours
                prjRatio = prjHours / totalHours
                prjCostShare = salary[name]["应发工资"] * prjRatio
                pTotal += prjCostShare
                verifyResult[name]["project_cost"] += prjCostShare
                verifyResult[name]["project_detail"][project] = {
                    "hours": prjHours,
                    "cost": prjCostShare
                }

        processResult[project]["verified_cost"] = pTotal
        #print("verify:", project, processResult[project]["应发工资"], pTotal)


init()
