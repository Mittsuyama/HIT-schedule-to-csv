import xlrd
from datetime import datetime, timedelta
import re

schedule_normal = [
    ("08:00 AM", "09:45 AM"),  # 1,2
    ("10:00 AM", "11:45 AM"),  # 3,4
    ("01:45 PM", "03:30 PM"),  # 5,6
    ("03:45 PM", "05:30 PM"),  # 7,8
    ("06:30 PM", "08:15 PM"),  # 9,10
    ("08:30 PM", "10:15 PM"),  # 11,12
]

schedule_exam = [
    ("08:00 AM", "10:00 AM"),  # 1,2
    ("10:00 AM", "12:00 PM"),  # 3,4
    ("01:00 PM", "03:00 PM"),  # 5,6
    ("03:45 PM", "05:45 PM"),  # 7,8
    ("06:30 PM", "08:30 PM"),  # 9,10
]

schedule_experiment = [
    ("07:20 AM", "09:50 AM"),  # 1,2
    ("10:00 AM", "12:30 PM"),  # 3,4
    ("01:00 PM", "03:30 PM"),  # 5,6
    ("03:40 PM", "06:10 PM"),  # 7,8
    ("06:30 PM", "09:00 PM"),  # 9,10
]

csv_value = ""


def writeCSV():
    keys = [
        "Subject",
        "Start Date",
        "Start Time",
        "End Date",
        "End Time",
        "All Day Event",
        "Description",
        "Location",
        "Private",
    ]
    csv = open("canlendar.csv", "w")
    csv.write(','.join(keys) + '\n')
    csv.write(csv_value)
    csv.close()


def dealCellValue(s_date, x, y, value):
    global csv_value
    if len(value) < 2:
        return
    subjects = value.split('\n')
    for index in range(0, int(len(subjects) / 2)):
        name = subjects[index * 2]
        information = subjects[index * 2 + 1]
        infos = information.split('周')
        teachers = []
        times = []
        for item in infos:
            if not '[' in item:
                break
            result = re.search(r'，?(.+)\[(.+)\]', item, re.M | re.I)
            teachers.append(result.group(1))
            times.append(result.group(2))
        for index1 in range(0, len(teachers)):
            s_time = times[index1]
            teacher = teachers[index1]
            time_list = s_time.split('，')
            for item in time_list:
                time_range = item.split('-')
                startWeek = int(time_range[0])
                endWeek = 0
                flag = 2
                if len(time_range) == 1:
                    endWeek = startWeek
                elif '单' in time_range[1]:
                    endWeek = int(time_range[1][:-1])
                    flag = 0
                elif '双' in time_range[1]:
                    endWeek = int(time_range[1][:-1])
                    flag = 1
                else:
                    endWeek = int(time_range[1])
                while startWeek <= endWeek:
                    if startWeek % 2 != flag:
                        _date = (s_date +
                                 timedelta(days=(startWeek - 1) * 7 +
                                           y)).strftime("%Y-%m-%d")
                        _start = schedule_normal[x][0]
                        _end = schedule_normal[x][1]
                        result = [name, _date, _start, _date,
                                  _end, 'False', teacher, infos[-1], 'True']
                        csv_value += (','.join(result) + '\n')
                    startWeek += 1


def main():
    start_day = input("输入本学期开始日期，参考格式：2020-2-17：")
    # start_day = "2020-2-24"
    s_date = datetime.strptime(start_day, "%Y-%m-%d")
    data = xlrd.open_workbook("kb.xls")
    table = data.sheets()[0]
    for x in range(2, 8):
        for y in range(2, 9):
            dealCellValue(s_date, x - 2, y - 2, table.cell_value(x, y))
    writeCSV()


if __name__ == "__main__":
    main()
