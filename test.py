import csv
from datetime import datetime


def work_calendar(year, month):
    days = []

    # 生成日历
    if month in [1, 3, 5, 7, 8, 10, 12]:
        days_num = 31
    elif month in [4, 6, 9, 11]:
        days_num = 30
    else:
        if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0):
            days_num = 29
        else:
            days_num = 28
    for day in range(1, days_num + 1):
        days.append('%d/%d/%d' % (year, month, day))

    # 刨除节假日
    # TODO: 输入法定节假日
    holidays = input('输入法定节日，空格隔开: ')
    holidays = holidays.split(' ')
    holidays = map(lambda x: int(x), holidays)
    for d in holidays:
        days.remove('%d/%d/%d' % (year, month, d))

    # 刨除周末
    # days = list(filter(lambda x: datetime.strptime(x, '%Y/%m/%d').isoweekday() < 6, days))
    days = [day for day in days if datetime.strptime(day, '%Y/%m/%d').isoweekday() < 6]

    return days


def find_absence():
    workdays = set(work_calendar(2017, 5))
    dic = dict()
    with open('/Users/hayden/Desktop/new_new.csv', 'r', encoding='gbk') as csv_file:
        spamreader = csv.reader(csv_file)
        for row in spamreader:
            if row[0] != '部门':
                if row[1] not in dic:
                    dic[row[1]] = [row[3]]
                else:
                    dic[row[1]].append(row[3])
        #print(dic)

    for staff in dic:
        dic[staff] = list(workdays - set(dic[staff]))

    print('缺勤记录')
    for staff in dic:
        if dic[staff]:
            print('%s: 缺勤%d天 %s' % (staff, len(dic[staff]), sorted(dic[staff], key=lambda x: int(x.split('/')[2]))))

find_absence()
