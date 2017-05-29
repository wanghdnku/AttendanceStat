import xlrd
import csv
from datetime import datetime

DEPT = 0
NAME = 1
TIME = 3


def string_to_time(time_string):
    time = time_string.split(':')
    numeric = float(time[0])
    numeric += float(time[1]) / 60
    return numeric


def turnout_checking(department, get_on_time, get_off_time):
    on_duty = 9.0
    off_duty = 18.0
    if department == '技术部':
        on_duty = 10.0
        off_duty = 19.0

    result = ''

    get_on = string_to_time(get_on_time)
    get_off = string_to_time(get_off_time)
    if on_duty < get_on <= (on_duty + 0.5):
        result += '迟到'
    elif get_on > on_duty + 0.5:
        result += '旷工'

    if (off_duty - 0.5) <= get_off < off_duty:
        result += '早退'
    elif get_off < off_duty - 0.5:
        result += '旷工'

    if result == '':
        result = '正常'

    # 疑似漏打卡的情况
    if get_on_time == get_off_time:
        result = '疑似漏打卡'

    return result


def work_calendar(input_year, input_month):
    days = []

    # 生成日历
    if input_month in [1, 3, 5, 7, 8, 10, 12]:
        days_num = 31
    elif input_month in [4, 6, 9, 11]:
        days_num = 30
    else:
        if input_year % 4 == 0 and (input_year % 100 != 0 or input_year % 400 == 0):
            days_num = 29
        else:
            days_num = 28
    for day in range(1, days_num + 1):
        days.append('%d/%d/%d' % (input_year, input_month, day))

    # 刨除法定节假日
    while True:
        holidays = input('输入法定节日，空格隔开 (若没有，输入N): ')
        if holidays in 'Nn':
            break
        holidays = holidays.split(' ')
        holidays = list(map(lambda x: int(x), holidays))

        # 检查输入的合法性
        if not set(holidays).issubset(set(range(1, days_num+1))):
            print('请检查输入的合法性! (1 ~ %d)' % days_num)
            continue

        for holiday in holidays:
            days.remove('%d/%d/%d' % (input_year, input_month, holiday))
        break

    # 刨除周末
    # days = list(filter(lambda x: datetime.strptime(x, '%Y/%m/%d').isoweekday() < 6, days))
    days = [day for day in days if datetime.strptime(day, '%Y/%m/%d').isoweekday() < 6]

    # 添加法定补班的周末
    while True:
        work_weekends = input('输入补班周末，空格隔开 (若没有，输入N): ')
        if work_weekends in 'Nn':
            break
        work_weekends = work_weekends.split(' ')
        work_weekends = list(map(lambda x: int(x), work_weekends))

        # 检查输入的合法性
        if not set(work_weekends).issubset(set(range(1, days_num + 1))):
            print('请检查输入的合法性! (1 ~ %d)' % days_num)
            continue

        for work_weekend in work_weekends:
            days.append('%d/%d/%d' % (input_year, input_month, work_weekend))
        break

    return days


if __name__ == '__main__':

    # 读取Excel数据
    data = xlrd.open_workbook('/Users/hayden/Desktop/checkin.xlsx')

    # 获取第一张Sheet的表格
    table = data.sheets()[0]

    # 计算行数
    rows_number = table.nrows

    # 创建要写入的csv文件，记录统计结果
    csv_path = '/Users/hayden/Desktop/new_new.csv'
    csv_file = open(csv_path, 'w', encoding='gbk')
    writer = csv.writer(csv_file)

    # 写入表头
    header = table.row_values(0)
    header[3] = '工作日'
    header[4] = '上班时间'
    header[5] = '下班时间'
    header[7] = '出勤情况'
    writer.writerow(header)

    previous_row = table.row_values(1)

    # 获取表格的年份、月份
    year = int(previous_row[TIME].split('/')[0])
    month = int(previous_row[TIME].split('/')[1])
    workdays = work_calendar(year, month)

    # 记录第一行
    new_row = previous_row
    new_row[4] = previous_row[TIME].split(' ')[1]

    # 记录出勤情况的字典，键是员工名字，值是一个保存日期的列表
    absence = dict()

    # 遍历每行
    for index in range(2, rows_number):
        row = table.row_values(index)
        if row[NAME] == previous_row[NAME] and row[TIME].split(' ')[0] == previous_row[TIME].split(' ')[0]:
            pass
        else:
            # 将该行写入
            new_row[5] = previous_row[TIME].split(' ')[1]
            new_row[3] = previous_row[TIME].split(' ')[0]
            new_row[7] = turnout_checking(new_row[DEPT], new_row[4], new_row[5])
            writer.writerow(new_row)
            # 计入出勤统计
            if new_row[NAME] not in absence:
                absence[new_row[NAME]] = [new_row[TIME]]
            else:
                absence[new_row[NAME]].append(new_row[TIME])

            # 缓存下一行
            new_row = row
            new_row[4] = row[TIME].split(' ')[1]

        previous_row = row

    # 将最后一行写入
    new_row[5] = previous_row[TIME].split(' ')[1]
    new_row[3] = previous_row[TIME].split(' ')[0]
    new_row[7] = turnout_checking(new_row[DEPT], new_row[4], new_row[5])
    writer.writerow(new_row)
    # 计入出勤统计
    if new_row[NAME] not in absence:
        absence[new_row[NAME]] = [new_row[TIME]]
    else:
        absence[new_row[NAME]].append(new_row[TIME])

    csv_file.close()

    print(workdays)

    # 统计最终缺勤信息
    print('缺勤记录')
    for staff in absence:
        absence[staff] = list(set(workdays) - set(absence[staff]))
        if absence[staff]:
            print('%s: 缺勤%d天 %s' % (staff, len(absence[staff]), sorted(absence[staff], key=lambda x: int(x.split('/')[2]))))
