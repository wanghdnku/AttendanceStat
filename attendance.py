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

    # 刨除周末
    # days = list(filter(lambda x: datetime.strptime(x, '%Y/%m/%d').isoweekday() < 6, days))
    days = [day for day in days if datetime.strptime(day, '%Y/%m/%d').isoweekday() < 6]

    # 刨除节假日
    # TODO: 输入法定节假日

    return days


if __name__ == '__main__':

    # 读取Excel数据
    data = xlrd.open_workbook('/Users/hayden/Desktop/checkin.xlsx')

    # 获取第一张Sheet的表格
    table = data.sheets()[0]

    # 计算行数
    rows_number = table.nrows

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
    #year = previous_row[TIME].split('/')[0]
    #month = previous_row[TIME].split('/')[1]
    #workdays = work_calendar(year, month)

    # 记录第一行
    new_row = previous_row
    new_row[4] = previous_row[TIME].split(' ')[1]

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

            # 缓存下一行
            new_row = row
            new_row[4] = row[TIME].split(' ')[1]

        previous_row = row

    # 将最后一行写入
    new_row[5] = previous_row[TIME].split(' ')[1]
    new_row[3] = previous_row[TIME].split(' ')[0]
    new_row[7] = turnout_checking(new_row[DEPT], new_row[4], new_row[5])
    writer.writerow(new_row)

    csv_file.close()
