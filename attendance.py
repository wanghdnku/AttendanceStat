import xlrd
import csv
import platform
from datetime import datetime

'''
输入表格：
0:部门 1:姓名 2:考勤号码 3:日期时间 4:机器号 5:编号(空) 6:比对方式 7:卡号

输出表格：
0:部门 1:姓名 2:考勤号码 3:工作日 4:上班时间 5:下班时间 6:比对方式 7:出勤情况 8:工时 

缺勤统计表：
0:部门 1:姓名 2:考勤号码 3:工作日 4:上班时间 5:下班时间 6:比对方式 7:出勤情况 8:工时
'''

DEPT = 0
NAME = 1
NUMB = 2
TIME = 3

'''
将字符串类型时间转化为数字类型时间
:param time_string - 时间字符串，如 9:30:59
:return double类型的时间
'''
def string_to_time(time_string):
    time = time_string.split(':')
    numeric = float(time[0])
    numeric += float(time[1]) / 60
    return numeric


'''
出勤时间检测
:param department - 部门名称，如'技术部' 
:param get_on_time - 上班时间，如'8:59:59'
:param get_off_time - 下班时间，如'18:00:00'
:return 出勤检测结果，字符串类型
'''
def turnout_checking(department, get_on_time, get_off_time):

    result = ''

    # 得到实际上班时间和下班时间
    get_on = string_to_time(get_on_time)
    get_off = string_to_time(get_off_time)

    # 根据部门调整基准上下班时间
    on_duty = 9.0
    off_duty = 18.0
    # 技术部实行弹性工作制，9:30-10:00均可
    if department == '技术部':
        on_duty = 10.0
        # 根据上班时间调整下班时间
        if get_on < 9.5:
            off_duty = 18.5
        elif get_on > 10.0:
            off_duty = 19.0
        else:
            off_duty = get_on + 9.0

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
    if get_off - get_on <= 1.0:
        result = '疑似漏打卡'

    return result


'''
生成工作日的日历，将周末与法定休息日删去
:param input_year - 输入一个年份，整型
:param input_month - 输入一个月份，整型
:return 一个数组，保存了所有工作日，数组元素为字符串格式，如'2017/5/30'
'''
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
        if not set(holidays).issubset(set(range(1, days_num + 1))):
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


'''
统计出勤信息
:param excel_path - 输入Excel文件存储的路径
:return void，将统计结果写入到输入文件同文件夹下的statistics.csv中
Notice: Excel文件必须是.xlsx，否则会出现编码错误
'''
def statistics(excel_path='/Users/hayden/Desktop/checkin.xlsx'):

    # TODO: 加一列计算当天的工时

    # 根据操作系统确定路径分隔符
    path_separator = '\\'
    if platform.system() == 'Darwin':
        path_separator = '/'

    # 读取Excel数据
    data = xlrd.open_workbook(excel_path)

    # 获取第一张Sheet的表格
    table = data.sheets()[0]

    # 计算行数
    rows_number = table.nrows

    # 创建要写入的csv文件，记录统计结果
    csv_path = path_separator.join(excel_path.split(path_separator)[:-1])
    csv_path += path_separator
    csv_path += 'statistics.csv'
    csv_file = open(csv_path, 'w', encoding='gbk')
    writer = csv.writer(csv_file)

    # 写入表头
    header = table.row_values(0)
    header[3] = '工作日'
    header[4] = '上班时间'
    header[5] = '下班时间'
    header[7] = '出勤情况'
    # 再加一行工时
    header.append('工时')
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
        if row[NUMB] == previous_row[NUMB] and row[TIME].split(' ')[0] == previous_row[TIME].split(' ')[0]:
            # 将同一个人同一天的其他打卡记录略过
            pass
        else:
            # 将该行写入
            new_row[5] = previous_row[TIME].split(' ')[1]
            new_row[3] = previous_row[TIME].split(' ')[0]
            new_row[7] = turnout_checking(new_row[DEPT], new_row[4], new_row[5])
            new_row.append(round(string_to_time(new_row[5]) - string_to_time(new_row[4]), 2))
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
    new_row.append(round(string_to_time(new_row[5]) - string_to_time(new_row[4]), 2))
    writer.writerow(new_row)
    # 计入出勤统计
    if new_row[NAME] not in absence:
        absence[new_row[NAME]] = [new_row[TIME]]
    else:
        absence[new_row[NAME]].append(new_row[TIME])

    csv_file.close()

    print('%d月工作日:' % month)
    print(workdays, '\n')

    # 统计最终缺勤信息
    print('缺勤记录')
    for staff in absence:
        absence[staff] = list(set(workdays) - set(absence[staff]))
        if absence[staff]:
            print('%s: 缺勤%d天 %s'
                  % (staff, len(absence[staff]), sorted(absence[staff], key=lambda x: int(x.split('/')[2]))))


'''
程序主入口
'''
if __name__ == '__main__':
    file_path = input('输入文件路径: ')
    statistics(file_path)
