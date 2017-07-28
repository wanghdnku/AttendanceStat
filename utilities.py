import platform
from datetime import datetime


SEPARATOR = '/' if platform.system() == 'Darwin' else '\\'


'''
获取文件路径，删除路径前后的特殊字符
'''
def get_input_path():
    return input('输入文件路径: ').strip('$ ')


'''
根据输入文件路径，输出所在文件夹的路径
'''
def get_folder_path(file_path):
    file_path = file_path.strip('$ ')
    return SEPARATOR.join(file_path.split(SEPARATOR)[:-1]) + SEPARATOR


'''
生成输出文件路径，指定输出文件名
'''
def get_output_path(input_path, file_name):
    folder = get_folder_path(input_path)
    return folder + file_name


'''
根据输入的文件，自动检测文件名中的月份、开始日期、结束日期
'''
def get_start_end(file_path):
    file_path = file_path.strip('$ ')
    file_name = file_path.split(SEPARATOR)[-1]
    start_date, end_date = file_name[:-5].split('-')
    assert start_date.split('.')[0] == end_date.split('.')[0]
    month = int(start_date.split('.')[0])
    start = int(start_date.split('.')[1])
    end = int(end_date.split('.')[1])

    return month, start, end


'''
将字符串类型时间转化为数字类型时间。
:param string_time - 时间字符串，如 '9:30:59'。
:return double类型的时间，如 9.5。
'''
def string_to_time(string_time):
    time = string_time.split(':')
    numeric_time = float(time[0])
    numeric_time += float(time[1]) / 60
    return numeric_time


'''
判断各部门人员应该的上班和下班时间。
:param department - 部门名称，如 '技术部'。
:param get_on - 上班时间，用来动态判断应该的下班时间，适用于技术部。
:return 应该的上班和下班时间, double类型。 
'''
def schedule_for_department(department, get_on):

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

    return on_duty, off_duty


'''
出勤情况检测
:param department - 部门名称，如 '技术部'。 
:param get_on_time - 上班时间，如 '8:59:59'。
:param get_off_time - 下班时间，如 '18:00:00'。
:return 出勤检测结果，字符串类型，如 '正常'、'早退'、'旷工'等。
'''
def turnout_checking(department, get_on_time, get_off_time):

    result = ''

    # 得到实际上班时间和下班时间
    get_on = string_to_time(get_on_time)
    get_off = string_to_time(get_off_time)

    # 根据部门和上班时间，计算出上班和下班的时间
    on_duty, off_duty = schedule_for_department(department, get_on)

    # 检查上班时间
    if on_duty < get_on <= (on_duty + 0.5):
        result += '迟到'
    elif get_on > on_duty + 0.5:
        result += '旷工'

    # 检查下班时间
    if (off_duty - 0.5) <= get_off < off_duty:
        result += '早退'
    elif get_off < off_duty - 0.5:
        result += '旷工'

    # 输出结果
    if result == '':
        result = '正常'

    # 疑似漏打卡的情况
    if get_off - get_on <= 1.0:
        result = '疑似漏打卡'

    return result

'''
:param input_month - 月份
:param input_year - 年份
:return 一个月的天数，int类型
'''
def number_of_days(input_month, input_year):
    if input_month in [1, 3, 5, 7, 8, 10, 12]:
        days_num = 31
    elif input_month in [4, 6, 9, 11]:
        days_num = 30
    else:
        if input_year % 4 == 0 and (input_year % 100 != 0 or input_year % 400 == 0):
            days_num = 29
        else:
            days_num = 28

    return days_num


'''
检查某一天是星期几，每周从周一开始
:param input_date - 输入的日期。如 '2017/7/28'
:return int类型，1-7
'''
def days_in_week(input_date):
    return datetime.strptime(input_date, '%Y/%m/%d').isoweekday()


'''
生成工作日的日历，将周末与法定休息日删去
:param input_year - 输入一个年份，整型
:param input_month - 输入一个月份，整型
:return 一个数组，保存了所有工作日，数组元素为字符串格式，如 '2017/5/30'
'''
def work_calendar(input_year, input_month):
    days = []

    # 得出每月的天数
    days_num = number_of_days(input_month, input_year)

    # 设置可以输入起止时间
    while True:
        date_period = input('输入起止时间 (如: 19-23): ')
        start_date, end_date = map(lambda x: int(x), date_period.split('-'))
        if start_date < 1 or end_date > days_num:
            print('输入日期超出了当月范围! (1 ~ %d)' % days_num)
            continue
        break

    for day in range(start_date, end_date + 1):
        days.append('%d/%d/%d' % (input_year, input_month, day))

    # 刨除法定节假日
    while True:
        holidays = input('输入法定节日，空格隔开 (若没有，输入N): ')
        if holidays in 'Nn':
            break
        holidays = holidays.split(' ')
        holidays = list(map(lambda x: int(x), holidays))

        # 检查输入的合法性
        if not set(holidays).issubset(set(range(start_date, end_date + 1))):
            print('请检查输入的合法性! (%d ~ %d)' % (start_date, end_date))
            continue

        # 从日历列表中移除法定假日
        for holiday in holidays:
            days.remove('%d/%d/%d' % (input_year, input_month, holiday))
        break

    # 刨除周末
    days = [day for day in days if days_in_week(day) < 6]

    # 添加法定补班的周末
    while True:
        work_weekends = input('输入补班周末，空格隔开 (若没有，输入N): ')
        if work_weekends in 'Nn':
            break
        work_weekends = work_weekends.split(' ')
        work_weekends = list(map(lambda x: int(x), work_weekends))

        # 检查输入的合法性
        if not set(work_weekends).issubset(set(range(start_date, end_date + 1))):
            print('请检查输入的合法性! (%d ~ %d)' % (start_date, end_date))
            continue

        for work_weekend in work_weekends:
            days.append('%d/%d/%d' % (input_year, input_month, work_weekend))
        break

    return days
