#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
import xlrd
from absence import find_absence
from utilities import *

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
DATE = 3
ON = 4
OFF = 5
HOURS = 6
ATTENDANCE = 7


'''
统计出勤信息
:param excel_path - 输入Excel文件存储的路径
:return void，将统计结果写入到输入文件同文件夹下的statistics.csv中
Notice: Excel文件必须是.xlsx，否则会出现编码错误
'''
def statistics(excel_path='/Users/hayden/Desktop/checkin.xlsx', encoding='gbk'):

    # 读取Excel数据
    data = xlrd.open_workbook(excel_path)
    # 获取第一张工作表
    table = data.sheets()[0]
    # 计算行数
    rows_number = table.nrows

    csv_path = get_output_path(excel_path, 'statistics.csv')

    with open(csv_path, 'w', encoding=encoding) as csv_file:
        writer = csv.writer(csv_file)

        # 写入表头
        header = ['部门', '姓名', '考勤号码', '工作日', '上班时间', '下班时间', '工时', '出勤情况']
        writer.writerow(header)

        previous_row = table.row_values(1)

        # 记录第一行
        new_row = previous_row
        new_row[4] = previous_row[DATE].split(' ')[1]

        # 记录出勤情况的字典，键是员工名字，值是一个保存日期的列表
        absence = dict()

        # 遍历每行
        for index in range(2, rows_number):
            excel_row = table.row_values(index)
            if excel_row[NUMB] == previous_row[NUMB] and excel_row[DATE].split(' ')[0] == previous_row[DATE].split(' ')[0]:
                # 将同一个人同一天的其他打卡记录略过
                pass
            else:
                # 将该行写入
                new_row[5] = previous_row[DATE].split(' ')[1]
                new_row[3] = previous_row[DATE].split(' ')[0]
                new_row[7] = turnout_checking(new_row[DEPT], new_row[4], new_row[5])
                new_row[6] = round(string_to_time(new_row[5]) - string_to_time(new_row[4]), 2)
                if new_row[DEPT] not in '管理层':
                    writer.writerow(new_row)
                    # 计入出勤统计
                    if new_row[NAME] not in absence:
                        absence[new_row[NAME]] = [new_row[DATE]]
                    else:
                        absence[new_row[NAME]].append(new_row[DATE])

                # 缓存下一行
                new_row = excel_row
                new_row[4] = excel_row[DATE].split(' ')[1]

            previous_row = excel_row

        # 将最后一行写入
        new_row[5] = previous_row[DATE].split(' ')[1]
        new_row[3] = previous_row[DATE].split(' ')[0]
        new_row[7] = turnout_checking(new_row[DEPT], new_row[4], new_row[5])
        new_row[6] = round(string_to_time(new_row[5]) - string_to_time(new_row[4]), 2)
        if new_row[DEPT] not in '管理层':
            writer.writerow(new_row)
            # 计入出勤统计
            if new_row[NAME] not in absence:
                absence[new_row[NAME]] = [new_row[DATE]]
            else:
                absence[new_row[NAME]].append(new_row[DATE])

    # 获取表格的年份、月份
    year = int(previous_row[DATE].split('/')[0])
    month = int(previous_row[DATE].split('/')[1])
    workdays = work_calendar(year, month)

    print('%d月工作日:' % month)
    print(workdays, '\n')

    # 统计最终缺勤信息
    print('缺勤记录')
    for staff in absence:
        absence[staff] = list(set(workdays) - set(absence[staff]))
        if absence[staff]:
            print('%s: 缺勤%d天 %s'
                  % (staff, len(absence[staff]), sorted(absence[staff], key=lambda x: int(x.split('/')[2]))))
    print('\n')

    # 检查缺勤记录
    find_absence(csv_path, workdays, encoding)


'''
程序主入口
'''
if __name__ == '__main__':
    file_path = input('输入文件路径: ')
    statistics(file_path)
