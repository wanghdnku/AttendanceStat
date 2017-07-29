#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import xlrd


'''
将输入的日期转化为用'/'分隔
'''
def date_formatter(input_date):
    return input_date.replace('-', '/')


'''
从钉钉导出的请假表中检查每个人的请假记录
:param 钉钉导出的请假表的路径
:return {姓名:[(日期, 假期类型+时长)]}
'''
def check_vacation(excel_path='/Users/hayden/Desktop/20170723105500047.xlsx'):

    data = xlrd.open_workbook(excel_path)
    table = data.sheets()[0]
    rows_number = table.nrows

    vacation_list = []

    for index in range(1, rows_number):
        table_row = table.row_values(index)
        if table_row[2] != '已撤销' and table_row[3] != '拒绝':
            vacation_list.append(table_row)

    vacation_dict = dict()

    for record in vacation_list:

        name = record[7]
        types = record[13]
        start = date_formatter(record[14].split(' ')[0])
        end = date_formatter(record[15].split(' ')[0])
        length = record[16]

        if name not in vacation_dict:
            vacation_dict[name] = []

        # 如果只请一天假，那就好办了
        if float(length) <= 1 and start == end:
            vacation_dict[name].append((start, '%s%s天' % (types, length)))
        else:
            start_year, start_month, start_day = start.split('/')
            end_year, end_month, end_day = end.split('/')
            # 须保证假期不跨月
            assert start_year == end_year and start_month == end_month
            start_day, end_day = int(start_day), int(end_day)

            # print('%s %s天 %s - %s' %(name, length, start, end))

            for date in range(start_day, end_day+1):
                current_date = '%s/%s/%s' % (start_year, start_month, date)
                number_of_days = '1'
                if current_date == end:
                    remain = '%.2f' % (float(length) % 1.0)
                    number_of_days = '1' if remain == '0.00' else remain
                vacation_dict[name].append((current_date, '%s%s天' % (types, number_of_days)))

    return vacation_dict

for k, v in check_vacation().items():
    print(k, v)
