#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from statistics import *
from remove_records import *
from utilities import *

'''
程序主入口
'''
if __name__ == '__main__':

    # 默认使用 GBK 编码
    encoding = 'gbk'

    # 获取输入文件位置
    file_path = get_input_path()

    # 计算出输出文件位置
    statistics_path = get_output_path(file_path, 'statistics.csv')
    absence_path = get_output_path(file_path, 'absence.csv')

    # 从文件名中获取统计表的起止日期
    month, start, end = get_start_end(file_path)

    # 根据起止日期生成工作日列表
    workdays = work_calendar_for_period(2017, month, start, end)

    # 统计出勤信息
    statistics(file_path, workdays, encoding)

    # 生成缺勤记录
    find_absence(statistics_path, workdays, encoding)

    # 删除上周四的统计
    start_date = '2017/%s/%s' % (month, start)
    if end - start == 7 and days_in_week(start_date):
        print('\n%s 记录已删除' % start_date)
        remove_records(statistics_path, start_date)
        remove_records(absence_path, start_date)
