#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from statistics import *
from utilities import *

'''
程序主入口
'''
if __name__ == '__main__':

    encoding_format = input('输入编码格式(1: utf-8, 2: gbk): ')

    # 默认使用 GBK 编码
    encoding = 'gbk'

    # 如果输入1, 使用 UTF-8 编码
    if encoding_format == '1':
        encoding = 'utf-8'

    file_path = get_input_path()

    csv_path = get_output_path(file_path, 'statistics.csv')

    workdays = work_calendar(2017, 7)

    # 统计出勤信息
    statistics(file_path, workdays, encoding)

    # 生成缺勤记录
    find_absence(csv_path, workdays, encoding)
