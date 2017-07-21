#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv
from utilities import get_output_path

# FIXME: 将输入区间少设一天，来解决此问题？

def remove_records(file_path, remove_date, chosen_encode='gbk'):

    records = []

    with open(file_path, 'r', encoding=chosen_encode) as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            if row[3] != remove_date:
                records.append(row)

    with open(file_path, 'w', encoding=chosen_encode) as csv_file:
        writer = csv.writer(csv_file)
        for row in records:
            writer.writerow(row)

if __name__ == '__main__':

    encoding_format = input('输入编码格式(1: utf-8, 2: gbk): ')
    # 默认使用 GBK 编码
    encoding = 'gbk'
    # 如果输入1, 使用 UTF-8 编码
    if encoding_format == '1':
        encoding = 'utf-8'

    path = input('请输入文件夹路径: ')
    statistics_path = path + 'statistics.csv'
    absence_path = path + 'absence.csv'

    day = input('请输入去除日期: ')

    remove_records(statistics_path, day, encoding)
    remove_records(absence_path, day, encoding)
