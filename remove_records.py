#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import csv

# FIXME: 将输入区间少设一天，来解决此问题？

def remove_records(path, day, encoding='gbk'):

    records = []

    with open(path, 'r', encoding=encoding) as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            if row[3] != day:
                records.append(row)

    with open(path, 'w', encoding=encoding) as csv_file:
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
