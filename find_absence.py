from attendance import *

'''
读取统计表，查找缺勤日期，最后生成一张统计表记录一个月期间不正常打卡的记录
'''
def find_absence():
    workdays = work_calendar(2017, 5)
    # 字典类型，记录员工出勤日期，键为姓名，值为一个列表，存储了有打卡记录的日期
    staff_attendance = dict()
    # 字典类型，记录员工的信息，键为姓名，值为一个元组: (部门名称, 考勤号)
    staff_info = dict()
    # 出勤表，记录Excel表格中每一行的信息
    attendance_list = []

    # 读取前面生成的统计表
    with open('/Users/hayden/Desktop/statistics.csv', 'r', encoding='gbk') as csv_file:
        reader = csv.reader(csv_file)
        for row in reader:
            if row[0] != '部门':
                # 将记录一条加入到出勤表中
                attendance_list.append(row)
                # 将出勤天数加入到字典中
                if row[1] not in staff_attendance:
                    staff_info[row[1]] = (row[0], row[2])
                    staff_attendance[row[1]] = [row[3]]
                else:
                    staff_attendance[row[1]].append(row[3])

    # 用当月工作日减去实际出勤天数，得出每个员工当月未上班天数
    for staff in staff_attendance:
        staff_attendance[staff] = list(set(workdays) - set(staff_attendance[staff]))

    # print('缺勤记录: ')
    for staff in staff_attendance:
        if staff_attendance[staff]:
            # 对未出勤日期进行排序
            staff_attendance[staff] = sorted(staff_attendance[staff], key=lambda x: int(x.split('/')[2]))
            # print('%s: 缺勤%d天 %s' % (staff, len(dic[staff]), sorted(dic[staff], key=lambda x: int(x.split('/')[2]))))

    '''
    遍历attancanceList，同时向其中添加未打卡的天数的信息
    '''
    absence_list = attendance_list[:]

    for (staff, dates) in staff_attendance.items():
        for date in dates:
            absence_list.append(
                [staff_info[staff][0], staff, staff_info[staff][1], date, '00:00:00', '00:00:00', '意念', '未打卡'])

    # 各条记录排序: 1.按部门 2.按姓名 3.按日期
    absence_list = sorted(absence_list, key=lambda x: (x[0], x[1], int(x[3].split('/')[2])))

    '''
    前一天晚上奋斗到22点以后的，抵扣第二天的迟到
    '''
    for index in range(1, len(absence_list)):
        if '迟到' in absence_list[index][7]:
            # 检查前一天
            if absence_list[index - 1][1] == absence_list[index][1] and int(
                    absence_list[index - 1][5].split(':')[0]) >= 22:
                # 打印一下抵扣迟到的记录
                print(absence_list[index])
                # 如果只有迟到的，标为正常。如果当天还有别的记录，仅仅抹去迟到
                if absence_list[index][7] == '迟到':
                    absence_list[index][7] = '正常'
                else:
                    absence_list[index][7] = absence_list[index][7][2:]

    # 移去正常的记录，只保留不正常的
    absence_list = [x for x in absence_list if x[7] != '正常']
    # print(absence_list)

    '''
    写入到csv文件中
    '''
    with open('/Users/hayden/Desktop/absence.csv', 'w', encoding='gbk') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['部门', '姓名', '考勤号码', '工作日', '上班时间', '下班时间', '比对方式', '出勤情况'])
        # 循环将每一行依次写入到csv文件中
        for row in absence_list:
            writer.writerow(row)


'''
主函数入口
'''
if __name__ == '__main__':
    find_absence()
