import csv
from attendance import *

def find_absence():
    workdays = set(work_calendar(2017, 5))
    dic = dict()
    staff_info = dict()
    attendance_list = []
    with open('/Users/hayden/Desktop/statistics.csv', 'r', encoding='gbk') as csv_file:
        spamreader = csv.reader(csv_file)
        for row in spamreader:
            if row[0] != '部门':
                #print(row)
                attendance_list.append(row)
                if row[1] not in dic:
                    staff_info[row[1]] = (row[0], row[2])
                    dic[row[1]] = [row[3]]
                else:
                    dic[row[1]].append(row[3])
        #print(dic)

    for staff in dic:
        dic[staff] = list(workdays - set(dic[staff]))

    #print('缺勤记录')
    for staff in dic:
        if dic[staff]:
            absence_list = sorted(dic[staff], key=lambda x: int(x.split('/')[2]))
            if absence_list:
                dic[staff] = absence_list



            #print('%s: 缺勤%d天 %s' % (staff, len(dic[staff]), sorted(dic[staff], key=lambda x: int(x.split('/')[2]))))

    #print(dic)
    #print(attendance_list)

    '''
    遍历attancanceList，同时向其中添加未打卡的天数的信息
    '''
    absence_list = attendance_list[:]

    for (staff, dates) in dic.items():
        for date in dates:
            absence_list.append([staff_info[staff][0], staff, staff_info[staff][1], date, '-', '-', '意念', '未打卡'])

    #print('--------')
    #print(absence_list)
    absence_list = sorted(absence_list, key=lambda x: (x[0], x[1], int(x[3].split('/')[2])))


    absence_list = [x for x in absence_list if x[7] != '正常']
    print(absence_list)

    '''
    写入到csv文件中
    '''
    with open('/Users/hayden/Desktop/absence.csv', 'w', encoding='gbk') as csv_file:
        writer = csv.writer(csv_file)
        writer.writerow(['部门', '姓名', '考勤号码', '工作日', '上班时间', '下班时间', '比对方式', '出勤情况'])
        for row in absence_list:
            writer.writerow(row)

if __name__ == '__main__':
    find_absence()
