# File name：
# Description:
# Author: zeng
# Time: 2020/4/11
# Change Activity:

import xlwt
import DataLoader
import datetime
from copy import deepcopy


best_mat = []
best_var = 999999999999999999999
flag = 0
counter = 0


def time_crash_check(task1, task2):
    '''
    输入为一个task_set，[[航班号1，航班号2，...]，起始时刻，终止时刻，总时间]
    :param task1:
    :param task2:
    :return: 0表示冲突 1表示无冲突
    '''
    if task2[1] <= task1[1] <= task2[2]:
        return 0
    if task1[1] <= task2[1] <= task1[2]:
        return 0
    return 1


def format_time(int_time):
    '''
    还原整形时间为字符串
    :param int_time: 距离第0天的00:00:00的过去了多少分钟，int
    :return: 形如'时钟：分钟'。string
    '''
    d, h, m = DataLoader.int_to_time(int_time)
    m = int(m)
    if m < 10:
        return str(h) + ':0' + str(m)
    return str(h) + ':' + str(m)


def var(l):
    '''
    输入一个数字列表，求方差
    :param l: 由数字构成的一维列表，list
    :return: 方差，float
    '''
    n = len(l)
    mean = sum(l)/n
    r = [(i-mean)**2 for i in l]
    return sum(r)/n


def count_time(tasks, mat):
    '''
    根据员工任务分配矩阵和任务集列表，计算该分配方案员工总工时的方差
    :param tasks: 特定时间段的任务集列表，每个任务集包括一辆车从3机坪出发到返回3机坪之间的全部任务，list
        [[[航班号(string) * M]，起始时刻(int)，终止时刻(int)，总时间(int)，车辆编号(int)] * N]
                 0                 1             2            3           4
    :param mat: 员工任务分配矩阵，行索引员工，列索引任务集，list
    :return: 当前任务时段员工总工时的方差，float
    '''

    staff_number = len(mat)
    col = len(tasks)

    time = [0 for _ in range(staff_number)]

    for i in range(staff_number):
        for j in range(col):
            if mat[i][j] == 1:
                time[i] += tasks[j][3]
    return var(time)


def staff_group(loc):
    '''
    读取员工信息表，按员工工作时间对其分组
    :param loc: 员工信息表路径，string
    :return: 按工作时间段分组的员工编号列表，list
    [[员工编号（string）*n], [员工编号（string）*p], [员工编号（string）*q]]
    '''
    staff_list = DataLoader.load_staff(loc)

    ##################################################
    # 注意：员工的工作时间不能交错，如果交错了要进一步细分 #
    ##################################################

    staff_by_part = []
    part = []
    flag = staff_list[0][1]
    while staff_list:
        number = staff_list[0][0]
        if staff_list[0][1] == flag:
            part.append(number)
            del staff_list[0]
        else:
            flag = staff_list[0][1]
            staff_by_part.append(part)
            part = []

    staff_by_part.append(part)

    return staff_by_part


def get_task_set(tasks, bus_number=36):
    '''
    处理车辆派工结果，得到按车辆划分的任务集合，每个任务集合按时间顺序包含同一辆车从3机坪出发到回到3机坪之间执行的全部任务
    :param tasks:
        [航班号，车辆编号，起始机坪，任务机坪，终点机坪，起始时刻，到位时刻，终止时刻，进出港类型，延误时间，发车车次，登机口，机位 ]
           0        1       2        3       4        5        6        7         8         9       10      11     12
    :param bus_number:
    :return:
        [[[航班号(string) * M]，起始时刻(int)，终止时刻(int)，总时间(int)，车辆编号(int)] * N]
                 0                 1             2            3           4
    '''

    task_by_bus = [[] for _ in range(bus_number)]   # 初始化按车辆划分的任务集合
    for task in tasks:
        number = task[1]  # 获取任务车辆编号
        task_by_bus[number].append(task)

    total_task_set = []  # 初始化以起点终点皆为3机坪的任务集合
    number = 0  # 记录当前车辆编号
    for bus in task_by_bus:
        # print(bus)
        task_set = [[]]  # [[航班号1，航班号2，...]，起始时刻，终止时刻，总时间，车辆编号]
        for task in bus:
            if task[2] == '3':  # 若起点为3机坪
                task_set[0].append(task[0])  # 航班列表添加航班
                task_set.append(task[5])  # 添加起始时刻
            else:
                task_set[0].append(task[0])  # 航班列表添加航班

            if task[4] == '3':  # 若终点为3机坪
                task_set.append(task[7])  # 添加终止时刻
                # print(task_set)
                total_time = task_set[2] - task_set[1]
                task_set.append(total_time)  # 添加总时间
                task_set.append(number)

                total_task_set.append(task_set)  # 当前任务集合完毕，加入total_task_set

                task_set = [[]]  # 重置task_set

        if task_set != [[]]:
            task_set.append(bus[-1][7])  # 添加终止时刻
            total_time = task_set[2] - task_set[1]
            task_set.append(total_time)  # 添加总时间
            task_set.append(number)
            total_task_set.append(task_set)

        number += 1

    total_task_set = sorted(total_task_set, key=(lambda x: x[1]))  # 按起始时刻排序

    return total_task_set


def allot_window(tasks, staff_number, m=5):
    lunch_time_start = 480
    lunch_time_end = 960

    # 11:00-2:00,分时用餐，用餐时间40分钟,则任务终止时间必须在13：20前
    # lunch_time_start = 660
    # lunch_time_end = 800

    #          0                 1        2       3      4
    # [[航班号1，航班号2，...]，起始时刻，终止时刻，总时间，车辆编号]
    task_set = tasks

    # 初始化人员时间表
    free_time = [0 for _ in range(staff_number)]  # 当前任务到来时，员工已执行任务的结束时间

    # 初始化人员工时表
    work_time = [0 for _ in range(staff_number)]

    # 初始化人员用餐表
    lunch_flags = [0 for _ in range(staff_number)]

    # 初始化任务分配列表
    row = staff_number
    col = len(task_set)
    mat = [[0 for _ in range(col)] for _ in range(row)]
    sum_lunch = 0

    task_number = 0
    while task_number < col:

        task = task_set[task_number]  # 当前任务
        start_time = task[1]  # 当前任务起始时间

        free_staff = []  # 当前任务到来时，空闲的员工
        for n in range(staff_number):
            if start_time >= free_time[n]:
                free_staff.append(n)


        free_staff_number = len(free_staff)  # 空闲员工数量
        if free_staff_number == 0:  # 无空闲员工，退出
            return 999999999, [], 0

        # print(task)
        # print(free_staff)
        # 空闲员工吃饭
        eating_staff = []
        if free_staff:
            for i in free_staff:
                # u = free_staff[i]
                if len(eating_staff) >= 2:
                    break
                if lunch_time_start <= free_time[i] % 1440 <= lunch_time_end and lunch_flags[i] == 0:
                    free_time[i] += 40  # 为午餐加上40分钟
                    lunch_flags[i] = 1  # 置标志为1
                    eating_staff.append(i)

        # print(lunch_flags)
        free_staff = []  # 剩下的空闲员工
        for n in range(staff_number):
            if start_time >= free_time[n]:
                free_staff.append(n)


        free_staff_number = len(free_staff)  # 空闲员工数量
        if free_staff_number == 0:  # 无空闲员工，拿一个人不吃饭
            free_staff.append(eating_staff[0])
            lunch_flags[eating_staff[0]] = 0  # 还原
            free_time[eating_staff[0]] -= 40  # 还原
            free_staff_number += 1


        # 当前空闲员工的工时
        part_time = [work_time[i] for i in free_staff]

        idx = sorted((range(free_staff_number)), key=lambda x: part_time[x])  # 记录part_time顺序排序后的索引
        staff_idx = [free_staff[i] for i in idx]

        # 确定当前窗口大小
        window_size = min((m, free_staff_number, len(task_set[task_number::])))

        # tasks_in_window = task_set[task_number:task_number + window_size]
        # 记录tasks_in_window按总时间逆序排序后的索引
        task_idx = sorted((range(task_number, task_number+window_size)), key=lambda x: task_set[x][3], reverse=True)

        for i in range(window_size):
            s = staff_idx[i]
            t = task_idx[i]

            mat[s][t] = 1
            # print(mat)
            # 更新工时
            work_time[s] += task_set[t][3]

            # 更新free_time和lunch_flags
            free_time[s] = task_set[t][2]

            # if lunch_time_start < task_set[t][2] % 1440 < lunch_time_end and lunch_flags[s] == 0:
            #
            #     free_time[s] += 40  # 为午餐加上40分钟
            #     lunch_flags[s] = 1  # 置标志为1
            #     print(lunch_flags)

        task_number += window_size

    v = count_time(task_set, mat)
    sum_lunch = 0
    for i in lunch_flags:
        sum_lunch += i
    return v, mat, sum_lunch


def allot_rec(tasks, staff_number, d=2):

    # 11:00-2:00,分时用餐，用餐时间40分钟,则任务终止时间必须在13：20前
    lunch_time_start = 660
    lunch_time_end = 800

    #          0                 1        2       3      4
    # [[航班号1，航班号2，...]，起始时刻，终止时刻，总时间，车辆编号]
    task_set = tasks

    # 初始化人员时间表
    free_moment = [0 for _ in range(staff_number)]  # 当前任务到来时，员工已执行任务的结束时间

    # 初始化人员工时表
    work_time = [0 for _ in range(staff_number)]

    # 初始化人员用餐表
    lunch_flags = [0 for _ in range(staff_number)]

    # 初始化任务分配列表
    row = staff_number
    col = len(task_set)
    mat = [[0 for _ in range(col)] for _ in range(row)]

    # 初始化结果
    global best_mat
    global best_var
    global flag
    global counter
    best_mat = []
    best_var = 999999999999999999999
    flag = 0
    counter = 0

    def rec(j):  # 第j个任务

        global best_var
        global best_mat
        global flag
        global counter

        if flag == 1:
            return 0

        if j >= col:  # 执行完毕的出口
            v = count_time(task_set, mat)
            if v <= best_var:
                best_mat = deepcopy(mat)
                best_var = v
            elif counter > d and v > best_var:
                flag = 1  # 剪枝标记
            counter = 0
            return 0

        task = task_set[j]  # 第j个任务
        start_time = task[1]  # 当前任务起始时间

        free_staff = []  # 当前任务到来时，空闲的员工
        for n in range(staff_number):
            if start_time > free_moment[n]:
                free_staff.append(n)

        # 空闲员工吃饭
        if free_staff:
            for i in free_staff:
                # u = free_staff[i]
                if lunch_time_start <= free_moment[i] % 1440 <= lunch_time_end and lunch_flags[i] == 0:
                    free_moment[i] += 40  # 为午餐加上40分钟
                    lunch_flags[i] = 1  # 置标志为1

        free_staff = []  # 当前任务到来时，空闲的员工
        for n in range(staff_number):
            if start_time > free_moment[n]:
                free_staff.append(n)

        # 当前空闲员工的工时
        part_time = [work_time[i] for i in free_staff]
        idx = sorted((range(len(part_time))), key=lambda x: part_time[x])
        part_time = sorted(part_time)  # 从小到大排序

        if not free_staff:  # 失败的出口，若任务到来无空闲员工，退出
            return 1

        for k in idx:

            i = free_staff[k]  # i, 空闲的员工i

            # 剪枝1
            if i > j:
                return 1

            # 更新任务分配矩阵
            mat[i][j] = 1
            p1 = i
            p2 = j

            # 更新工时
            record0 = work_time[i]
            work_time[i] += task[3]

            # 剪枝2
            # if max(work_time) - min(work_time) > 150:
            #     return 1

            # 更新free_time和lunch_flags
            record1 = free_moment[i]
            # record2 = lunch_flags[i]
            free_moment[i] = task[2]
            # if lunch_time_start < task[2] % 1440 < lunch_time_end and lunch_flags[i] == 0:
            #     free_moment[i] = task[2] + 40  # 为午餐加上40分钟
            #     lunch_flags[i] = 1  # 置标志为1

            # 递归到下一个任务
            rec(j+1)

            # 回溯还原
            mat[p1][p2] = 0
            work_time[i] = record0
            free_moment[i] = record1
            # lunch_flags[i] = record2
            # 计数器更新
            counter += 1
    rec(0)
    return best_var, best_mat


def control(tasks, loc, loc_air, b1, b2, p=3, method=1):
    '''
    人员调度主函数
    :param tasks:
    :param loc:
    :param p:
    :param method:
    :return:
    '''
    print('派车次数', len(tasks))

    staff_by_part = staff_group(loc)

    #    0      1        2         3        4        5       6        7        8         9        10      11    12
    # [航班号，车辆编号，起始机坪，任务机坪，终点机坪，起始时刻，到位时刻，终止时刻，进出港类型，延误时间，发车车次，登机口，机位 ]
    tasks = sorted(tasks, key=(lambda x: x[5]))  # 对任务按起始时间排序

    bus = [t[1] for t in tasks]
    bus_number = len(set(bus))
    #          0                 1        2       3      4
    # [[航班号1，航班号2，...]，起始时刻，终止时刻，总时间，车辆编号]
    task_by_bus = get_task_set(tasks, bus_number=bus_number)  # 以从3号机坪开始到返回3号机坪，划分每一辆摆渡车的任务为若干任务集合，并计算每个集合的任务总时长

    temp = 0
    for i in range(len(task_by_bus)):
        temp += len(task_by_bus[i][0])
    print('任务总数（3机坪起始任务集）：', len(task_by_bus), '车辆总数：', bus_number, temp)

    # 划分天，划分时间段
    days, _, _ = DataLoader.int_to_time(task_by_bus[-1][1])  # 以最后一个任务的起始时刻判断一共多少天
    days = days + 1

    # 初始化结果表及其表头
    r = []  # 派遣结果列表

    stat = []  # 员工派遣信息统计
    c = 1  # 任务计数器
    for day in range(days):

        # 取出第day天的任务，并将这些任务从task_by_bus里移除
        task_by_day = []
        while task_by_bus:
            if task_by_bus[0][1] < (day+1)*1440:
                task_by_day.append(task_by_bus[0])
                del task_by_bus[0]
            else:
                break

        print('\n第', day, '天任务总数：', len(task_by_day))
        r_by_day = []
        stat_by_day = []
        for part in range(3):

            # 取出第part个时间段的任务，并将这些任务从task_by_day里移除
            task_by_part = []
            while task_by_day:
                if task_by_day[0][1] < day*1440+(part+1)*8*60:
                    task_by_part.append(task_by_day[0])
                    del task_by_day[0]
                else:
                    break

            temp = 0
            for i in range(len(task_by_part)):
                temp += len(task_by_part[i][0])
            print('第', day, '天，第', part, '时间段，任务数:', len(task_by_part), temp)


            if task_by_part:

                # 确定该时刻员工人数
                number = len(staff_by_part[part])

                mat = []
                v = 99999999999999
                lunch_c = 0
                # 给第day天第part时间段的任务派工
                if method == 0:  # 贪心算法
                    v, mat, _ = allot_window(task_by_part, number, 1)

                elif method == 1:  # 时间窗算法
                    for i in range(1, 21):  # 以不同窗口计算选择最优
                        v1, mat1, lunch_c1 = allot_window(task_by_part, number, i)
                        # print(v1, lunch_c1)
                        if lunch_c1 > lunch_c:
                            lunch_c = lunch_c1
                            v = v1
                            mat = mat1
                        elif lunch_c1 == lunch_c and v1 < v:
                            v = v1
                            mat = mat1
                    # print(v, lunch_c)
                    # print()

                elif method == 2:  # 递归算法
                    for i in range(1, 21):  # 确保有解后才能使用递归算法
                        v1, mat1, _ = allot_window(task_by_part, number, i)
                        if v1 < v:
                            v = v1
                            mat = mat1
                    if mat:
                        v, mat = allot_rec(task_by_part, number, p)

                print(v, mat)
                r_part = format_out(task_by_part, tasks, staff_by_part[part], mat, b1, b2)
                stat_by_day.append(staff_out(c, r_part))
                c += len(r_part)

                r += r_part
                # r_by_day += r_part
            else:
                print('当前时段无任务')
                stat_by_day.append([])

        stat.append(stat_by_day)
        # ss = staff_out(c, r_by_day)
        # s.append(ss)
        # c += len(r_by_day)

    start_date = DataLoader.get_start_date(loc_air)  # (year,month,day,hour,minute,nearest_second）
    start_date = str(start_date[0]) + '-' + str(start_date[1]) + '-' + str(start_date[2])
    start_date = datetime.datetime.strptime(start_date, "%Y-%m-%d")

    if time_to_int(r[0][8]) > time_to_int(r[0][9]):
        start_date = (start_date + datetime.timedelta(days=-1))
    # start_date = str(start_date)[0:9]

    export_staff(stat, start_date, '派遣结果统计.xls')

    head = [['任务编号', '航班号', '任务执行登机口', '任务执行机位', '车型', '车辆编号', '任务类型', '发车车次', '到位时间', '发车时间', '结束时间', '任务时长', '人员编号']]
    r = head + r
    export_xls(r, '摆渡车任务.xls')


def judge(tasks, loc, p=3):
    '''
    计算人员缺口
    :param tasks:
    :param loc:
    :param p: 算法参数，对应时间窗算法中的窗口大小
    :return:
    '''
    staff_by_part = staff_group(loc)

    #    0      1        2         3        4        5       6        7        8         9        10      11    12
    # [航班号，车辆编号，起始机坪，任务机坪，终点机坪，起始时刻，到位时刻，终止时刻，进出港类型，延误时间，发车车次，登机口，机位 ]
    tasks = sorted(tasks, key=(lambda x: x[5]))  # 对任务按起始时间排序
    bus = [t[1] for t in tasks]
    bus_number = len(set(bus))
    #          0                 1        2       3      4
    # [[航班号1，航班号2，...]，起始时刻，终止时刻，总时间，车辆编号]
    # 以从3号机坪开始到返回3号机坪，划分每一辆摆渡车的任务为若干任务集合，并计算每个集合的任务总时长
    task_by_bus = get_task_set(tasks, bus_number=bus_number)

    # 划分天，划分时间段
    days, _, _ = DataLoader.int_to_time(task_by_bus[-1][1])  # 以最后一个任务的起始时刻判断一共多少天
    days = days + 1

    needed_list = []

    for day in range(days):

        # 取出第day天的任务，并将这些任务从task_by_bus里移除
        task_by_day = []
        while task_by_bus:
            if task_by_bus[0][1] < (day + 1) * 1440:
                task_by_day.append(task_by_bus[0])
                del task_by_bus[0]
            else:
                break

        needed_l = []
        for part in range(3):

            # 取出第part个时间段的任务，并将这些任务从task_by_day里移除
            task_by_part = []
            while task_by_day:
                if task_by_day[0][1] < day * 1440 + (part + 1) * 8 * 60:
                    task_by_part.append(task_by_day[0])
                    del task_by_day[0]
                else:
                    break

            # print('第', day, '天，第', part, '时间段')

            needed = -1  # 初始化人员缺口数
            if task_by_part:

                # 确定该时刻员工人数
                number = len(staff_by_part[part])

                mat = []
                v = 99999999999999

                while mat == []:
                    needed += 1
                    v, mat, sss = allot_window(task_by_part, number+needed, p)

            needed = 0 if needed < 0 else needed
            needed_l.append(needed)
        needed_list.append(needed_l)

    r = [0 for _ in range(3)]
    for i in range(3):
        a = [it[i] for it in needed_list]
        r[i] = max(a)
    return r


def export_xls(r, file_name):
    book = xlwt.Workbook()  # 新建一个Excel
    sheet = book.add_sheet('sheet1')  # 新建一个sheet页

    col = 0  # 列号
    for j in r[0]:  # 控制列的
        sheet.write(0, col, j)
        col += 1  # 列号

    count = 1
    row = 1  # 行号
    for i in r[1::]:  # 控制行
        sheet.write(row, 0, count)
        count += 1
        col = 1  # 列号
        for j in i:  # 控制列的
            sheet.write(row, col, j)
            col += 1  # 列号
        row += 1

    book.save(file_name)  # 保存内容


def export_staff(s, start_date, file_name):
    book = xlwt.Workbook()  # 新建一个Excel
    sheet = book.add_sheet('sheet1')  # 新建一个sheet页

    head = ['员工编号', '分配任务数', '任务时长', '工作时长', '工时利用率', '派遣任务编号']

    col = 0  # 列号
    for j in head:  # 控制列的
        sheet.write(0, col, j)
        col += 1  # 列号

    row = 1  # 行号
    for day in range(len(s)):

        date = (start_date + datetime.timedelta(days=day)).strftime("%Y-%m-%d")
        date = str(date)[0:10]

        sheet.write(row, 0, date)
        row += 1

        for part in range(len(s[day])):
            if s[day][part] == []:
                continue

            start = str(8 * part)
            end = str(8 + 8 * part)
            sheet.write(row, 0, date + ' ' + start + '-' + end + '时')
            row += 1

            for i in s[day][part]:  # 控制行
                col = 0  # 列号
                for j in i:  # 控制列的
                    sheet.write(row, col, j)
                    col += 1  # 列号
                row += 1

    book.save(file_name)  # 保存内容


def format_out(task_by_part, origin_tasks, staff_by_part, mat, b1, b2):

    # task_by_part
    #          0                 1        2       3      4
    # [[航班号1，航班号2，...]，起始时刻，终止时刻，总时间，车辆编号]

    r = []

    if not mat:
        return r

    number = len(staff_by_part)
    n = len(task_by_part)

    for i in range(number):
        for j in range(n):
            if mat[i][j] == 1:

                staff = staff_by_part[i]
                bus = task_by_part[j][4]

                b = 'COBUS3000' if bus < b1 else '北方'

                z = ''
                if bus < 9 or (b1 - 1) < bus < (b1 + 9):
                    z = '0'
                b_n = 'COBUS-' + z + str(bus + 1) if bus < b1 else '北方-' + z + str(bus - b1 + 1)

                for air in task_by_part[j][0]:

                    gate = ''  # 登机口
                    pos = ''  # 机位
                    tp = ''  # 进出港类型
                    num = ''  # 车次
                    time1 = 0  # 到位时间
                    time2 = 0  # 发车时间
                    time3 = 0  # 结束时间
                    time4 = 0  # 时长
#    0      1        2         3        4        5       6        7        8         9        10      11    12
# [航班号，车辆编号，起始机坪，任务机坪，终点机坪，起始时刻，到位时刻，终止时刻，进出港类型，延误时间，发车车次，登机口，机位 ]

                    for t in origin_tasks:
                        if air == t[0] and bus == t[1]:
                            if t[11] != 'Empty':
                                gate = t[11]  # 登机口
                            pos = t[12]  # 机位
                            tp = t[8]  # 进出港类型
                            num = t[10]  # 车次
                            time1 = format_time(t[6])  # 到位时间
                            time2 = format_time(t[5])  # 发车时间
                            time3 = format_time(t[7])  # 结束时间
                            time4 = int(t[7] - t[6])  # 时长
                            break

                    total = [air, gate, pos, b, b_n, tp, num, time1, time2, time3, time4, staff]
                    r.append(total)
    return r


def staff_out(start_i, r):

    if r == []:
        return r

    r = [[start_i+i]+r[i] for i in range(len(r))]
    # print(r)

#    0       1         2            3          4     5         6       7         8        9       10       11       12
# 任务编号, 航班号, 任务执行登机口, 任务执行机位, 车型, 车辆编号, 任务类型, 发车车次, 到位时间, 发车时间, 结束时间, 任务时长, 人员编号
    staffs = sorted(set([x[12] for x in r]))

    #   0          1         2          3           4           5
    # 员工编号，分配任务数， 任务时长， 起始时间集， 结束时间集， 派遣任务编号
    stat_by_day_init = [[staffs[x], 0, 0, [], [], []] for x in range(len(staffs))]

    #   0          1         2          3           4           5
    # 员工编号，分配任务数， 任务时长， 上班时长， 工时利用率， 派遣任务编号
    stat_by_day = [[staffs[x], 0, 0, 0, 0, ''] for x in range(len(staffs))]

    for t in r:
        staff_id = t[12]
        idx = staffs.index(staff_id)
        stat_by_day[idx][1] += 1

        a = time_to_int(t[10])
        b = time_to_int(t[8])
        c = time_to_int(t[9])
        d = a-b if a > b else a+1440-b
        e = a-c if a > c else a+1440-c

        stat_by_day[idx][2] += d
        stat_by_day[idx][3] += e

        stat_by_day[idx][5] += (str(t[0]) + ',')

    for i in stat_by_day:
        i[2] = i[2] / 60
        i[3] = i[3] / 60
        i[4] = i[2] / i[3]
        i[5] = i[5][0:-1]

    return stat_by_day


def time_to_int(s):
    hour, minute = [int(i) for i in s.split(':')]
    return hour*60 + minute
