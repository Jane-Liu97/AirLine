#改
from DataLoader import get_runtime, load_runtime, load_airline, load_staff,get_start_date

def get_task(d_airline):
    '''
    任务池初始化
    为方便计算，将时间简化为分钟，如：0:10 = 10,1:10 = 70
    列表中一个元素的构成为
        [航班编号（string），STA/STD（int），进出港类型（string），所在机坪（string），登机口所在机坪（string），
        首车到达时间（int），首车发车时间（int），次车到达时间（int），次车发车时间（int），尾车到达时间（int），尾车发车时间（int）,延误（int）]
    '''

    # 待响应任务
    task_todo = []

    # 待响应任务池初始化，将航班信息导入待响应任务
    for item in d_airline:

        # 航班号
        al_id = item[0]

        # 时间
        al_time = item[1]

        # 进出港类型
        al_type = item[2]

        # 机坪
        al_port = item[3]

        # 机位
        al_pos = item[7]

        # 登机口所在机坪
        al_aboard = None
        if al_type == '出港':
            al_aboard = item[6]
        if al_aboard is None:
            al_aboard = 'Empty'

        # 首车到达
        f_arrive = None
        # 首车发车
        f_depart = None
        # 次车到达
        s_arrive = None
        # 次车发车
        s_depart = None
        # 尾车到达
        l_arrive = None
        # 尾车发车
        l_depart = None
        if al_type == '进港':
            f_arrive = al_time - 5
            f_depart = al_time + 13
            s_arrive = al_time
            s_depart = al_time + 15
            l_arrive = al_time + 5
            l_depart = al_time + 17
        else:
            f_arrive = al_time - 35
            f_depart = al_time - 25
            s_arrive = al_time - 25
            s_depart = al_time - 19
            l_arrive = al_time - 19
            l_depart = al_time - 5

        if al_type == '出港':
            if item[6] is None:
                continue
        task_todo.append([al_id, al_time, al_type, al_port, al_aboard, f_arrive, f_depart, s_arrive, s_depart, l_arrive, l_depart, 0, al_pos])

    # 按首车到达时间排序
    task_todo = sorted(task_todo, key=(lambda x: x[5]))

    return task_todo


def get_bus(n_bus1, n_bus2):
    '''
    资源池（车辆）初始化
    车辆编号 —— 0,1,2,3...（0-25为大摆渡，26-35为大VIP）
    列表中一个元素的构成为
        [车辆编号（int），行驶时间（int），所在机坪（string）,可用时间（int）]
    '''

    # 大摆渡
    bus1 = []

    # 大VIP
    bus2 = []

    for i in range(n_bus1):
        bus1.append([i, 0, '3', -1000, None, None])

    for i in range(n_bus2):
        bus2.append([i + n_bus1,0, '3', -1000, None, None])

    return bus1, bus2

def get_bus_buffer(bus1, bus2, task, place, runtime, s1, s2, s3, bus_done):
    '''
        进行车辆调度匹配，获取缓冲池
        返回值bus_buffer,flag2(flag2=1:因人员不足产生资源缺口)
        列表中一个元素的构成为
            [车辆编号（int），车次（int）]
    '''
    flag2 = 0

    # flag标记：0：人数充足
    flag = None
    t = task[5] % 1440
    # 0-8
    if t < 480:
        flagt = 1
        if len(bus_done) >= s1:
            flag = 0
        else:
            flag = s1 - len(bus_done)
    # 8-16
    elif t < 960:
        flagt = 2
        if len(bus_done) >= s2:
            flag = 0
        else:
            flag = s2 - len(bus_done)
    # 16-24
    else:
        flagt = 3
        if len(bus_done) >= s3:
            flag = 0
        else:
            flag = s3 - len(bus_done)

    bus_buffer = []
    for bus in bus1:
        # 计算该车辆到达任务机坪的路程时间
        time1 = get_runtime(bus[2], place, runtime)
        # 计算该车辆可到达任务机坪的最早时间
        time2 = bus[3] + get_runtime(bus[2], place, runtime)
        bus[4] = time1
        bus[5] = time2
    for bus in bus2:
        # 计算该车辆到达任务机坪的路程时间
        time1 = get_runtime(bus[2], place, runtime)
        # 计算该车辆可到达任务机坪的最早时间
        time2 = bus[3] + get_runtime(bus[2], place, runtime)
        bus[4] = time1
        bus[5] = time2

    bus1 = sorted(bus1, key=(lambda x: (x[4], x[5], x[1])))
    bus2 = sorted(bus2, key=(lambda x: (x[4], x[5], x[1])))

    n = 0
    # 资源池（大摆渡）中依次寻找可用车辆
    for bus in bus1:

        o = 0
        # 首车 if 该车辆可按时到达任务机坪
        if bus[5] <= task[5]:
            for b in bus_buffer:
                if b[1] == 0:
                    o = 1
                if b[0] == bus[0]:
                    o = 1
                # 在3机坪且人员不够
                if bus[2] == '3' and flag < 1:
                    if o == 0:
                        flag2 = 1
                    o = 1
            if o != 1:
                # 加入缓冲池
                bus_buffer.append([bus[0], 0])
                n += 1
                if bus[2] == '3':
                    flag = flag - 1

        o = 0
        # 次车 if 该车辆可按时到达任务机坪
        if bus[5] <= task[7]:
            for b in bus_buffer:
                if b[1] == 1:
                    o = 2
                if b[0] == bus[0]:
                    o = 2
                # 在3机坪且人员不够
                if bus[2] == '3' and flag < 1:
                    if o == 0:
                        flag2 = 1
                    o = 2
            if o != 2:
                # 加入缓冲池
                bus_buffer.append([bus[0], 1])
                n += 1
                if bus[2] == '3':
                    flag = flag - 1

        o = 0
        # 尾车 if 该车辆可按时到达任务机坪
        if bus[5] <= task[9]:
            for b in bus_buffer:
                if b[1] == 2:
                    o = 3
                if b[0] == bus[0]:
                    o = 3
                # 在3机坪且人员不够
                if bus[2] == '3' and flag < 1:
                    if o == 0:
                        flag2 = 1
                    o = 3
            if o != 3:
                # 加入缓冲池
                bus_buffer.append([bus[0], 2])
                n += 1
                if bus[2] == '3':
                    flag = flag - 1

        # 已有3辆大摆渡满足条件，寻找合适的大VIP
        if n == 3:
            o = 0
            # 资源池（大VIP）中依次寻找可用车辆
            for bus in bus2:
                # if 该车辆可按时到达任务机坪
                if bus[5] <= task[5]:
                    # 在3机坪且人员不够
                    if bus[2] == '3' and flag < 1:
                        flag2 = 1
                        o = 4
                    if o != 4:
                        # 加入缓冲池
                        bus_buffer.append([bus[0], 3])
                        n += 1
                        if bus[2] == '3':
                            flag = flag - 1
                        break
            break

    return bus_buffer, flag2, flagt

def get_buffer(bus1, bus2, task, place, runtime, bus_done):
    '''
        进行车辆调度匹配，获取缓冲池
        返回值bus_buffer,flag2(flag2=1:因人员不足产生资源缺口)
        列表中一个元素的构成为
            [车辆编号（int），车次（int）]
    '''
    

    bus_buffer = []
    for bus in bus1:
        # 计算该车辆到达任务机坪的路程时间
        time1 = get_runtime(bus[2], place, runtime)
        # 计算该车辆可到达任务机坪的最早时间
        time2 = bus[3] + get_runtime(bus[2], place, runtime)
        bus[4] = time1
        bus[5] = time2
    for bus in bus2:
        # 计算该车辆到达任务机坪的路程时间
        time1 = get_runtime(bus[2], place, runtime)
        # 计算该车辆可到达任务机坪的最早时间
        time2 = bus[3] + get_runtime(bus[2], place, runtime)
        bus[4] = time1
        bus[5] = time2

    bus1 = sorted(bus1, key=(lambda x: (x[4], x[5], x[1])))
    bus2 = sorted(bus2, key=(lambda x: (x[4], x[5], x[1])))

    n = 0
    # 资源池（大摆渡）中依次寻找可用车辆
    for bus in bus1:
        o = 0
        # 首车 if 该车辆可按时到达任务机坪
        if bus[5] <= task[5]:
            for b in bus_buffer:
                if b[1] == 0:
                    o = 1
                if b[0] == bus[0]:
                    o = 1
            if o != 1:
                # 加入缓冲池
                bus_buffer.append([bus[0], 0])
                n += 1

        o = 0
        # 次车 if 该车辆可按时到达任务机坪
        if bus[5] <= task[7]:
            for b in bus_buffer:
                if b[1] == 1:
                    o = 2
                if b[0] == bus[0]:
                    o = 2
            if o != 2:
                # 加入缓冲池
                bus_buffer.append([bus[0], 1])
                n += 1

        o = 0
        # 尾车 if 该车辆可按时到达任务机坪
        if bus[5] <= task[9]:
            for b in bus_buffer:
                if b[1] == 2:
                    o = 3
                if b[0] == bus[0]:
                    o = 3
                    
            if o != 3:
                # 加入缓冲池
                bus_buffer.append([bus[0], 2])
                n += 1

        # 已有3辆大摆渡满足条件，寻找合适的大VIP
        if n == 3:
            # 资源池（大VIP）中依次寻找可用车辆
            for bus in bus2:
                # if 该车辆可按时到达任务机坪
                if bus[5] <= task[5]:
                    # 加入缓冲池
                    bus_buffer.append([bus[0], 3])
                    n += 1
                    break
            break

    return bus_buffer

def get_num_staff(staff):
    '''
        获取每个时段的员工数量
        s1:0-8;s2:8-16;s3:16-24
    '''
    # 工时0-8
    s1 = 0
    #t1 = [0,480]
    # 工时8-16
    s2 = 0
    #t2 = [480,960]
    # 工时16-24
    s3 = 0
    #t3 = [960,1440]
    # 员工数量
    for s in staff:
        if s[1] == 0:
            s1 += 1
        elif s[1] == 8:
            s2 += 1
        else:
            s3 += 1
    return s1, s2, s3

def add_day(dispatch):
    for dis in dispatch:
        dis[5] += 1440
        dis[6] += 1440
        dis[7] += 1440
    return dispatch

def get_need1(n_bus1, n_bus2, loc_airline, loc_gate,excel_loc_runtime):
    '''
        计算资源缺口1（满足所有航班不延误）
        返回值依次为大摆渡缺口，大VIP缺口，调度结果，0-8员工缺口，8-16员工缺口，16-24员工缺口
    '''

    n1 = n_bus1
    n2 = n_bus2
    ok = 0
    
    
    task_todo = get_task(load_airline(loc_airline, loc_gate))
    # 距离矩阵
    runtime = load_runtime(excel_loc_runtime)
    

    while ok == 0:

        # 资源池初始化
        bus1, bus2 = get_bus(n1, n2)

        #ok,b1,b2 = run_need1(task_todo,bus1,bus2)
        ok, b1, b2, dispatch1 = run_need1(task_todo, bus1, bus2,runtime)
        n1 = n1 + b1
        n2 = n2 + b2
        
    dispatch1 = add_day(dispatch1)
    # return n1-n_bus1,n2-n_bus2
    return n1 - n_bus1, n2 - n_bus2, dispatch1

def run_need1(task_todo, bus1, bus2,runtime):
    '''
        考虑资源缺口1的调度
        从待响应任务池依次派遣任务
        返回ok，b1，b2
        ok=0表示航班延误情况存在，ok=1表示资源充足，航班不延误
        b1代表大摆渡缺口，b2代表大VIP缺口
    '''

    '''
        遣回列表初始化
        遣回列表用于在每一次任务调度前对派出车辆进行遣回
        列表中一个元素的构成为
            [车辆编号（int），所在机坪（string）,可用时间（int）,航班编号（string）]
    '''

    # 已派出且不在3机坪
    bus_done = []

    '''
        缓冲池初始化
        列表中一个元素的构成为
            [车辆编号（int），第几辆车（int）]
    '''

    # 缓冲池
    bus_buffer = []

    '''
        调度列表初始化
        列表中一个元素的构成为
            [航班编号（string），车辆编号（int），起始机坪（string），任务机坪（string），返回机坪（string），
            起始时刻（int），航班时刻（int），终止时刻（int）,进出港类型（string）,延误时间（int）,机位（string）]
    '''

    # 调度列表
    dispatch_bus = []

    # 已分配任务
    task_done = []

    # 延误任务
    task_delay = []

    n_bus1 = len(bus1)

    task_cp = task_todo.copy()

    for x in task_delay:
        while task_delay.count(x) > 1:
            del task_delay[task_delay.index(x)]
    # print(dispatch_bus)
    #x = 0
    flag0 = 0
    for task in task_todo:
        bus1 = sorted(bus1, key=(lambda x: (x[3], x[1])))
        bus2 = sorted(bus2, key=(lambda x: (x[3], x[1])))
        # print(len(dispatch_bus))
        # 按时间排序
        task_todo = sorted(task_todo, key=(lambda x: x[5]))

        # 任务起始地点，出港航班为登机口，进港航班为航班所在机坪
        if task[2] == '出港':
            place = task[4]
        else:
            place = task[3]

        #i = 0
        # 无任务车辆遣返
        # print(bus_done)
        if len(dispatch_bus) != 0:
            bu = []
            for i in range(len(bus1) + len(bus2)):
                bu.append([])

            for dis in dispatch_bus:
                # 航班号，起始机坪，结束机坪，结束时间，起始时间
                bu[dis[1]].append([dis[0], dis[2], dis[4], dis[7], dis[5]])

            # 最后一次从3机坪派任务-索引
            bun = []
            for i in range(len(bus1) + len(bus2)):
                bun.append([])
            bui = 0
            for b in bu:
                bi = 0
                for i in b:
                    if i[2] == '3':
                        break
                    bi += 1
                bi -= 1
                bun[bui].append(bi)
                bui += 1

            # 8:00-11:40派出
            flagb = []
            for i in range(len(bus1) + len(bus2)):
                flagb.append([])
            bi = 0

            for order in bun:

                if len(bu[bi]) != 0:
                    if bu[bi][order[0]][4] % 1440 > 480 and bu[bi][order[0]][4] % 1440 < 700:
                        flagb[bi].append(1)
                    else:
                        flagb[bi].append(0)
                else:
                    flagb[bi].append(0)
                bi += 1

            for f in flagb:
                if task[5] % 1440 > 770 and f == 1:
                    f.append(1)
                else:
                    f.append(0)

        for bus_d in bus_done:
            f = 0
            if flagb[bus_d[0]][0] == 1 and flagb[bus_d[0]][1] == 1:
                f = 1

            # 计算该车辆等待的时间
            time_wait = task[5] - bus_d[2]
            # 到达任务机坪
            time_go = get_runtime(bus_d[1], '3', runtime) + get_runtime('3', place, runtime)
            # 等待时间大于往返时间，遣回
            if time_go < time_wait or f == 1:
                # 大摆渡
                if bus_d[0] < n_bus1:
                    for bus in bus1:
                        # 资源池中找到该车辆
                        if bus[0] == bus_d[0]:

                            # 修改调度列表
                            for dis_bus in dispatch_bus:
                                if dis_bus[1] == bus_d[0]:
                                    if dis_bus[0] == bus_d[3]:
                                        bus[1] = bus[1] + get_runtime(bus_d[1], '3', runtime)
                                        bus[2] = '3'
                                        bus[3] = bus[3] + get_runtime(bus_d[1], '3', runtime)
                                        dis_bus[4] = '3'
                                        dis_bus[7] = bus[3]
                                        i = 0
                                        # 移出遣回列表
                                        for d in bus_done:
                                            if d[0] == bus_d[0]:
                                                bus_done.pop(i)
                                            i += 1
                                        break
                # 大VIP
                else:
                    for bus in bus2:
                        # 资源池中找到该车辆
                        if bus[0] == bus_d[0]:

                            # 修改调度列表
                            for dis_bus in dispatch_bus:
                                if dis_bus[1] == bus_d[0]:
                                    if dis_bus[0] == bus_d[3]:
                                        bus[1] = bus[1] + get_runtime(bus_d[1], '3', runtime)
                                        bus[2] = '3'
                                        bus[3] = bus[3] + get_runtime(bus_d[1], '3', runtime)
                                        dis_bus[4] = '3'
                                        dis_bus[7] = bus[3]
                                        i = 0
                                        # 移出遣回列表
                                        for d in bus_done:
                                            if d[0] == bus_d[0]:
                                                bus_done.pop(i)
                                            i += 1
                                        break

        # 车辆匹配-获取缓冲池
        bus_buffer.clear()
        bus_buffer = get_buffer(bus1, bus2, task, place, runtime,bus_done)

        # 有足够资源
        if len(bus_buffer) == 4:

            for bus_b in bus_buffer:
                # 移出遣返列表
                di = 0
                for bus_d in bus_done:
                    if bus_b[0] == bus_d[0]:
                        bus_done.pop(di)
                        break
                    di += 1

            # print(bus_buffer)
            # 缓冲池顺序：大摆渡1，大摆渡2，大摆渡3，大VIP
            # 将缓冲池中的车辆进行相关更新
            for bus_b in bus_buffer:

                # 大摆渡首车
                if bus_b[1] == 0:
                    for bus in bus1:
                        # 按编号在资源池中找到该车辆
                        if bus[0] == bus_b[0]:
                            # 任务开始时间
                            start_time = task[5] - get_runtime(bus[2], place, runtime)
                            # 起始机坪
                            start_place = bus[2]
                            # 进港
                            if task[2] == '进港':
                                # 计算新增里程（行驶时间）
                                # 路程时间(航班机坪-3机坪)
                                time_run = get_runtime(task[3], '3', runtime)
                                # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                time_all = task[6] + 1.5 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[6] + 1.5 * time_run + 1
                                # 修改所在机坪
                                bus[2] = '3'
                            # 出港
                            else:
                                # 计算新增里程（行驶时间）
                                # 路程时间(登机口-航班机坪)
                                time_run = get_runtime(task[4], task[3], runtime)
                                # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                time_all = task[6] + 2 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[6] + 2 * time_run + 1
                                # 修改所在机坪
                                bus[2] = task[3]
                                # 遣回列表更新
                                bus_done.append([bus[0], bus[2], bus[3], task[0]])

                            # 更新调度列表
                            dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[5], bus[3], task[2], task[11], 1, task[4], task[12]])

                            break

                # 大摆渡次车
                elif bus_b[1] == 1:
                    for bus in bus1:
                        # 按编号在资源池中找到该车辆
                        if bus[0] == bus_b[0]:
                            # 任务开始时间
                            start_time = task[7] - get_runtime(bus[2], place, runtime)
                            # 起始机坪
                            start_place = bus[2]
                            # 进港
                            if task[2] == '进港':
                                # 计算新增里程（行驶时间）
                                # 路程时间(航班机坪-3机坪)
                                time_run = get_runtime(task[3], '3', runtime)
                                # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                time_all = task[8] + 1.5 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[8] + 1.5 * time_run + 1
                                # 修改所在机坪
                                bus[2] = '3'
                            # 出港
                            else:
                                # 计算新增里程（行驶时间）
                                # 路程时间(登机口-航班机坪)
                                time_run = get_runtime(task[4], task[3], runtime)
                                # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                time_all = task[8] + 2 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[8] + 2 * time_run + 1
                                # 修改所在机坪
                                bus[2] = task[3]
                                # 遣回列表更新
                                bus_done.append([bus[0], bus[2], bus[3], task[0]])

                            # 更新调度列表
                            dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[7], bus[3], task[2], task[11], 2, task[4], task[12]])

                            break

                # 大摆渡尾车
                elif bus_b[1] == 2:
                    for bus in bus1:
                        # 按编号在资源池中找到该车辆
                        if bus[0] == bus_b[0]:
                            # 任务开始时间
                            start_time = task[9] - get_runtime(bus[2], place, runtime)
                            # 起始机坪
                            start_place = bus[2]
                            # 进港
                            if task[2] == '进港':
                                # 计算新增里程（行驶时间）
                                # 路程时间(航班机坪-3机坪)
                                time_run = get_runtime(task[3], '3', runtime)
                                # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                time_all = task[10] + 1.5 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[10] + 1.5 * time_run + 1
                                # 修改所在机坪
                                bus[2] = '3'
                            # 出港
                            else:
                                # 计算新增里程（行驶时间）
                                # 路程时间(登机口-航班机坪)
                                time_run = get_runtime(task[4], task[3], runtime)
                                # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                time_all = task[10] + 2 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[10] + 2 * time_run + 1
                                # 修改所在机坪
                                bus[2] = task[3]
                                # 遣回列表更新
                                bus_done.append([bus[0], bus[2], bus[3], task[0]])

                            # 更新调度列表
                            dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[9], bus[3], task[2], task[11], 3, task[4], task[12]])

                            break

                # 大VIP
                else:
                    for bus in bus2:
                        # 按编号在资源池中找到该车辆
                        if bus[0] == bus_b[0]:
                            # 任务开始时间
                            start_time = task[5] - get_runtime(bus[2], place, runtime)
                            # 起始机坪
                            start_place = bus[2]
                            # 进港
                            if task[2] == '进港':
                                # 计算新增里程（行驶时间）
                                # 路程时间(航班机坪-3机坪)
                                time_run = get_runtime(task[3], '3', runtime)
                                # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                time_all = task[6] + 1.5 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[6] + 1.5 * time_run + 1
                                # 修改所在机坪
                                bus[2] = '3'
                            # 出港
                            else:
                                # 计算新增里程（行驶时间）
                                # 路程时间(登机口-航班机坪)
                                time_run = get_runtime(task[4], task[3], runtime)
                                # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                time_all = task[6] + 10 + 2 * time_run + 1 - start_time
                                # 修改里程
                                bus[1] = bus[1] + time_all
                                # 修改可用时间
                                bus[3] = task[6] + 10 + 2 * time_run + 1
                                # 修改所在机坪
                                bus[2] = task[3]
                                # 遣回列表更新
                                bus_done.append([bus[0], bus[2], bus[3], task[0]])

                            # 更新调度列表
                            dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[5], bus[3], task[2], task[11], 0, task[4], task[12]])

                            break

            # 缓冲区车辆信息同步完毕，清空缓冲区
            bus_buffer.clear()

            ti = 0
            # 当前任务移出任务池
            for t in task_todo:
                if t[0] == task[0]:
                    task_todo.pop(ti)
                    break
                ti += 1

        # 资源不够
        else:

            if len(bus_buffer) < 3:
                ok = 0
                b1 = 3 - len(bus_buffer)
                b2 = 0
            else:
                ok = 0
                b1 = 0
                b2 = 1
            if b1 == 0 and b2 == 0:
                st_lack = 1

        if len(task_todo) == 0:
            ok = 1
            b1 = 0
            b2 = 0

    # return ok,b1,b2
    return ok, b1, b2, dispatch_bus

def get_need2(n_bus1, n_bus2, loc_airline, loc_gate, n1, n2,excel_loc_runtime):
    '''
        计算资源缺口2（满足进港航班不延误）
        返回值依次为大摆渡缺口，大VIP缺口，调度结果，0-8员工缺口，8-16员工缺口，16-24员工缺口
    '''
    
    task_todo = get_task(load_airline(loc_airline, loc_gate))
    
    # 距离矩阵
    runtime = load_runtime(excel_loc_runtime)

    n1 = n1 + n_bus1
    n2 = n2 + n_bus2


    ok = 1
    while ok == 1:
        n1 -= 1
        bus1, bus2 = get_bus(n1, n2)
        ok, b1, b2, dispatch2 = run_need2(task_todo, bus1, bus2,runtime)
    n1 += 1

    ok = 1
    while ok == 1:
        n2 -= 1
        bus1, bus2 = get_bus(n1, n2)
        ok, b1, b2, dispatch2 = run_need2(task_todo, bus1, bus2,runtime)
    n2 += 1
    
    bus1, bus2 = get_bus(n1, n2)
    ok, b1, b2, dispatch2 = run_need2(task_todo, bus1, bus2,runtime)

    if n1 < n_bus1:
        n1 = n_bus1
    if n2 < n_bus2:
        n2 = n_bus2
        
    dispatch2 = add_day(dispatch2)
    # return n1-n_bus1,n2-n_bus2
    return n1 - n_bus1, n2 - n_bus2, dispatch2

def run_need2(task_todo, bus1, bus2,runtime):
    '''
        考虑资源缺口2的调度（进港航班不延误）
        从待响应任务池依次派遣任务
        返回ok，b1，b2
        ok=0表示航班延误情况存在，ok=1表示资源充足，航班不延误
        b1代表大摆渡缺口，b2代表大VIP缺口
    '''

    ok = 0
    b1 = 0
    b2 = 0

    '''
        遣回列表初始化
        遣回列表用于在每一次任务调度前对派出车辆进行遣回
        列表中一个元素的构成为
            [车辆编号（int），所在机坪（string）,可用时间（int）,航班编号（string）]
    '''

    # 已派出且不在3机坪
    bus_done = []

    '''
        缓冲池初始化
        列表中一个元素的构成为
            [车辆编号（int），第几辆车（int）]
    '''

    # 缓冲池
    bus_buffer = []

    '''
        调度列表初始化
        列表中一个元素的构成为
            [航班编号（string），车辆编号（int），起始机坪（string），任务机坪（string），返回机坪（string），
            起始时刻（int），航班时刻（int），终止时刻（int）,进出港类型（string）,延误时间（int）,机位（string）]
    '''

    # 调度列表
    dispatch_bus = []

    # 已分配任务
    task_done = []

    # 延误任务
    task_delay = []

    n_bus1 = len(bus1)

    task_cp = task_todo.copy()
    while task_todo:
        for x in task_delay:
            while task_delay.count(x) > 1:
                del task_delay[task_delay.index(x)]
        # print(dispatch_bus)
        #x = 0
        flag0 = 0
        for task in task_todo:
            bus1 = sorted(bus1, key=(lambda x: (x[3], x[1])))
            bus2 = sorted(bus2, key=(lambda x: (x[3], x[1])))

            # print(len(dispatch_bus))
            # 按时间排序
            task_todo = sorted(task_todo, key=(lambda x: x[5]))

            # 任务起始地点，出港航班为登机口，进港航班为航班所在机坪
            if task[2] == '出港':
                place = task[4]
            else:
                place = task[3]

            #i = 0
            # 无任务车辆遣返
            # print(bus_done)
            if len(dispatch_bus) != 0:
                bu = []
                for i in range(len(bus1) + len(bus2)):
                    bu.append([])

                for dis in dispatch_bus:
                    # 航班号，起始机坪，结束机坪，结束时间，起始时间
                    bu[dis[1]].append([dis[0], dis[2], dis[4], dis[7], dis[5]])

                # 最后一次从3机坪派任务-索引
                bun = []
                for i in range(len(bus1) + len(bus2)):
                    bun.append([])
                bui = 0
                for b in bu:
                    bi = 0
                    for i in b:
                        if i[2] == '3':
                            break
                        bi += 1
                    bi -= 1
                    bun[bui].append(bi)
                    bui += 1

                # 8:00-11:40派出
                flagb = []
                for i in range(len(bus1) + len(bus2)):
                    flagb.append([])
                bi = 0

                for order in bun:

                    if len(bu[bi]) != 0:
                        if bu[bi][order[0]][4] % 1440 > 480 and bu[bi][order[0]][4] % 1440 < 700:
                            flagb[bi].append(1)
                        else:
                            flagb[bi].append(0)
                    else:
                        flagb[bi].append(0)
                    bi += 1

                for f in flagb:
                    if task[5] % 1440 > 770 and f == 1:
                        f.append(1)
                    else:
                        f.append(0)

            for bus_d in bus_done:
                f = 0
                if flagb[bus_d[0]][0] == 1 and flagb[bus_d[0]][1] == 1:
                    f = 1

                # 计算该车辆等待的时间
                time_wait = task[5] - bus_d[2]
                # 到达任务机坪
                time_go = get_runtime(bus_d[1], '3', runtime) + get_runtime('3', place, runtime)
                # 等待时间大于往返时间，遣回
                if time_go < time_wait or f == 1:
                    # 大摆渡
                    if bus_d[0] < n_bus1:
                        for bus in bus1:
                            # 资源池中找到该车辆
                            if bus[0] == bus_d[0]:

                                # 修改调度列表
                                for dis_bus in dispatch_bus:
                                    if dis_bus[1] == bus_d[0]:
                                        if dis_bus[0] == bus_d[3]:
                                            bus[1] = bus[1] + get_runtime(bus_d[1], '3', runtime)
                                            bus[2] = '3'
                                            bus[3] = bus[3] + get_runtime(bus_d[1], '3', runtime)
                                            dis_bus[4] = '3'
                                            dis_bus[7] = bus[3]
                                            i = 0
                                            # 移出遣回列表
                                            for d in bus_done:
                                                if d[0] == bus_d[0]:
                                                    bus_done.pop(i)
                                                i += 1
                                            break
                    # 大VIP
                    else:
                        for bus in bus2:
                            # 资源池中找到该车辆
                            if bus[0] == bus_d[0]:

                                # 修改调度列表
                                for dis_bus in dispatch_bus:
                                    if dis_bus[1] == bus_d[0]:
                                        if dis_bus[0] == bus_d[3]:
                                            bus[1] = bus[1] + get_runtime(bus_d[1], '3', runtime)
                                            bus[2] = '3'
                                            bus[3] = bus[3] + get_runtime(bus_d[1], '3', runtime)
                                            dis_bus[4] = '3'
                                            dis_bus[7] = bus[3]
                                            i = 0
                                            # 移出遣回列表
                                            for d in bus_done:
                                                if d[0] == bus_d[0]:
                                                    bus_done.pop(i)
                                                i += 1
                                            break

            # 车辆匹配-获取缓冲池
            bus_buffer.clear()
            bus_buffer = get_buffer(bus1, bus2, task, place, runtime, bus_done)

            # 有足够资源
            if len(bus_buffer) == 4:

                for bus_b in bus_buffer:
                    # 移出遣返列表
                    di = 0
                    for bus_d in bus_done:
                        if bus_b[0] == bus_d[0]:
                            bus_done.pop(di)
                            break
                        di += 1

                # print(bus_buffer)
                # 缓冲池顺序：大摆渡1，大摆渡2，大摆渡3，大VIP
                # 将缓冲池中的车辆进行相关更新
                for bus_b in bus_buffer:

                    # 大摆渡首车
                    if bus_b[1] == 0:
                        for bus in bus1:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[5] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[6] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[6] + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[5], bus[3], task[2], task[11], 1, task[4], task[12]])

                                break

                    # 大摆渡次车
                    elif bus_b[1] == 1:
                        for bus in bus1:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[7] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[8] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[8] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[8] + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[8] + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[7], bus[3], task[2], task[11], 2, task[4], task[12]])

                                break

                    # 大摆渡尾车
                    elif bus_b[1] == 2:
                        for bus in bus1:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[9] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[10] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[10] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[10] + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[10] + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[9], bus[3], task[2], task[11], 3, task[4], task[12]])

                                break

                    # 大VIP
                    else:
                        for bus in bus2:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[5] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[6] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[6] + 10 + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 10 + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[5], bus[3], task[2], task[11], 0, task[4], task[12]])

                                break

                # 缓冲区车辆信息同步完毕，清空缓冲区
                bus_buffer.clear()

                ti = 0
                # 当前任务移出任务池
                for t in task_todo:
                    if t[0] == task[0]:
                        task_todo.pop(ti)
                        break
                    ti += 1

            # 资源不够
            else:
                #flag0 = 0
                # x+=1
                # print(task[0],task[1],task[2])
                bus_buffer2 = []
                al_id = None

                # 延误航班
                # 进港
                if task[2] == '进港':
                    # 回溯
                    for d in dispatch_bus:
                        if d[8] == '出港':
                            # 航班编号
                            al_id = d[0]
                            # print(d[0],d[8])
                            break
                    if al_id is None:
                        #print('No depart before.')
                        # 延误该航班
                        ok = 0
                        if len(bus_buffer) < 3:
                            b1 = 1
                            b2 = 0
                        else:
                            b1 = 0
                            b2 = 1
                        return ok, b1, b2, dispatch_bus
                        # return ok,b1,b2
                        break

                    # 找到最近的出港航班
                    else:
                        bus_buffer2.clear()
                        for d in dispatch_bus:
                            if d[0] == al_id:
                                bus_buffer2.append([d[1]])

                        # 该资源释放后是否被再次调用
                        flag = 0
                        for d in dispatch_bus:
                            if d[0] != al_id:
                                for b in bus_buffer2:
                                    if b == d[1]:
                                        flag = 1
                                        break
                                if flag == 1:
                                    #print('The car is not free.')
                                    # 延误航班
                                    ok = 0
                                    if len(bus_buffer) < 3:
                                        b1 = 1
                                        b2 = 0
                                    else:
                                        b1 = 0
                                        b2 = 1
                                    return ok, b1, b2, dispatch_bus
                                    # return ok,b1,b2

                                    break  # break while
                            else:
                                break

                            if flag == 1:
                                break  # break for d

                        # 该航班资源释放后未被再次调用
                        if flag == 0:
                            flag2 = 0
                            # 判断每辆车能否按时到达
                            for b in bus_buffer2:
                                # 找到该车上一次任务
                                for d in dispatch_bus:
                                    if b[0] == d[1] and d[0] != al_id:
                                        # 计算到达时间
                                        time_arrive = d[7] + get_runtime(place, d[4], runtime)
                                        # 释放后无法按时到达
                                        if time_arrive > task[5]:
                                            # 延误
                                            ok = 0
                                            if len(bus_buffer) < 3:
                                                b1 = 1
                                                b2 = 0
                                            else:
                                                b1 = 0
                                                b2 = 1
                                            return ok, b1, b2, dispatch_bus
                                            # return ok,b1,b2
                                            flag2 = 1
                                        break  # break for d
                                if flag2 == 1:
                                    break  # break for b
                            # 可按时到达，释放资源
                            if flag2 == 0:
                                flag0 = 1
                                # 延误航班
                                task_delay.append([al_id])
                                # 修改任务池
                                for t in task_todo:
                                    if al_id == t[0]:
                                        t[11] = t[11] + 10
                                        t[1] = t[1] + 10
                                        t[5] = t[5] + 10
                                        t[6] = t[6] + 10
                                        t[7] = t[7] + 10
                                        t[8] = t[8] + 10
                                        t[9] = t[9] + 10
                                        t[10] = t[10] + 10
                                #print('Delay the last depart.')
                                # print(al_id)
                                # 在任务池中添加该航班
                                for t in task_cp:
                                    if t[0] == al_id:
                                        task_todo.append([t[0], t[1] + 10, t[2], t[3], t[4], t[5] + 10, t[6] + 10, t[7] + 10, t[8] + 10, t[9] + 10, t[10] + 10, t[11] + 10, t[12]])

                                # 在调度队列中删除该航班
                                num = 4
                                for n in range(num):
                                    di = 0
                                    for dis in dispatch_bus:
                                        if dis[0] == al_id:
                                            dispatch_bus.pop(di)
                                            break
                                        di += 1

                                # 修改资源池（车辆）
                                bus_new = []
                                # 找到前序任务
                                for b in bus_buffer2:
                                    for d in dispatch_bus:
                                        if d[1] == b[0]:
                                            bus_new.append([b[0], d[7] - d[5], d[4], d[7]])
                                            break

                                # 没有前序任务
                                if len(bus_new) == 0:
                                    #print('No before.')
                                    for b in bus_buffer2:
                                        # 大摆渡
                                        if b[0] < n_bus1:
                                            for bus in bus1:
                                                if b == bus[0]:
                                                    bus[1] = 0
                                                    bus[2] = '3'
                                                    bus[3] = 0
                                        # 大VIP
                                        else:
                                            for bus in bus2:
                                                if b == bus[0]:
                                                    bus[1] = 0
                                                    bus[2] = '3'
                                                    bus[3] = 0
                                # 有前序任务
                                else:
                                    # 修改资源池
                                    for new in bus_new:
                                        # 大摆渡
                                        if new[0] < n_bus1:
                                            for bus in bus1:
                                                if new[0] == bus[0]:
                                                    bus[1] = bus[1] - new[1]
                                                    bus[2] = new[2]
                                                    bus[3] = new[3]
                                        # 大VIP
                                        else:
                                            for bus in bus2:
                                                if new[0] == bus[0]:
                                                    bus[1] = bus[1] - new[1]
                                                    bus[2] = new[2]
                                                    bus[3] = new[3]
                                    bus_new.clear()

                # 出港
                else:
                    # 直接延误
                    task_delay.append([task[0]])
                    # 修改任务池
                    for t in task_todo:
                        if task[0] == t[0]:
                            t[11] = t[11] + 10
                            t[1] = t[1] + 10
                            t[5] = t[5] + 10
                            t[6] = t[6] + 10
                            t[7] = t[7] + 10
                            t[8] = t[8] + 10
                            t[9] = t[9] + 10
                            t[10] = t[10] + 10
                    # print('Delay.')

                bus_buffer2.clear()
                if flag0 == 1:
                    break
            break

    ok = 1
    b1 = 0
    b2 = 0

    return ok, b1, b2, dispatch_bus

def bus_out(n1,n2,dispatch,date):
    '''
        输出Excel文件
        车辆行驶时间统计.xls
    '''
    import datetime
    #获取日期
    date = str(date[0])+'/'+str(date[1])+'/'+str(date[2])+' '+str(date[3])+':'+str(date[4])+':'+str(date[5])
    date = datetime.datetime.strptime(date,'%Y/%m/%d %H:%M:%S')
    
    #获取天数(按起始时刻)
    last = dispatch[0][5]
    first = dispatch[len(dispatch)-1][5]
    if first < 0:
        days = int(last/1440) + 2
        date = date + datetime.timedelta(days=-1)
        dispatch = add_day(dispatch)
    else:
        days = int(last/1440) + 1
    
    #初始化
    bus = []
    for d in range(days):
        for i in range(n1+n2):
            #[编号，里程，天]
            bus.append([i,0,d+1])
    
    for dis in dispatch:
        #第几天
        day = int(dis[5]/1440) + 1
        #里程
        run = dis[7] - dis[5]
        for b in bus:
            if dis[1]==b[0] and b[2]==day:
                #里程+
                b[1] += run
    
    s = 0
    bus_output = []
    head = [['车辆编号','总行驶时间（分钟）']]
    bus_output +=head
    
    for b in bus:
        if b[0] < n1:
            if b[0] < 9:
                id = 'COBUS-0' + str(b[0]+1)
            else:
                id = 'COBUS-' + str(b[0]+1)
        else:
            if b[0]-n1 < 9:
                id = '北方-0' + str(b[0]-n1+1)
            else:
                id = '北方-' + str(b[0]-n1+1)
        if b[0]==0:
            info = date + datetime.timedelta(days=b[2]-1)
            info = info.strftime('%Y') + '年' + info.strftime('%m') + '月' + info.strftime('%d') + '日'
            bus_output += [[info,'']]
        if b[1]!=0:
            bus_output += [[id,b[1]]]
            s += b[1]
        
    end = [['总计',s]]
    bus_output += end
    
    file_name = '车辆行驶时间统计.xls'
    
    import xlwt
    book = xlwt.Workbook()  # 新建一个Excel
    sheet = book.add_sheet('sheet1')  # 新建一个sheet页
    
    row = 0
    for bu in bus_output:
        col = 0
        for b in bu:
            sheet.write(row, col, b)
            col+=1
        row+=1

    book.save(file_name)  # 保存内容
    
    
def add_day(dispatch):
    dispatch.reverse()
    if dispatch[0][5] < 0:
        for dis in dispatch:
            dis[5] += 1440
            dis[6] += 1440
            dis[7] += 1440
    dispatch.reverse()
    return dispatch
    




def ferrybus_dispatch(loc_airline, loc_gate, n_bus1, n_bus2, loc_staff,excel_loc_runtime):
    '''
        主要部分
        从待响应任务池依次派遣任务
        返回dispatch_bus为调度结果
    '''

    '''
        遣回列表初始化
        遣回列表用于在每一次任务调度前对派出车辆进行遣回
        列表中一个元素的构成为
            [车辆编号（int），所在机坪（string）,可用时间（int）,航班编号（string）]
    '''
    
    staff = load_staff(loc_staff)
    
    task_todo = get_task(load_airline(loc_airline, loc_gate))
    
    # 已派出且不在3机坪
    bus_done = []

    '''
        缓冲池初始化
        列表中一个元素的构成为
            [车辆编号（int），第几辆车（int）]
    '''

    # 缓冲池
    bus_buffer = []

    '''
        调度列表初始化
        列表中一个元素的构成为
            [航班编号（string），车辆编号（int），起始机坪（string），任务机坪（string），返回机坪（string），
            起始时刻（int），航班时刻（int），终止时刻（int）,进出港类型（string）,延误时间（int）,机位（string）]
    '''

    # 调度列表
    dispatch_bus = []

    # 已分配任务
    task_done = []

    # 延误任务
    task_delay = []

    # 距离矩阵
    runtime = load_runtime(excel_loc_runtime)
    
    s1, s2, s3 = get_num_staff(staff)
    min = s1
    if s2<min:
        min = s2
    if s3<min:
        min = s3
    if n_bus1 > int(min*3/4):
        n_bus1 = int(min*3/4)
    n_bus2 = int(n_bus1/3)

    bus1,bus2 = get_bus(n_bus1,n_bus2)

    task_cp = task_todo.copy()
    while task_todo:
        for x in task_delay:
            while task_delay.count(x) > 1:
                del task_delay[task_delay.index(x)]
        # print(dispatch_bus)
        #x = 0
        flag0 = 0
        for task in task_todo:
            #print(len(dispatch_bus)/4)
            #print(len(task_todo))
            bus1 = sorted(bus1, key=(lambda x: (x[3], x[1])))
            bus2 = sorted(bus2, key=(lambda x: (x[3], x[1])))
            # print(len(dispatch_bus))
            # 按时间排序
            task_todo = sorted(task_todo, key=(lambda x: x[5]))

            # 任务起始地点，出港航班为登机口，进港航班为航班所在机坪
            if task[2] == '出港':
                place = task[4]
            else:
                place = task[3]

            #i = 0
            # 无任务车辆遣返
            # print(bus_done)
            if len(dispatch_bus) != 0:
                bu = []
                for i in range(len(bus1) + len(bus2)):
                    bu.append([])

                for dis in dispatch_bus:
                    # 航班号，起始机坪，结束机坪，结束时间，起始时间
                    bu[dis[1]].append([dis[0], dis[2], dis[4], dis[7], dis[5]])

                # 最后一次从3机坪派任务-索引
                bun = []
                for i in range(len(bus1) + len(bus2)):
                    bun.append([])
                bui = 0
                for b in bu:
                    bi = 0
                    for i in b:
                        if i[2] == '3':
                            break
                        bi += 1
                    bi -= 1
                    bun[bui].append(bi)
                    bui += 1

                # 8:00-11:40派出
                flagb = []
                for i in range(len(bus1) + len(bus2)):
                    flagb.append([])
                bi = 0

                for order in bun:

                    if len(bu[bi]) != 0:
                        if bu[bi][order[0]][4] % 1440 > 480 and bu[bi][order[0]][4] % 1440 < 700:
                            flagb[bi].append(1)
                        else:
                            flagb[bi].append(0)
                    else:
                        flagb[bi].append(0)
                    bi += 1

                for f in flagb:
                    if task[5] % 1440 > 770 and f == 1:
                        f.append(1)
                    else:
                        f.append(0)

            for bus_d in bus_done:
                f = 0
                if flagb[bus_d[0]][0] == 1 and flagb[bus_d[0]][1] == 1:
                    f = 1

                # 计算该车辆等待的时间
                time_wait = task[5] - bus_d[2]
                # 到达任务机坪
                time_go = get_runtime(bus_d[1], '3', runtime) + get_runtime('3', place, runtime)
                # 等待时间大于往返时间，遣回
                if time_go < time_wait or f == 1:
                    # 大摆渡
                    if bus_d[0] < n_bus1:
                        for bus in bus1:
                            # 资源池中找到该车辆
                            if bus[0] == bus_d[0]:

                                # 修改调度列表
                                for dis_bus in dispatch_bus:
                                    if dis_bus[1] == bus_d[0]:
                                        if dis_bus[0] == bus_d[3]:
                                            bus[1] = bus[1] + get_runtime(bus_d[1], '3', runtime)
                                            bus[2] = '3'
                                            bus[3] = bus[3] + get_runtime(bus_d[1], '3', runtime)
                                            dis_bus[4] = '3'
                                            dis_bus[7] = bus[3]
                                            i = 0
                                            # 移出遣回列表
                                            for d in bus_done:
                                                if d[0] == bus_d[0]:
                                                    bus_done.pop(i)
                                                i += 1
                                            break
                    # 大VIP
                    else:
                        for bus in bus2:
                            # 资源池中找到该车辆
                            if bus[0] == bus_d[0]:

                                # 修改调度列表
                                for dis_bus in dispatch_bus:
                                    if dis_bus[1] == bus_d[0]:
                                        if dis_bus[0] == bus_d[3]:
                                            bus[1] = bus[1] + get_runtime(bus_d[1], '3', runtime)
                                            bus[2] = '3'
                                            bus[3] = bus[3] + get_runtime(bus_d[1], '3', runtime)
                                            dis_bus[4] = '3'
                                            dis_bus[7] = bus[3]
                                            i = 0
                                            # 移出遣回列表
                                            for d in bus_done:
                                                if d[0] == bus_d[0]:
                                                    bus_done.pop(i)
                                                i += 1
                                            break

            # 车辆匹配-获取缓冲池
            bus_buffer.clear()
            bus_buffer, st_lack, flagt = get_bus_buffer(bus1, bus2, task, place, runtime, s1, s2, s3, bus_done)

            # 有足够资源
            if len(bus_buffer) == 4:

                for bus_b in bus_buffer:
                    # 移出遣返列表
                    di = 0
                    for bus_d in bus_done:
                        if bus_b[0] == bus_d[0]:
                            bus_done.pop(di)
                            break
                        di += 1

                # print(bus_buffer)
                # 缓冲池顺序：大摆渡1，大摆渡2，大摆渡3，大VIP
                # 将缓冲池中的车辆进行相关更新
                for bus_b in bus_buffer:
                    # 大摆渡首车
                    if bus_b[1] == 0:
                        for bus in bus1:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[5] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[6] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[6] + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[5], bus[3], task[2], task[11], 1, task[4], task[12]])

                                break

                    # 大摆渡次车
                    elif bus_b[1] == 1:
                        for bus in bus1:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[7] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[8] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[8] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[8] + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[8] + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[7], bus[3], task[2], task[11], 2, task[4], task[12]])

                                break

                    # 大摆渡尾车
                    elif bus_b[1] == 2:
                        for bus in bus1:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[9] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[10] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[10] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[10] + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[10] + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[9], bus[3], task[2], task[11], 3, task[4], task[12]])

                                break

                    # 大VIP
                    else:
                        for bus in bus2:
                            # 按编号在资源池中找到该车辆
                            if bus[0] == bus_b[0]:
                                # 任务开始时间
                                start_time = task[5] - get_runtime(bus[2], place, runtime)
                                # 起始机坪
                                start_place = bus[2]
                                # 进港
                                if task[2] == '进港':
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(航班机坪-3机坪)
                                    time_run = get_runtime(task[3], '3', runtime)
                                    # 新增里程（结束时间=发车时间+1.5×路程时间+1分钟）
                                    time_all = task[6] + 1.5 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 1.5 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = '3'
                                # 出港
                                else:
                                    # 计算新增里程（行驶时间）
                                    # 路程时间(登机口-航班机坪)
                                    time_run = get_runtime(task[4], task[3], runtime)
                                    # 新增里程（结束时间=发车时间+2×路程时间+1分钟）
                                    time_all = task[6] + 10 + 2 * time_run + 1 - start_time
                                    # 修改里程
                                    bus[1] = bus[1] + time_all
                                    # 修改可用时间
                                    bus[3] = task[6] + 10 + 2 * time_run + 1
                                    # 修改所在机坪
                                    bus[2] = task[3]
                                    # 遣回列表更新
                                    bus_done.append([bus[0], bus[2], bus[3], task[0]])

                                # 更新调度列表
                                dispatch_bus.insert(0, [task[0], bus[0], start_place, place, bus[2], start_time, task[5], bus[3], task[2], task[11], 0, task[4], task[12]])

                                break

                # 缓冲区车辆信息同步完毕，清空缓冲区
                bus_buffer.clear()

                ti = 0
                # 当前任务移出任务池
                for t in task_todo:
                    if t[0] == task[0]:
                        task_todo.pop(ti)
                        break
                    ti += 1

            # 资源不够
            else:
                # 清空缓冲区
                bus_buffer.clear()
                #flag0 = 0
                # x+=1
                # print(task[0],task[1],task[2])
                bus_buffer2 = []
                al_id = None

                # 延误航班
                # 进港
                if task[2] == '进港':
                    # 回溯
                    for d in dispatch_bus:
                        if d[8] == '出港':
                            # 航班编号
                            al_id = d[0]
                            # print(d[0],d[8])
                            break
                    if al_id is None:
                        #print('No depart before.')
                        # 延误该航班
                        # 加入延误列表
                        task_delay.append([task[0]])
                        # 修改任务池
                        for t in task_todo:
                            if task[0] == t[0]:
                                t[11] = t[11] + 10
                                t[1] = t[1] + 10
                                t[5] = t[5] + 10
                                t[6] = t[6] + 10
                                t[7] = t[7] + 10
                                t[8] = t[8] + 10
                                t[9] = t[9] + 10
                                t[10] = t[10] + 10
                        break

                    # 找到最近的出港航班
                    else:
                        bus_buffer2.clear()
                        for d in dispatch_bus:
                            if d[0] == al_id:
                                bus_buffer2.append([d[1]])

                        # 该资源释放后是否被再次调用
                        flag = 0
                        for d in dispatch_bus:
                            if d[0] != al_id:
                                for b in bus_buffer2:
                                    if b == d[1]:
                                        flag = 1
                                        break
                                if flag == 1:
                                    #print('The car is not free.')
                                    # 延误航班
                                    task_delay.append([task[0]])
                                    # 修改任务池
                                    for t in task_todo:
                                        if task[0] == t[0]:
                                            t[11] = t[11] + 10
                                            t[1] = t[1] + 10
                                            t[5] = t[5] + 10
                                            t[6] = t[6] + 10
                                            t[7] = t[7] + 10
                                            t[8] = t[8] + 10
                                            t[9] = t[9] + 10
                                            t[10] = t[10] + 10
                                    break  # break while
                            else:
                                break

                            if flag == 1:
                                break  # break for d

                        # 该航班资源释放后未被再次调用
                        if flag == 0:
                            flag2 = 0
                            # 判断每辆车能否按时到达
                            for b in bus_buffer2:
                                # 找到该车上一次任务
                                for d in dispatch_bus:
                                    if b[0] == d[1] and d[0] != al_id:
                                        # 计算到达时间
                                        time_arrive = d[7] + get_runtime(place, d[4], runtime)
                                        # 释放后无法按时到达
                                        if time_arrive > task[5]:
                                            # 延误
                                            task_delay.append([task[0]])
                                            #print('Can not arrive on time.')
                                            # 修改任务池
                                            for t in task_todo:
                                                if task[0] == t[0]:
                                                    t[11] = t[11] + 10
                                                    t[1] = t[1] + 10
                                                    t[5] = t[5] + 10
                                                    t[6] = t[6] + 10
                                                    t[7] = t[7] + 10
                                                    t[8] = t[8] + 10
                                                    t[9] = t[9] + 10
                                                    t[10] = t[10] + 10
                                            flag2 = 1
                                        break  # break for d
                                if flag2 == 1:
                                    break  # break for b
                            # 可按时到达，释放资源
                            if flag2 == 0:
                                flag0 = 1
                                # 延误航班
                                task_delay.append([al_id])
                                #print('Delay the last depart.')
                                #print(al_id)
                                #print(task[0])
                                # 在任务池中添加该航班
                                for t in task_cp:
                                    if t[0] == al_id:
                                        t[11] = t[11] + 10
                                        t[1] = t[1] + 10
                                        t[5] = t[5] + 10
                                        t[6] = t[6] + 10
                                        t[7] = t[7] + 10
                                        t[8] = t[8] + 10
                                        t[9] = t[9] + 10
                                        t[10] = t[10] + 10
                                        task_todo.append(t)
                                        #print(t[11])
                                        #task_todo.append([t[0], task[1] + 10, t[2], t[3], t[4], task[5] + 10, task[6] + 10, task[7] + 10,
                                        #                  task[8] + 10, task[9] + 10, task[10] + 10, t[11] + 10, t[12]])

                                # 在调度队列中删除该航班
                                num = 4
                                for n in range(num):
                                    di = 0
                                    for dis in dispatch_bus:
                                        if dis[0] == al_id:
                                            dispatch_bus.pop(di)
                                            break
                                        di += 1

                                # 修改资源池（车辆）
                                bus_new = []
                                # 找到前序任务
                                for b in bus_buffer2:
                                    for d in dispatch_bus:
                                        if d[1] == b[0]:
                                            bus_new.append([b[0], d[7] - d[5], d[4], d[7]])
                                            break

                                # 没有前序任务
                                if len(bus_new) == 0:
                                    #print('No before.')
                                    for b in bus_buffer2:
                                        # 大摆渡
                                        if b[0] < n_bus1:
                                            for bus in bus1:
                                                if b == bus[0]:
                                                    bus[1] = 0
                                                    bus[2] = '3'
                                                    bus[3] = -100
                                        # 大VIP
                                        else:
                                            for bus in bus2:
                                                if b == bus[0]:
                                                    bus[1] = 0
                                                    bus[2] = '3'
                                                    bus[3] = -100
                                # 有前序任务
                                else:
                                    # 修改资源池
                                    for new in bus_new:
                                        # 大摆渡
                                        if new[0] < n_bus1:
                                            for bus in bus1:
                                                if new[0] == bus[0]:
                                                    bus[1] = bus[1] - new[1]
                                                    bus[2] = new[2]
                                                    bus[3] = new[3]
                                        # 大VIP
                                        else:
                                            for bus in bus2:
                                                if new[0] == bus[0]:
                                                    bus[1] = bus[1] - new[1]
                                                    bus[2] = new[2]
                                                    bus[3] = new[3]
                                    bus_new.clear()

                # 出港
                else:
                    
                    # 直接延误
                    task_delay.append([task[0]])
                    # 修改任务池
                    for t in task_todo:
                        if task[0] == t[0]:
                            t[11] = t[11] + 10
                            t[1] = t[1] + 10
                            t[5] = t[5] + 10
                            t[6] = t[6] + 10
                            t[7] = t[7] + 10
                            t[8] = t[8] + 10
                            t[9] = t[9] + 10
                            t[10] = t[10] + 10
                    #print('Delay.')

                bus_buffer2.clear()
                if flag0 == 1:
                    break
            break
            
    

    
    bus_out(len(bus1),len(bus2),dispatch_bus,get_start_date(loc_airline))
    
    dispatch = add_day(dispatch_bus)
    
    return dispatch,len(bus1),len(bus2)
    

if __name__ == '__main__':
    loc_airline = '航班数据.xlsx'
    loc_runtime = '机坪运行时间.xlsx'
    loc_gate = '候机楼登机口信息.xlsx'
    loc_staff = '人员信息.xlsx'

    # 大摆渡数量39
    n_bus1 = 26
    # 大VIP数量13
    n_bus2 = 10
    
    
    task_todo = get_task(load_airline(loc_airline, loc_gate))
    runtime = load_runtime(loc_runtime)
    air = load_airline(loc_airline, loc_gate)
    
    #选择算法：0：时间优先；1：距离优先
    type = 0
    #type = 1
    
    #大摆渡，n2：大VIP，dispatch1：调度结果
    #n1, n2, dispatch1 = get_need1(n_bus1, n_bus2, loc_airline, loc_gate,loc_runtime)
    
    # 进港航班不延误的资源缺口
    # m1：大摆渡，m2：大VIP，dispatch2：调度结果
    #m1, m2, dispatch2 = get_need2(n_bus1, n_bus2,loc_airline, loc_gate, n1, n2,loc_runtime)
    
    #调度结果dispatch_bus
    dispatch_bus,b1,b2 = ferrybus_dispatch(loc_airline, loc_gate, n_bus1, n_bus2, loc_staff,loc_runtime)
    
    #print(n1, n2)
    #print(m1, m2)