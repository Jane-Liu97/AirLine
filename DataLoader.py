# File name：DataLoader
# Description:
# Author: zeng
# Time: 2020/3/31
# Change Activity:

from re import match
from xlrd import open_workbook, xldate_as_tuple

port_index = ['3', '4', '5', '93', '95', 'N2', 'N1', '2', '6', '7', '8', 'W1']


def read_excel(loc, head=False):
    excel = open_workbook(loc)
    table = excel.sheet_by_index(0)
    data = [table.row_values(i) for i in range(table.nrows)]
    if head:
        return data
    else:
        return data[1::]


def time_to_int(time_str):
    '''
    将24小时表示的时间字符串转换为整形
    :param time_str: 24小时表示的时间，string，形如'时钟:分钟:秒钟'
    :return: 距离00:00:00的过去了多少分钟，int
    '''
    hour, minute, second = [int(i) for i in time_str.split(':')]
    return hour*60 + minute


def int_to_time(time_int):
    '''
    :param time_int: 距离第0天的00:00:00的过去了多少分钟，int
    :return: day：第几天；hour：第几小时；minute：第几分钟。int
    '''
    day = int(time_int/1440)
    hour = int((time_int - day*1440) / 60)
    minute = time_int - day*1440 - hour*60
    return day, hour, minute


def get_runtime(port1, port2, mat):
    '''
    输入两个机坪代号和运行时间矩阵，返回摆渡车在两者间的运行时间
    :param port1: 机坪代号, string
    :param port2: 机坪代号, string
    :param mat: 摆渡车机坪运行时间矩阵，可由load_runtime()得到
    :return: 运行时间，int

    例子：
        >>> r = get_runtime('N2', '3', load_runtime())
        >>> 15
    '''

    i = port_index.index(port1)
    j = port_index.index(port2)

    return mat[i][j]


def get_gate_port(gate, d_port):
    for i in d_port:
        if gate == i[0]:
            return i[1]
    return None


def load_runtime(loc):
    '''
    读取机坪运行时间，返回对称的运行时间矩阵列表，机坪顺序按port_index定义，port_index在文件开头预定义
    :param loc: excel文件路径，string，默认为excel_loc_runtime，在文件开头预定义
    :return: 对称的运行时间矩阵列表，list
    '''

    data_runtime = read_excel(loc)
    # data_runtime = data_runtime.fillna(value=0)
    # data_runtime = data_runtime.values

    data_runtime = data_runtime[1:-1]
    data_runtime = [i[1::] for i in data_runtime]

    for i in range(len(data_runtime)):
        for j in range(len(data_runtime[i])):
            if i > j:
                data_runtime[i][j] = data_runtime[j][i]
            data_runtime[i][j] = int(data_runtime[i][j])
    return data_runtime


def load_gate(loc):
    '''
    读取登机口信息
    :param loc: excel文件路径，string，默认为excel_loc_gate，在文件开头预定义
    :return: 登机口列表，list，列表中一个元素的构成为
        [登机口号（string），登机口所在机坪（string）]
    注：若登机口表中其他元素无用，将返回的数据结构换为字典
    '''
    data_gate = read_excel(loc)
    # data_gate = data_gate.values

    d_gate = []
    for item in data_gate:
        # 登机口号
        gate = item[0]
        # print(gate)
        gate = match(r'\w?\d\d', gate)
        gate = gate.group()
        # 机坪
        port = str(int(item[4]))

        # 将单个元素加入列表
        d_gate.append([gate, port])

    return d_gate


def load_staff(loc):
    '''
    读取人员列表
    :param loc: excel文件路径，string，默认为excel_loc_staff，在文件开头预定义
    :return: 人员信息列表，list，列表中一个元素的构成为
        [编号（string），工作开始时间（int），工作结束时间（int）]
    '''

    data_staff = read_excel(loc)
    # data_staff = data_staff.values

    d_staff = []
    for item in data_staff:

        # 编号
        num = item[0]

        # 工作时间
        time = item[2]
        res = match(r'(\d+)-(\d+)', time)
        time_start = int(res.group(1))
        time_end = int(res.group(2))

        # 将单个元素加入列表
        d_staff.append([num, time_start, time_end])

    return d_staff


def load_airline(loc, loc1):
    '''
        读取航班数据，并按时间顺序排列，返回列表中的时间以航班进出港类型确定，若为进港，取STA为时间，反之取STD
        :param loc: excel文件路径，string，默认为excel_loc_airline，在文件开头预定义
        :return: 航班信息列表，list，列表中一个元素的构成为
            [航班号（string），时间（string），航班进出港类型（string），机坪（string），
            登机口1（string），登机口2（string），登机口机坪（string），机位（string）]
    '''

    data_airline = read_excel(loc)
    # data_airline = data_airline.fillna(value='empty')
    # data_airline = data_airline.values

    # 加载登机口信息
    gate_info = load_gate(loc1)
    # print(gate_info)
    d_al = []
    for item in data_airline:

        # 航班号
        al_id = item[1]

        # 进出港类型
        al_type = item[8]

        # 机位
        al_pos = item[9]
        if type(al_pos) == float:
            al_pos = str(int(al_pos))

        # 机坪
        al_port = item[9]
        if type(al_port) == float:
            al_port = str(int(al_port))

        # print(item[9])
        res = match('2|3|4|5|6|7|8|M|93|95|N2|N1|W1', str(al_port))
        if res == None:
            print('无法识别的机坪：', al_port)
            continue
        al_port = res.group()
        if al_port == 'M':
            al_port = 'N2'

        # 时间和登机口
        al_time = None
        al_gate1 = None
        al_gate2 = None
        al_gate_port = None
        # print()
        if al_type == '进港':
            al_time = item[7]
            # 进港航班无登机口信息
        else:
            al_time = item[6]
            # 为出港航班添加登机口和登机口所在机坪信息
            al_gate1 = item[10] if item[10] != '' else None
            al_gate2 = item[11] if item[11] != '' else None
            al_gate_port1 = get_gate_port(al_gate1, gate_info)
            al_gate_port2 = get_gate_port(al_gate2, gate_info)
            al_gate_port = al_gate_port1 if al_gate_port1 is not None else al_gate_port2

        al_time = xldate_as_tuple(al_time, 0)  # (year,month,day,hour,minute,nearest_second）

        # 将单个元素加入列表
        single = [al_id, al_time, al_type, al_port, al_gate1, al_gate2, al_gate_port, al_pos]
        for x in range(len(single)):
            if single[x] == ' ':
                single[x] = None
        d_al.append(single)

    # 按时间排序
    d_al = sorted(d_al, key=(lambda x: x[1]))

    # 处理时间为整形
    day_pointer = d_al[0][1][2]
    day_counter = 0
    for item in d_al:
        date = item[1][2]
        hour = item[1][3]
        minute = item[1][4]
        if date != day_pointer:
            day_counter += 1
            day_pointer = date
        item[1] = minute + hour*60 + day_counter*1440

    return d_al


def get_start_date(loc):
    '''

    '''

    data_airline = read_excel(loc)

    d_al = []
    for item in data_airline:

        # 进出港类型
        al_type = item[8]

        # 时间
        al_time = None

        if al_type == '进港':
            al_time = item[7]
            # 进港航班无登机口信息
        else:
            al_time = item[6]
        al_time = xldate_as_tuple(al_time, 0)  # (year,month,day,hour,minute,nearest_second）

        # 将单个元素加入列表
        d_al.append(al_time)

    # 按时间排序
    d_al = sorted(d_al)

    return d_al[0]


# 代码测试
if __name__ == '__main__':

    excel_loc_airline = '航班数据.xlsx'
    excel_loc_runtime = '机坪运行时间.xlsx'
    execl_loc_gate = '候机楼登机口信息.xlsx'
    excel_loc_staff = '人员信息.xlsx'

    d_airline = load_airline(excel_loc_airline, execl_loc_gate)
    d_runtime = load_runtime(excel_loc_runtime)
    d_staff = load_staff(excel_loc_staff)
    d_gate = load_gate(execl_loc_gate)
    #
    # for i in d_gate:
    #     print(i)
    for i in d_airline:
        print(i)
    print(get_start_date(excel_loc_airline))  # (year,month,day,hour,minute,nearest_second）
