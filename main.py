# File name：
# Description:
# Author: zeng
# Time: 2020/4/14
# Change Activity:
import BusDispatch as bd
import StaffDispatch as sd


def judge_max(dispatch, loc_staff):
    lack = [1000, 1000, 1000]
    for i in range(1, 21):
        n = sd.judge(dispatch, loc_staff, i)
        lack = [min(lack[x], n[x]) for x in range(3)]
    return lack


# 运行选择，回溯算法，时间窗算法
def run(loc_airline, loc_runtime, loc_gate, loc_staff, method1=0, method2=1, p=3, n_bus1=26, n_bus2=10):
    # n_bus1 大摆渡数量
    # n_bus2 大VIP数量

    dispatch_bus, n1, n2, = bd.ferrybus_dispatch(loc_airline, loc_gate, n_bus1, n_bus2, loc_staff, loc_runtime)
    sd.control(dispatch_bus, loc_staff, loc_airline, n1, n2, p, method=method2)


# 运行判断，保障进港不延误，需要车辆资源和人员资源，保障全部不延误，需要车辆资源和人员资源
def resource(loc_airline, loc_runtime, loc_gate, loc_staff, n_bus1=26, n_bus2=10):
    # n_bus1 大摆渡数量
    # n_bus2 大VIP数量

    # 所有航班不延误的资源缺口
    # n1：大摆渡，n2：大VIP，dispatch1：调度结果，s1:0-8员工，s2:8-16员工，s3:16-24员工
    n1, n2, dispatch1 = bd.get_need1(n_bus1, n_bus2, loc_airline, loc_gate, loc_runtime)
    # print('w')
    lack = judge_max(dispatch1, loc_staff)

    # lack1 = [s1, s2, s3]
    # r1 = [max(lack[i], lack1[i], 0) for i in range(3)]
    r_a = lack

    print('若要所有航班不延误，则额外需要')
    print('大摆渡:', n1)
    print('大VIP：', n2)
    print('0-8时间员工：', r_a[0])
    print('8-16时间员工：', r_a[1])
    print('16-24时间员工：', r_a[2])

    # 进港航班不延误的资源缺口
    # m1：大摆渡，m2：大VIP，dispatch2：调度结果，w1:0-8员工，w2:8-16员工，w3:16-24员工
    m1, m2, dispatch2 = bd.get_need2(n_bus1, n_bus2, loc_airline, loc_gate, n1, n2, loc_runtime)
    # print('w')
    lack = judge_max(dispatch2, loc_staff)

    # lack1 = [w1, w2, w3]
    # r1 = [max(lack[i], lack1[i], 0) for i in range(3)]
    r_b = lack

    print()
    print('若要进港航班不延误，则额外需要')
    print('大摆渡:', m1)
    print('大VIP：', m2)
    print('0-8时间员工：', r_b[0])
    print('8-16时间员工：', r_b[1])
    print('16-24时间员工：', r_b[2])

    with open('缺少资源.txt', 'w') as f:

        f.write('若要所有航班不延误，则额外需要\n')
        f.write('大摆渡:'+str(n1)+'\n')
        f.write('大VIP：'+str(n2)+'\n')
        f.write('0-8时间员工：'+str(r_a[0])+'\n')
        f.write('8-16时间员工：'+str(r_a[1])+'\n')
        f.write('16-24时间员工：'+str(r_a[2])+'\n')

        f.write('\n若要进港航班不延误，则额外需要\n')
        f.write('大摆渡:'+str(m1)+'\n')
        f.write('大VIP：'+str(m2)+'\n')
        f.write('0-8时间员工：'+str(r_b[0])+'\n')
        f.write('8-16时间员工：'+str(r_b[1])+'\n')
        f.write('16-24时间员工：'+str(r_b[2])+'\n')


if __name__ == '__main__':

    loc_airline = '航班数据.xlsx'
    loc_runtime = '机坪运行时间.xlsx'
    loc_gate = '候机楼登机口信息.xlsx'
    loc_staff = '人员信息.xlsx'

    # dispatch_bus = bd.ferrybus_dispatch(loc_airline, loc_gate, 26, 10, loc_staff, loc_runtime)
    # sd.judge(dispatch_bus, loc_staff)

    run(loc_airline, loc_runtime, loc_gate, loc_staff, method1=0, method2=1, p=3, n_bus1=26, n_bus2=10)