
# coding: utf-8
'''
M.run方法的执行移到了最下方的selectFrame类中
大摆渡数量 小摆渡数量 算法选择返回的参数已注明
对于结果展示：
    将结果文件命名为 结果菜单项名.xls 生成在同一文件夹即可点击结果菜单项展示
'''

import wx.grid as gridlib
import xlwt
import xlrd
import wx
import os
import main as M
import myico
from datetime import datetime
from xlrd import xldate_as_tuple

APP_EXIT = 0

result_list = [['test', 'test', 'test', 'test', 'test', 'test', 'test', 'test', 'test', 'test', 'test', 5, 'test'],
               ['test', 'test', 'test', 'test', 'test', 'test1', 'test', 'test', 'test', 'test', 'test', 6, 'test'],
               ['test', 'test', 'test', 'test', 'test', 'test1', 'test', 'test', 'test', 'test', 'test', 4, 'test']]


class selectFrame(wx.Frame):
    def __init__(self):
        list_1 = [u'window', u'rec']
        super().__init__(parent=None, title=u'参数选择', size=(500, 300))
        panel = wx.Panel(self, -1, size=(500, 300))
        self.selectBox_1 = wx.Choice(panel, -1, (300, 200), (200, 40), list_1)
        self.Show(True)


class mainFrame(wx.Frame):
    '''程序主窗口类，继承自wx.Frame'''

    def __init__(self):
        '''构造函数'''
        super().__init__(parent=None, title=u'机场调度', size=(1000, 800))
        self.SetIcon(myico.AppIcon.GetIcon())

        self.PATH_1 = u'航班数据.xlsx'
        self.PATH_2 = u'人员信息.xlsx'
        self.PATH_3 = u'候机楼登机口信息.xlsx'
        self.PATH_4 = u'机坪运行时间.xlsx'  # 机坪运行时间

        menubar = wx.MenuBar()  # 生成菜单栏
        '''文件菜单'''
        filemenu = wx.Menu()  # 生成一个菜单
        menuitemQuit = wx.MenuItem(filemenu, APP_EXIT, '退出')  # 生成一个菜单项
        self.Bind(wx.EVT_MENU, self.OnQuit, id=APP_EXIT)
        menuitem1_1 = wx.MenuItem(filemenu, 11, u'航班数据')
        self.Bind(wx.EVT_MENU, self.OnOpenFile_1, id=11)
        menuitem1_2 = wx.MenuItem(filemenu, 12, u'人员信息')
        self.Bind(wx.EVT_MENU, self.OnOpenFile_2, id=12)
        menuitem1_3 = wx.MenuItem(filemenu, 13, u'候机楼登机口信息')
        self.Bind(wx.EVT_MENU, self.OnOpenFile_3, id=13)
        menuitem1_4 = wx.MenuItem(filemenu, 14, u'机坪运行时间')
        self.Bind(wx.EVT_MENU, self.OnOpenFile_4, id=14)
        filemenu.Append(menuitem1_1)  # 把菜单项加入到菜单中
        filemenu.Append(menuitem1_2)
        filemenu.Append(menuitem1_3)
        filemenu.Append(menuitem1_4)
        filemenu.AppendSeparator()
        filemenu.Append(menuitemQuit)
        menubar.Append(filemenu, u'&文件(F)')  # 把菜单加入到菜单栏中

        ##########################################
        ##############修改处 曾####################
        ##########################################
        '''操作菜单'''
        actionmenu = wx.Menu()
        self.menuitem2_1 = wx.MenuItem(actionmenu, 21, u'生成排班表')
        self.Bind(wx.EVT_MENU, self.OnMain, id=21)
        actionmenu.Append(self.menuitem2_1)

        self.menuitem2_2 = wx.MenuItem(actionmenu, 22, u'计算缺少资源')
        self.Bind(wx.EVT_MENU, self.OnMain1, id=22)
        actionmenu.Append(self.menuitem2_2)
        self.PathListen()

        ##########################################
        ################修改处 曾##################
        ##########################################

        menubar.Append(actionmenu, u'&操作(A)')
        self.SetMenuBar(menubar)
        '''结果菜单'''
        resultmenu = wx.Menu()
        self.menuitem3_1 = wx.MenuItem(actionmenu, 31, u'摆渡车任务')
        self.Bind(wx.EVT_MENU, lambda event, PATH=u'摆渡车任务.xls': self.OnOpenResult(event, PATH), id=31)
        self.menuitem3_1.Enable(False)
        self.menuitem3_2 = wx.MenuItem(actionmenu, 32, u'车辆行驶时间统计')
        self.Bind(wx.EVT_MENU, lambda event, PATH=u'车辆行驶时间统计.xls': self.OnOpenResult(event, PATH), id=32)
        self.menuitem3_2.Enable(False)
        self.menuitem3_3 = wx.MenuItem(actionmenu, 33, u'派遣结果统计')
        self.Bind(wx.EVT_MENU, lambda event, PATH=u'派遣结果统计.xls': self.OnOpenResult(event, PATH), id=33)
        self.menuitem3_3.Enable(False)
        resultmenu.Append(self.menuitem3_1)
        resultmenu.Append(self.menuitem3_2)
        resultmenu.Append(self.menuitem3_3)
        menubar.Append(resultmenu, u'&结果(R)')

        self.panel = wx.Panel(self, -1, size=(1920, 1080))
        text1 = wx.StaticText(self.panel, -1, u'航班数据：')
        self.FileName_1 = wx.TextCtrl(self.panel, size=(70, 10), style=(wx.TE_READONLY))
        if os.path.exists(u'航班数据.xlsx'):
            self.FileName_1.SetValue(u'就绪')
        else:
            self.FileName_1.SetValue(u'文件不存在')
        text2 = wx.StaticText(self.panel, -1, u'人员信息：')
        self.FileName_2 = wx.TextCtrl(self.panel, size=(70, 10), style=(wx.TE_READONLY))
        if os.path.exists(u'人员信息.xlsx'):
            self.FileName_2.SetValue(u'就绪')
        else:
            self.FileName_2.SetValue(u'文件不存在')
        text3 = wx.StaticText(self.panel, -1, u'候机楼登机口信息：')
        self.FileName_3 = wx.TextCtrl(self.panel, size=(70, 10), style=(wx.TE_READONLY))
        if os.path.exists(u'候机楼登机口信息.xlsx'):
            self.FileName_3.SetValue(u'就绪')
        else:
            self.FileName_3.SetValue(u'文件不存在')
        text4 = wx.StaticText(self.panel, -1, u'机坪运行时间：')
        self.FileName_4 = wx.TextCtrl(self.panel, size=(70, 10), style=(wx.TE_READONLY))
        if os.path.exists(u'机坪运行时间.xlsx'):
            self.FileName_4.SetValue(u'就绪')
        else:
            self.FileName_4.SetValue(u'文件不存在')

        box1 = wx.BoxSizer()
        box1.Add(text1, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(self.FileName_1, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(text2, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(self.FileName_2, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(text3, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(self.FileName_3, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(text4, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(self.FileName_4, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)

        self.grid = gridlib.Grid(self.panel, size=(1920, 1080))
        self.grid.CreateGrid(1, 1)
        self.grid.SetColLabelValue(0, u'使用说明')
        self.grid.SetCellValue(0, 0, u'请将源文件以“航班数据.xlsx”“人员信息.xlsx”“候机楼登机口信息.xlsx”“机坪运行时间.xlsx”命名放在程序目录下')
        self.grid.AutoSize()
        self.grid.SetLabelBackgroundColour('white')
        box2 = wx.BoxSizer()
        box2.Add(self.grid, proportion=3, flag=wx.EXPAND | wx.ALL, border=0)
        v_box = wx.BoxSizer(wx.VERTICAL)
        v_box.Add(box1, proportion=0, flag=wx.ALIGN_LEFT, border=3)  # 随着窗口托大控件一直保持居中或居左
        v_box.Add(box2, proportion=0, flag=wx.ALIGN_LEFT, border=3)
        self.SetSizer(v_box)
        self.Show(True)

    def OnQuit(self, event):
        '''关闭文件'''
        self.Close()

    ##########################################
    ##############修改处 曾####################
    ##########################################

    def OnMain(self, event):
        # l1 = self.PATH_1
        # l2 = self.PATH_2
        # l3 = self.PATH_3
        # l4 = self.PATH_4
        # self.PATH_1 = ''  # 航班
        # self.PATH_2 = ''  # 人员
        # self.PATH_3 = ''  # 候机楼登机口信息
        # self.PATH_4 = ''  # 机坪运行时间
        # write_result(main(l1, l4, l3, l2))
        #M.run(l1, l4, l3, l2, method='window', p=5, n_bus1=26, n_bus2=10)
        # self.menuitem3_1.Enable(True)
        # self.menuitem3_2.Enable(True)
        # self.menuitem3_3.Enable(True)
        frame = selectFrame(self.menuitem3_1, self.menuitem3_2, self.menuitem3_3)

    def OnMain1(self, event):
        frame = selectFrame_2()

    def PathListen(self):
        '''监听是否文件都存在'''
        if os.path.exists(u'航班数据.xlsx') and os.path.exists(u'人员信息.xlsx') and os.path.exists(u'候机楼登机口信息.xlsx') and os.path.exists(u'机坪运行时间.xlsx'):
            self.menuitem2_1.Enable(True)
            self.menuitem2_2.Enable(True)
        else:
            self.menuitem2_1.Enable(False)
            self.menuitem2_2.Enable(False)

    ##########################################
    ##############修改处 曾####################
    ##########################################

    def OnOpenFile_1(self, event):
        '''
        文件选择预留代码
        wildcard = 'Excel files(*.xlsx)|*.xlsx'
        dialog = wx.FileDialog(None,'select',os.getcwd(),'',wildcard,wx.FD_OPEN)
        if dialog.ShowModal() == wx.ID_OK:
            self.FileName_1.SetValue(dialog.GetPath())
            self.PATH_1=self.FileName_1.GetValue()
            dialog.Destroy
        '''
        if os.path.exists(u'航班数据.xlsx'):
            self.ShowExcel(self.PATH_1)
            self.FileName_1.SetValue(u'就绪')
        else:
            xxx = wx.MessageDialog(parent=None,
                                   message=u'当前路径不存在 航班数据.xlsx',
                                   caption=u'错误')
            xxx.ShowModal()
            self.FileName_1.SetValue(u'文件不存在')

    def OnOpenFile_2(self, event):
        if os.path.exists(u'人员信息.xlsx'):
            self.ShowExcel(self.PATH_2)
            self.FileName_2.SetValue(u'就绪')
        else:
            xxx = wx.MessageDialog(parent=None,
                                   message=u'当前路径不存在 人员信息.xlsx',
                                   caption=u'错误')
            xxx.ShowModal()
            self.FileName_2.SetValue(u'文件不存在')

    def OnOpenFile_3(self, event):
        if os.path.exists(u'候机楼登机口信息.xlsx'):
            self.ShowExcel(self.PATH_3)
            self.FileName_3.SetValue(u'就绪')
        else:
            xxx = wx.MessageDialog(parent=None,
                                   message=u'当前路径不存在 候机楼登机口信息.xlsx',
                                   caption=u'错误')
            xxx.ShowModal()
            self.FileName_3.SetValue(u'文件不存在')

    def OnOpenFile_4(self, event):
        if os.path.exists(u'机坪运行时间.xlsx'):
            self.ShowExcel(self.PATH_4)
            self.FileName_4.SetValue(u'就绪')
        else:
            xxx = wx.MessageDialog(parent=None,
                                   message=u'当前路径不存在 机坪运行时间.xlsx',
                                   caption=u'错误')
            xxx.ShowModal()
            self.FileName_3.SetValue(u'文件不存在')

    def OnOpenResult(self, event, PATH):
        if os.path.exists(PATH):
            self.ShowExcel(PATH)
        else:
            print(u'文件不存在')

    def ShowExcel(self, PATH):
        '''展示Excel文件内容'''
        data = dataframe(read_excel(io=PATH))  # data.shape[1]列数；data.shape[0]行数
        self.RecreateGrid(data)
        '''遍历赋值'''
        for i in range(0, data.shape[0]):
            for j in range(0, data.shape[1]):
                if isinstance(data.iat(i, j), str):
                    self.grid.SetCellValue(i, j, data.iat(i, j))
                else:
                    if str(data.iat(i, j)) == 'nan':
                        self.grid.SetCellValue(i, j, '')
                    else:
                        self.grid.SetCellValue(i, j, str(data.iat(i, j)))
        '''重构表头'''
        count = 0
        for each in data.columns:
            strtemp = u'Unnamed: ' + str(count)
            if str(each) != strtemp:
                self.grid.SetColLabelValue(count, str(each))
            else:
                self.grid.SetColLabelValue(count, '')
            count += 1
        '''自动调整cell大小'''
        self.grid.AutoSize()
        self.PathListen()
        '''刷新界面'''
        self.Layout()

    def RecreateGrid(self, data):
        '''重构表格'''
        # 行数=self.grid.getNumberRows()
        # 列数=self.grid.getNumberCols()
        '''清除表格中内容'''
        self.grid.ClearGrid()
        '''重构行'''
        if self.grid.GetNumberRows() > data.shape[0]:
            self.grid.DeleteRows(numRows=self.grid.GetNumberRows() - data.shape[0])
        elif self.grid.GetNumberRows() < data.shape[0]:
            self.grid.AppendRows(numRows=data.shape[0] - self.grid.GetNumberRows())
        '''重构列'''
        if self.grid.GetNumberCols() > data.shape[1]:
            self.grid.DeleteCols(numCols=self.grid.GetNumberCols() - data.shape[1])
        elif self.grid.GetNumberCols() < data.shape[1]:
            self.grid.AppendCols(numCols=data.shape[1] - self.grid.GetNumberCols())


def read_excel(io):
    '''读Excel方法'''
    temp = xlrd.open_workbook(io)
    temp = temp.sheets()[0]
    return temp


class dataframe():
    '''对于xlrd的数据处理类'''
    columns = []
    shape = [-1, -1]

    def __init__(self, data):
        self.data = data
        self.shape[1] = data.ncols
        self.shape[0] = data.nrows - 1
        for i in range(data.ncols):
            self.columns.append(data.cell_value(0, i))

    def iat(self, x, y):
        ctype = self.data.cell(x + 1, y).ctype  # 表格的数据类型
        cell = self.data.cell_value(x + 1, y)
        if ctype == 2 and cell % 1 == 0:  # 如果是整形
            cell = int(cell)
        elif ctype == 3:  # 转成datetime对象
            date = datetime(*xldate_as_tuple(cell, 0))
            if cell % 1 == 0:
                cell = date.strftime('%Y-%m-%d')
            else:
                cell = date.strftime('%Y-%m-%d %H:%M:%S')
        elif ctype == 4:
            cell = True if cell == 1 else False
        return cell

    def __del__(self):
        del self.columns[:]


class write_result():
    '''创建输出Excel的类，由构造函数直接创建'''

    def __init__(self, data):
        self.data = data

        book1 = xlwt.Workbook(encoding='utf-8')
        self.sheet1 = book1.add_sheet(u'摆渡车任务')
        self.write_book1()
        book1.save(u'摆渡车任务.xls')

        book2 = xlwt.Workbook(encoding='utf-8')
        self.sheet2 = book2.add_sheet(u'车辆行驶时间统计')
        self.write_book2()
        book2.save(u'车辆行驶时间统计.xls')

        book3 = xlwt.Workbook(encoding='utf-8')
        self.sheet3 = book3.add_sheet(u'派遣结果统计')
        self.write_book3()
        book3.save(u'派遣结果统计.xls')

    def write_book1(self):
        self.write_row(self.sheet1, 0, [u'航班号', u'任务执行登机口', u'任务执行机位', u'车型', u'车辆编号', u'任务类型', u'发车车次', u'到位时间', u'发车时间', u'结束时间', u'任务时长', u'人员编号'])
        for i in range(1, len(self.data)):
            count = 0
            for j in range(0, len(self.data[i])):
                if j != 7:
                    self.sheet1.write(i, count, self.data[i][j])
                    count += 1

    def write_book2(self):
        result = []
        self.write_row(self.sheet2, 0, [u'车辆编号', u'总行驶时间（分钟）'])
        for each in self.data:
            mark = 0
            for each_res in result:
                if each_res[0] == each[5]:
                    each_res[1] += each[11]
                    mark += 1
            if mark == 0:
                result.append([each[5], each[11]])
        for i in range(len(result)):
            self.write_row(self.sheet2, i + 1, result[i])

    def write_book3(self):
        result = []  # 12            0              11
        self.write_row(self.sheet3, 0, [u'员工编号', u'分配任务数量', u'任务时长（小时）', u'上班时长（小时）', u'工时利用率', u'派遣任务编号', '说明：上班时长=上班时间-下班时间，工时利用率=任务时长/上班时长'])
        for each in self.data:
            mark = 0
            for each_res in result:
                if each_res[0] == each[12]:
                    each_res[1] += 1
                    each_res[2] += each[11]
                    each_res[4] = each_res[2] / each_res[3]
                    each_res[5] = each_res[5] + ',' + str(each[0])
                    mark += 1
            if mark == 0:
                result.append([each[12], 1, int(each[11]) / 60, 8, int(each[11]) / 60 / 8, str(each[0])])
        for i in range(len(result)):
            self.write_row(self.sheet3, i + 1, result[i])

    def write_row(self, sheet, row_count, row):
        for i in range(0, len(row)):
            sheet.write(row_count, i, row[i])


def main(l1, l4, l3, l2):
    return result_list


class selectFrame(wx.Frame):

    def __init__(self, x, y, z):
        self.PATH_1 = u'航班数据.xlsx'
        self.PATH_2 = u'人员信息.xlsx'
        self.PATH_3 = u'候机楼登机口信息.xlsx'
        self.PATH_4 = u'机坪运行时间.xlsx'
        '''结果菜单的三个按钮'''
        self.x = x
        self.y = y
        self.z = z

        list_2 = [u'贪心算法', u'自底向上算法', u'向前看算法', u'递归回溯算法']
        list_1 = [u'任务优先', u'距离优先']
        super().__init__(parent=None, title=u'输入车辆数目', size=(500, 300))
        panel = wx.Panel(self, -1, size=(500, 300))
        text_1 = wx.StaticText(panel, -1, u'大摆渡数量：', size=(100, 25))
        self.number_1 = wx.TextCtrl(panel, size=(100, 25))
        text_2 = wx.StaticText(panel, -1, u'VIP摆渡数量：', size=(100, 25))
        self.number_2 = wx.TextCtrl(panel, size=(100, 25))
        # text_3=wx.StaticText(panel, -1, u'车辆排班算法：',size=(100,25))
        # self.selectBox_1 = wx.Choice(panel,-1,(0,0),(100, 25),list_1)
        # text_4=wx.StaticText(panel, -1, u'派工算法：',size=(100,25))
        # self.selectBox_2 = wx.Choice(panel,-1,(0,0),(100, 25),list_2)
        btn = wx.Button(panel, 0, label=u'确定', size=(50, 20))
        self.Bind(wx.EVT_BUTTON, self.OnClick, id=0)

        box1 = wx.BoxSizer()
        box1.Add(text_1, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(self.number_1, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box2 = wx.BoxSizer()
        box2.Add(text_2, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box2.Add(self.number_2, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        # box3= wx.BoxSizer()
        # box3.Add(text_3, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        # box3.Add(self.selectBox_1, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        # box4= wx.BoxSizer()
        # box4.Add(text_4, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        # box4.Add(self.selectBox_2, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box5 = wx.BoxSizer()
        box5.Add(btn, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)

        v_box = wx.BoxSizer(wx.VERTICAL)
        v_box.Add(box1, proportion=0, flag=wx.ALIGN_LEFT, border=3)
        v_box.Add(box2, proportion=0, flag=wx.ALIGN_LEFT, border=3)
        # v_box.Add(box3, proportion=0, flag=wx.ALIGN_LEFT, border=3)
        # v_box.Add(box4, proportion=0, flag=wx.ALIGN_LEFT, border=3)
        v_box.Add(box5, proportion=0, flag=wx.ALIGN_LEFT, border=3)

        self.SetSizer(v_box)
        self.Show(True)

    def OnClick(self, event):

        number_1 = self.number_1.GetValue()  # 大摆渡数量
        number_2 = self.number_2.GetValue()  # 小摆渡数量
        # selection_1=self.selectBox_1.GetSelection()#车辆排班算法类型
        # selection_2=self.selectBox_2.GetSelection()#派工算法类型
        # print(number_1, 'xxx', number_2, 'xxx', type(number_1))
        '''
        两个选择返回值均为int
        索引为：      0             1             2              3
        list_2=[u'贪心算法',u'自底向上算法',u'向前看算法',u'递归回溯算法']
        list_1=[u'任务优先',u'距离优先']
        '''
        # if selection_1==-1 and selection_2==-1:
        #     xxx = wx.MessageDialog(parent=None,
        #                            message=u'未选择算法',
        #                            caption=u'错误')
        #     xxx.ShowModal()
        # elif selection_1==-1 and selection_2!=-1:
        #     xxx = wx.MessageDialog(parent=None,
        #                            message=u'未选择车辆排班算法',
        #                            caption=u'错误')
        #     xxx.ShowModal()
        # elif selection_2==-1 and selection_1!=-1:
        #     xxx = wx.MessageDialog(parent=None,
        #                            message=u'未选择派工算法',
        #                            caption=u'错误')
        #     xxx.ShowModal()
        if number_1 == '' or number_2 == '':
            dialog = wx.MessageDialog(parent=None,
                                      message=u'未输入摆渡车数量',
                                      caption=u'错误')
            dialog.ShowModal()
        else:
            l1 = self.PATH_1
            l2 = self.PATH_2
            l3 = self.PATH_3
            l4 = self.PATH_4

            # self.PATH_1 = ''  # 航班
            # self.PATH_2 = ''  # 人员
            # self.PATH_3 = ''  # 候机楼登机口信息
            # self.PATH_4 = ''  # 机坪运行时间

            number_1 = 26 if number_1 == '' else number_1
            number_2 = 26 if number_2 == '' else number_2
            number_1 = int(number_1)
            number_2 = int(number_2)

            M.run(l1, l4, l3, l2, method1=0, method2=1, p=3, n_bus1=number_1, n_bus2=number_2)
            self.x.Enable(True)
            self.y.Enable(True)
            self.z.Enable(True)
            self.Close()


class selectFrame_2(wx.Frame):

    def __init__(self):
        self.PATH_1 = u'航班数据.xlsx'
        self.PATH_2 = u'人员信息.xlsx'
        self.PATH_3 = u'候机楼登机口信息.xlsx'
        self.PATH_4 = u'机坪运行时间.xlsx'

        super().__init__(parent=None, title=u'输入车辆数目', size=(500, 300))
        panel = wx.Panel(self, -1, size=(500, 300))
        text_1 = wx.StaticText(panel, -1, u'大摆渡数量：', size=(100, 25))
        self.number_1 = wx.TextCtrl(panel, size=(100, 25))
        text_2 = wx.StaticText(panel, -1, u'VIP摆渡数量：', size=(100, 25))
        self.number_2 = wx.TextCtrl(panel, size=(100, 25))
        btn = wx.Button(panel, 0, label=u'确定', size=(50, 20))
        self.Bind(wx.EVT_BUTTON, self.OnClick, id=0)

        box1 = wx.BoxSizer()
        box1.Add(text_1, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box1.Add(self.number_1, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box2 = wx.BoxSizer()
        box2.Add(text_2, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box2.Add(self.number_2, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)
        box5 = wx.BoxSizer()
        box5.Add(btn, proportion=0, flag=wx.EXPAND | wx.ALL, border=3)

        v_box = wx.BoxSizer(wx.VERTICAL)
        v_box.Add(box1, proportion=0, flag=wx.ALIGN_LEFT, border=3)
        v_box.Add(box2, proportion=0, flag=wx.ALIGN_LEFT, border=3)
        v_box.Add(box5, proportion=0, flag=wx.ALIGN_LEFT, border=3)

        self.SetSizer(v_box)
        self.Show(True)

    def OnClick(self, event):

        number_1 = self.number_1.GetValue()  # 大摆渡数量
        number_2 = self.number_2.GetValue()  # 小摆渡数量
        if number_1 == '' or number_2 == '':
            dialog = wx.MessageDialog(parent=None,
                                      message=u'未输入摆渡车数量',
                                      caption=u'错误')
            dialog.ShowModal()
        else:
            l1 = self.PATH_1
            l2 = self.PATH_2
            l3 = self.PATH_3
            l4 = self.PATH_4

            # self.PATH_1 = ''  # 航班
            # self.PATH_2 = ''  # 人员
            # self.PATH_3 = ''  # 候机楼登机口信息
            # self.PATH_4 = ''  # 机坪运行时间

            '''
            将此处M.run改为计算缺少资源的函数
            M.resource(l1, l4, l3, l2)
            '''
            number_1 = 26 if number_1 == '' else number_1
            number_2 = 26 if number_2 == '' else number_2
            number_1 = int(number_1)
            number_2 = int(number_2)

            M.resource(l1, l4, l3, l2, n_bus1=number_1, n_bus2=number_2)
            # self.x.Enable(True)
            # self.y.Enable(True)
            # self.z.Enable(True)
            self.Close()


if __name__ == '__main__':
    app = wx.App()
    frame = mainFrame()
    app.MainLoop()
