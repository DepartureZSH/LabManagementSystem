import configparser
import re
import sys
import logging
import pandas as pd
import numpy as np
import os
from PyQt5.QtCore import Qt, pyqtSignal, QModelIndex
from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtWidgets import *

from qt_material import QtStyleTools

from PandasModel import pandasModel
from SeatTableModel import SeatTableModel
from QRCode import QRTool
from AddLine import add_view
from ModifyWidget import modify_view
from Setting import setting
from quick_updater import UpdateTool


class Demo(QWidget, QtStyleTools):
    def __init__(self):
        super(Demo, self).__init__()
        self.resize(1400, 900)
        self.setWindowTitle("实验室人员管理")
        self.logger_init()
        self.conf_affirm()
        self.elements()
        self.Layout()
        self.Style()
        self.Parameters()

    ##################################################################################################
    # 初始化
    ##################################################################################################
    # 启动日志
    def logger_init(self):
        self.logger = logging.getLogger("logger")
        handler = logging.FileHandler(filename="debug.log")
        self.logger.setLevel(logging.DEBUG)
        handler.setLevel(logging.DEBUG)
        formatter = logging.Formatter("%(asctime)s %(name)s %(levelname)s %(message)s")
        handler.setFormatter(formatter)
        self.logger.addHandler(handler)

    # 配置文件检查
    def conf_affirm(self):
        try:
            setting.ConfigCheck()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 参数
    def Parameters(self):
        try:
            config = configparser.ConfigParser()
            config.read("configs/config.ini")
            self.OutputPath = config.get("Path", "Outputs")
            self.DataPath = config.get("Path", "Data")
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    ##################################################################################################
    # 组件库
    ##################################################################################################
    def elements(self):
        #################################################
        # 菜单界面
        #################################################
        self.menu = QWidget()
        self.menu_UI()
        #################################################
        # 表格界面
        #################################################
        self.MainTableView = QWidget()
        self.MainTableView_UI()
        #################################################
        # 操作界面
        #################################################
        self.StackView = QStackedWidget(self)
        self.StackView_UI()
        #################################################
        # 查找栏界面
        #################################################
        self.searchView = QWidget()
        self.searchView_UI()
        #################################################
        # 其他界面
        #################################################
        # 查询无结果视图
        empty_df = pd.DataFrame(["请添加工位图"], columns=["查询无结果"])
        self.empty_Model = SeatTableModel(empty_df)
        # 实验室门牌号添加视图
        self.InputDialog = QInputDialog()
        self.InputDialog_UI()
        #################################################

    def menu_UI(self):
        #################################################
        # elements
        #################################################
        self.LM_Button = QPushButton("实验室管理")
        self.LMS_Button = QPushButton("实验室人员管理")
        self.setting_Button = QPushButton("设置")
        self.exit_Button = QPushButton("退出")
        #################################################
        # connections
        #################################################
        # 实验室管理
        self.LM_Button.clicked.connect(self.LM_Show)
        # 实验室学生管理
        self.LMS_Button.clicked.connect(self.LMS_Show)
        # 设置
        self.setting_Button.clicked.connect(self.Setting)
        # 退出
        self.exit_Button.clicked.connect(self.Exit)
        #################################################
        # Layout
        #################################################
        v_layout = QVBoxLayout()
        v_layout.addWidget(self.LM_Button)
        v_layout.addWidget(self.LMS_Button)
        vSpacer = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        v_layout.addItem(vSpacer)
        v_layout.addWidget(self.setting_Button)
        v_layout.addWidget(self.exit_Button)
        self.menu.setLayout(v_layout)
        #################################################

    def MainTableView_UI(self):
        #################################################
        # elements
        #################################################
        Label = QLabel("实验室人员名单")
        Label.setFont(QFont("Roman times", 10, QFont.Bold))
        Label.setAlignment(Qt.AlignHCenter)
        # 表格
        self.MainTable = QTableView()
        self.MainModel = pandasModel()
        self.MainTable.setModel(self.MainModel)
        # 保存
        self.save_button = QPushButton('保存', self)
        # 另存为
        self.save_as_button = QPushButton('另存为', self)
        # 增
        self.add_line_button = QPushButton('新增', self)
        # 删
        self.remove_line_button = QPushButton('删除', self)
        #################################################
        # connections
        #################################################
        # 展示详细信息
        self.MainTable.clicked.connect(self.ShowInfo)
        # 保存
        self.save_button.clicked.connect(self.Save)
        # 另存为
        self.save_as_button.clicked.connect(self.Save_as)
        # 添加记录
        self.add_line_button.clicked.connect(self.AddLine_Show)
        # 删除一行
        self.remove_line_button.clicked.connect(self.Remove_filter)
        #################################################
        # Layout
        #################################################
        h_layout1 = QHBoxLayout()
        h_layout1.addWidget(self.save_button)
        h_layout1.addWidget(self.save_as_button)
        h_layout2 = QHBoxLayout()
        h_layout2.addWidget(self.add_line_button)
        h_layout2.addWidget(self.remove_line_button)
        v_layout = QVBoxLayout()
        v_layout.addWidget(Label)
        v_layout.addLayout(h_layout1)
        v_layout.addWidget(self.MainTable)
        v_layout.addLayout(h_layout2)
        self.MainTableView.setLayout(v_layout)
        #################################################

    def StackView_UI(self):
        #################################################
        # elements
        #################################################
        # 普通人员操作界面
        self.MemberView = QWidget()
        self.MemberView_UI()
        # 实验室管理人员操作界面
        self.LMView = QWidget()
        self.LMView_UI()
        #################################################
        # Layout
        #################################################
        self.StackView.addWidget(self.MemberView)
        self.StackView.addWidget(self.LMView)
        self.StackView.setCurrentIndex(0)
        #################################################

    def MemberView_UI(self):
        #################################################
        # elements
        #################################################
        header = self.MainModel.data.columns.values
        self.LineEdit = {}
        self.modify_button = {}
        Label1 = QLabel("详细信息")
        Label1.setFont(QFont("Roman times", 10, QFont.Bold))
        Label1.setAlignment(Qt.AlignHCenter)
        for i in range(len(header)):
            if i == 0:
                self.LineEdit[header[i]] = QLabel("")
            else:
                self.LineEdit[header[i]] = QLineEdit("")
                self.modify_button[header[i]] = QPushButton('修改', self)
        Label2 = QLabel("工位图")
        Label2.setFont(QFont("Roman times", 10, QFont.Bold))
        Label2.setAlignment(Qt.AlignHCenter)
        # 工位视图
        self.SeatTable = QTableView()
        self.SeatTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 打印
        self.print_Word_button = QPushButton('生成工牌打印文件', self)
        #################################################
        # connections
        #################################################
        self.modify_button[header[1]].clicked.connect(lambda: self.Modify(header[1]))
        self.modify_button[header[2]].clicked.connect(lambda: self.Modify(header[2]))
        self.modify_button[header[3]].clicked.connect(lambda: self.Modify(header[3]))
        self.modify_button[header[4]].clicked.connect(lambda: self.Modify(header[4]))
        self.modify_button[header[5]].clicked.connect(lambda: self.Modify(header[5]))
        self.modify_button[header[6]].clicked.connect(lambda: self.Modify(header[6]))
        self.modify_button[header[7]].clicked.connect(lambda: self.Modify(header[7]))
        self.modify_button[header[8]].clicked.connect(lambda: self.Modify(header[8]))
        # 通过工位图修改
        self.SeatTable.doubleClicked.connect(self.Seats_modify)
        self.print_Word_button.clicked.connect(self.print_Word)
        #################################################
        # Layout
        #################################################
        GridLayout = QGridLayout()
        for i in range(len(header)):
            each = header[i]
            if each == "序号":
                GridLayout.addWidget(QLabel(each), i + 1, 0)
                GridLayout.addWidget(self.LineEdit[each], i + 1, 1)
            else:
                GridLayout.addWidget(QLabel(each), i + 1, 0)
                GridLayout.addWidget(self.LineEdit[each], i + 1, 1)
                GridLayout.addWidget(self.modify_button[each], i + 1, 2)
            if each in ["实验室号", "实验室门牌号", "座位号", "工位位置X", "工位位置Y"]:
                self.modify_button[each].setDisabled(True)
        VLayout = QVBoxLayout()
        VLayout.addWidget(Label1)
        VLayout.addLayout(GridLayout)
        VLayout.addWidget(Label2)
        VLayout.addWidget(self.SeatTable)
        VLayout.addWidget(self.print_Word_button)
        self.MemberView.setLayout(VLayout)
        #################################################

    def LMView_UI(self):
        #################################################
        # elements
        #################################################
        self.LabSelect = QComboBox(self)
        self.AddLabButton = QPushButton("添加", self)
        self.DelLabButton = QPushButton("删除", self)
        self.RoomSelect = QComboBox(self)
        self.AddRoomButton = QPushButton("添加", self)
        self.DelRoomButton = QPushButton("删除", self)
        self.LMSaveButton = QPushButton("保存实验室信息", self)
        # 工位视图
        self.SeatTable1 = QTableView()
        self.SeatTable1.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 打印
        self.print_Word_button1 = QPushButton('生成工牌打印文件', self)
        # 一键更新
        self.quick_update_button = QPushButton('一键更新', self)
        #################################################
        # connections
        #################################################
        # 实验室选择框
        self.LabSelect.currentTextChanged.connect(self.RoomComboBox)
        # 添加实验室
        self.AddLabButton.clicked.connect(self.AddLab)
        # 删除实验室
        self.DelLabButton.clicked.connect(self.DelLab)
        # 实验室门牌号选择框
        self.RoomSelect.currentTextChanged.connect(self.RoomShow)
        # 添加实验室门牌
        self.AddRoomButton.clicked.connect(self.AddRoom)
        # 删除实验室门牌
        self.DelRoomButton.clicked.connect(self.DelRoom)
        # 保存实验室信息
        self.LMSaveButton.clicked.connect(self.LMSave)
        # 通过工位图修改
        self.SeatTable1.doubleClicked.connect(self.Seats_modify)
        # 打印
        self.print_Word_button1.clicked.connect(self.print_Word)
        # 一键更新
        self.quick_update_button.clicked.connect(self.quick_update)
        #################################################
        # Layout
        #################################################
        GridLayout = QGridLayout()
        GridLayout.addWidget(QLabel("实验室"), 0, 0)
        GridLayout.addWidget(self.LabSelect, 0, 1)
        GridLayout.addWidget(self.AddLabButton, 0, 2)
        GridLayout.addWidget(self.DelLabButton, 0, 3)
        GridLayout.addWidget(QLabel("实验室门牌号"), 1, 0)
        GridLayout.addWidget(self.RoomSelect, 1, 1)
        GridLayout.addWidget(self.AddRoomButton, 1, 2)
        GridLayout.addWidget(self.DelRoomButton, 1, 3)
        # 布局
        VLayout = QVBoxLayout()
        Label1 = QLabel("406实验室")
        Label1.setFont(QFont("Roman times", 10, QFont.Bold))
        Label1.setAlignment(Qt.AlignHCenter)
        VLayout.addWidget(Label1)
        VLayout.addLayout(GridLayout)
        Label2 = QLabel("工位图")
        Label2.setFont(QFont("Roman times", 10, QFont.Bold))
        Label2.setAlignment(Qt.AlignHCenter)
        VLayout.addWidget(Label2)
        VLayout.addWidget(self.SeatTable1)
        VLayout.addWidget(self.quick_update_button)
        VLayout.addWidget(self.print_Word_button1)
        VLayout.addWidget(self.LMSaveButton)
        self.LMView.setLayout(VLayout)
        #################################################

    def searchView_UI(self):
        #################################################
        # elements
        #################################################
        self.search_label = []
        self.search_line = []
        for each in self.MainModel.DataDisplay.columns.values:
            self.search_label.append(QLabel(each))
            self.search_line.append(QLineEdit(self))
        self.reg_flag = False
        self.reg_checkbox = QCheckBox('正则匹配', self)
        self.search_button = QPushButton('搜索', self)
        self.back_button = QPushButton('返回', self)
        #################################################
        # connections
        #################################################
        # 正则选择
        self.reg_checkbox.clicked.connect(self.Reg)
        # 搜索
        self.search_button.clicked.connect(self.Search)
        # 返回
        self.back_button.clicked.connect(self.Back)
        #################################################
        # Layout
        #################################################
        HLayout = QHBoxLayout()
        for i in range(len(self.search_line)):
            HLayout.addWidget(self.search_label[i])
            HLayout.addWidget(self.search_line[i])
        HLayout.addWidget(self.reg_checkbox)
        HLayout.addWidget(self.search_button)
        HLayout.addWidget(self.back_button)
        self.searchView.setLayout(HLayout)
        #################################################

    def InputDialog_UI(self):
        #################################################
        # elements
        #################################################
        self.InputDialog.setWindowTitle("添加实验室门牌号")
        self.InputDialog.setLabelText("请输入实验室门牌号:")
        self.InputDialog.setOkButtonText(u"确定")
        self.InputDialog.setCancelButtonText(u"取消")
        #################################################
        # Layout
        #################################################
        self.InputDialog.setFixedSize(350, 400)
        self.InputDialog.setWindowFlags(Qt.WindowStaysOnTopHint)
        #################################################

    def Layout(self):
        HLayout1 = QHBoxLayout()
        VLayout1 = QVBoxLayout()
        HLayout = QHBoxLayout()

        HLayout1.addWidget(self.StackView)
        HLayout1.addWidget(self.MainTableView)

        VLayout1.addLayout(HLayout1)
        VLayout1.addWidget(self.searchView)

        HLayout.addWidget(self.menu)
        HLayout.addLayout(VLayout1)

        self.setLayout(HLayout)

    ##################################################################################################
    # 槽函数
    ##################################################################################################

    # 1、menu
    # 保存
    def Save(self):
        try:
            self.MainModel.PrimarySave()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 另存为
    def Save_as(self):
        try:
            path, _ = QFileDialog.getSaveFileName(self, '保存', self.OutputPath, 'Files (*.xlsx)')
            self.MainModel.CustomizedSave(path)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 设置
    def Setting(self):
        try:
            self.settingView = setting()
            self.settingView.show()
            self.settingView.signal_UpdateSetting.connect(self.StyleChange)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 退出
    def Exit(self):
        try:
            res = QMessageBox.question(self, "确认退出?", "请先保存后退出。", QMessageBox.Yes | QMessageBox.No)
            if res == QMessageBox.Yes:
                self.close()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    ##################################################################################################

    # 2、MainTableView
    # 添加一行组件视图
    def AddLine_Show(self):
        try:
            if self.StackView.currentIndex() == 0:
                QMessageBox.information(self, "提示", "该功能只能由管理人员操作", QMessageBox.Yes)
            elif self.StackView.currentIndex() == 1:
                try:
                    self.add_view = add_view(self.MainModel.df, self.MainModel.df3)
                    self.add_view.show()
                    self.add_view.signal_addLine.connect(self.Insert)
                except Exception as e:
                    self.logger.debug(
                        "{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 增
    def Insert(self, data):
        try:
            self.MainModel.temp = pd.DataFrame([data])
            self.MainModel.insertRows(self.MainModel.data.shape[0])
            if self.StackView.currentIndex() == 0:
                self.ShowInfo()
            elif self.StackView.currentIndex() == 1:
                self.RoomSeatShow()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 删除-过滤器
    def Remove_filter(self):
        try:
            if self.StackView.currentIndex() == 0:
                self.Remove(self.index)
            elif self.StackView.currentIndex() == 1:
                if self.MainTable.currentIndex().row() == -1:
                    if self.SeatTable1.currentIndex().row() == -1:
                        QMessageBox.information(self, '提示', '请选择一位学生再进行此操作', QMessageBox.Yes)
                        return
                    else:
                        row = self.SeatTable1.currentIndex().row()
                        col = self.SeatTable1.currentIndex().column()
                        SeatNum = self.SeatModel.DataDisplay.iloc[row, col]
                        if str(SeatNum) == 'nan':
                            QMessageBox.information(self, '提示', '请选择一位学生再进行此操作', QMessageBox.Yes)
                            return
                        index = self.MainModel.data[self.MainModel.data["座位号"] == SeatNum].index.values
                        self.Remove(index[0])
                else:
                    index = self.MainTable.currentIndex().row()
                    self.Remove(index)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 删除
    def Remove(self, index):
        try:
            for each in self.MainModel.data.columns.values:
                if each == "序号" or each == "座位号" or each == "实验室门牌号" or each == "实验室号":
                    pass
                else:
                    try:
                        Type = type(self.MainModel.data.loc[index, each])
                        if Type == str:
                            Input = ''
                        elif Type == np.int64:
                            Input = 0
                        else:
                            Input = np.NAN
                        row = self.MainModel.data.loc[index, "序号"]
                        loc = self.MainModel.df[self.MainModel.df["序号"] == row].index.values
                        self.MainModel.data.loc[index, each] = Input
                        self.MainModel.data[each].astype(Type)
                        self.MainModel.pre_data = self.MainModel.data[["实验室号", "座位号", "姓名", "学号", "导师"]]
                        self.MainModel.beginResetModel()
                        self.MainModel.DataDisplay = self.MainModel.pre_data
                        self.MainModel.endResetModel()
                        self.MainModel.df.loc[loc, each] = Input
                        self.MainModel.df[each].astype(Type)
                    except Exception as e:
                        self.logger.debug(
                            "{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 普通修改
    def Modify(self, header):
        try:
            if self.index:
                pass
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
            return
        try:
            Input = self.LineEdit[header].text()
            if Input == '':
                raise Exception("请输入{}".format(header))
            if header == '姓名':
                if re.match(r"^(?:[\u4e00-\u9fa5]+)(?:●[\u4e00-\u9fa5]+)*$|^[a-zA-Z0-9]+\s?[\.·\-()a-zA-Z]*[a-zA-Z]+$",
                            Input):
                    pass
                else:
                    raise Exception("请输入正确的姓名（中文名或英文名）")
            if header == '学号':
                if re.match(r"^(S[0-9]{9})*$|^(B[0-9]{9})*$", Input):
                    pass
                else:
                    raise Exception("请输入正确的学号（S或B+9位数字）")
            if header == '导师':
                if re.match(r"^(?:[\u4e00-\u9fa5]+)(?:●[\u4e00-\u9fa5]+)*$|^[a-zA-Z0-9]+\s?[\.·\-()a-zA-Z]*[a-zA-Z]+$",
                            Input):
                    pass
                else:
                    raise Exception("请输入正确的导师姓名（中文名或英文名）")
            if header == '实验室门牌号':
                if re.match(r"^[0-9]+$", Input):
                    Input = int(Input)
                    # 无此实验室异常
                else:
                    raise Exception("请输入正确的实验室门牌号（数字）")
            if header == '座位号':
                if re.match(r"^[0-9]{1,2}$", Input):
                    Input = int(Input)
                    # 被占用异常
                else:
                    raise Exception("请输入正确的座位号（正整数）")
            if header == '工位位置X':
                if re.match(r"^[0-9]{1,2}$", Input):
                    Input = float(Input)
                else:
                    raise Exception("请输入正确的工位位置X（整数）")
            if header == '工位位置Y':
                if re.match(r"^[0-9]{1,2}$", Input):
                    Input = float(Input)
                    # 被占用异常
                else:
                    raise Exception("请输入正确的工位位置Y（整数）")
            if header == '实验室号':
                if re.match(r"^[0-9]+$", Input):
                    Input = int(Input)
                else:
                    raise Exception("请输入正确的实验室号（数字）")
            row = self.MainModel.data.loc[self.index, "序号"]
            loc = self.MainModel.df[self.MainModel.df["序号"] == row].index.values
            self.MainModel.data.loc[self.index, header] = Input
            self.MainModel.pre_data = self.MainModel.data[["实验室号", "座位号", "姓名", "学号", "导师"]]
            self.MainModel.beginResetModel()
            self.MainModel.DataDisplay = self.MainModel.pre_data
            self.MainModel.endResetModel()
            self.MainModel.df.loc[loc, header] = Input
            self.LineEdit[header].setText("")
            self.LineEdit[header].setPlaceholderText(str(Input))

        except Exception as e:
            QMessageBox.information(self, '提示', '{}'.format(str(e)), QMessageBox.Yes)

    ##################################################################################################

    # 3、StackView
    ##################################################################################################
    # 实验室人员管理切换
    def LMS_Show(self):
        if self.StackView.currentIndex() == 0:
            return
        self.StackView.setCurrentIndex(0)
        self.MainModel.dataUpdate()
        try:
            self.back_button.clicked.disconnect(self.RoomStaffShow)
            self.back_button.clicked.connect(self.Back)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 实验室管理切换
    def LM_Show(self):
        if self.StackView.currentIndex() == 1:
            return
        self.StackView.setCurrentIndex(1)
        self.LabComboBox()
        try:
            self.back_button.clicked.disconnect(self.Back)
            self.back_button.clicked.connect(self.RoomStaffShow)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    ##################################################################################################

    # 4、MemberView
    # 详细信息展示
    def ShowInfo(self):
        self.index = self.MainTable.currentIndex().row()
        Info = {}
        try:
            for i in range(len(self.MainModel.data.columns.values)):
                Info[self.MainModel.data.columns.values[i]] = self.MainModel.data.iloc[self.index, i]
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
        try:
            if str(Info['实验室门牌号']) == 'nan':
                raise Exception("实验室门牌号 == nan")
            else:
                df = pd.read_excel("data/{}.xlsx".format(int(Info['实验室门牌号'])), header=None)
                df2 = self.MainModel.data[self.MainModel.data["实验室门牌号"] == int(Info['实验室门牌号'])]
                self.Seats_Update(df, df2)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
            self.SeatTable.setModel(self.empty_Model)
        try:
            for each in self.MainModel.data.columns.values:
                if each == "序号":
                    self.LineEdit[each].setText(str(Info[each]))
                else:
                    if str(Info[each]) == 'nan':
                        self.LineEdit[each].setPlaceholderText("")
                    elif type(Info[each]) == np.float64 or type(Info[each]) == float:
                        self.LineEdit[each].setPlaceholderText(str(int(Info[each])))
                    else:
                        self.LineEdit[each].setPlaceholderText(str(Info[each]))
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    ##################################################################################################

    # 5、LMView
    # 添加实验室
    def AddLab(self):
        try:
            num = str(self.LabSelect.count() + 1)
            self.LabSelect.addItem(num)
            QMessageBox.information(self, "提示", "实验室{}已添加\n请添加实验室门牌并保存使此操作生效".format(num), QMessageBox.Yes)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 删除实验室
    def DelLab(self):
        try:
            Lab = int(self.LabSelect.currentText())
            res = self.MainModel.df3[self.MainModel.df3["实验室号"] == Lab]
            info1 = "该实验室号包含{}个实验室门牌号".format(res.shape[0])
            choice1 = QMessageBox.question(self, "确定删除", info1, QMessageBox.Yes | QMessageBox.No)
            if choice1 == QMessageBox.Yes:
                staff = self.MainModel.df[self.MainModel.df["实验室号"] == Lab]
                if staff.shape[0]:
                    info2 = "该实验室号还拥有{}个学生\n建议先删除或迁移该实验室所有人员后进行此操作".format(staff.shape[0])
                    choice2 = QMessageBox.information(self, "建议", info2, QMessageBox.Yes | QMessageBox.No)
                    if choice2 == QMessageBox.Yes:
                        rows = res.index.values
                        for row_index in rows:
                            res = self.MainModel.df3.loc[row_index]
                            self.MainModel.df3 = self.MainModel.df3.drop(labels=row_index, axis=0)
                            self.MainModel.AdvancedUpdate()
                        self.LabComboBox()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 添加实验室门牌号
    def AddRoom(self):
        try:
            if self.InputDialog.exec_():
                Input = self.InputDialog.textValue()
                if re.match(r"^[0-9]+$", Input):
                    res = self.MainModel.df3[self.MainModel.df3["实验室门牌号"] == int(Input)]
                    if res.shape[0]:
                        QMessageBox.information(self, "提示", "该实验室门牌号已被实验室{}拥有".format(res["实验室号"].values[0]),
                                                QMessageBox.Yes)
                        return
                    self.RoomSelect.addItem(Input)
                    new_room = {
                        "实验室门牌号": int(Input),
                        "实验室号": int(self.LabSelect.currentText())
                    }
                    temp = pd.DataFrame([new_room])
                    self.MainModel.df3 = self.MainModel.df3.append(temp)
                    self.MainModel.AdvancedUpdate()
                else:
                    QMessageBox.information(self, "提示", "请输入正确格式的实验室门牌号", QMessageBox.Yes)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 删除实验室门牌号
    def DelRoom(self):
        try:
            Room = int(self.RoomSelect.currentText())
            staff = self.MainModel.df[self.MainModel.df["实验室门牌号"] == Room]
            info1 = "该实验室号还拥有{}个学生\n建议先删除或迁移该实验室所有人员后进行此操作".format(staff.shape[0])
            choice1 = QMessageBox.question(self, "确定删除", info1, QMessageBox.Yes | QMessageBox.No)
            if choice1 == QMessageBox.Yes:
                res = self.MainModel.df3[self.MainModel.df3["实验室门牌号"] == Room]
                rows = res.index.values
                for row_index in rows:
                    res = self.MainModel.df3.loc[row_index]
                    self.MainModel.df3 = self.MainModel.df3.drop(labels=row_index, axis=0)
                    self.MainModel.AdvancedUpdate()
                self.RoomComboBox()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 实验室保存
    def LMSave(self):
        try:
            res = self.MainModel.PrimarySave()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
        try:
            res = self.MainModel.AdvancedSave()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
        try:
            path = self.DataPath + self.RoomSelect.currentText() + ".xlsx"
            self.SeatModel.df.to_excel(path, sheet_name='Sheet1', index=False, header=None)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 实验室选择框
    def LabComboBox(self):
        try:
            try:
                self.LabSelect.currentTextChanged.disconnect(self.RoomComboBox)
            except:
                pass
            self.LabSelect.clear()
            self.LabSelect.currentTextChanged.connect(self.RoomComboBox)
            LabList = np.unique(self.MainModel.df3["实验室号"].values).tolist()
            self.LabSelect.addItems([str(each) for each in LabList])
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 实验室门牌号选择框
    def RoomComboBox(self):
        try:
            try:
                self.RoomSelect.currentTextChanged.disconnect(self.RoomShow)
            except:
                pass
            self.RoomSelect.clear()
            self.RoomSelect.currentTextChanged.connect(self.RoomShow)
            df = self.MainModel.df3.loc[self.MainModel.df3["实验室号"] == int(self.LabSelect.currentText())]
            RoomList = df["实验室门牌号"].tolist()
            self.RoomSelect.addItems([str(each) for each in RoomList])
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 实验室管理之人员视图
    def RoomStaffShow(self):
        if self.RoomSelect.currentText() == "":
            return
        try:
            df = self.MainModel.df.loc[self.MainModel.df["实验室门牌号"] == int(self.RoomSelect.currentText())]
            if not df.empty:
                self.MainModel.beginResetModel()
                df = df.reset_index(drop=True)
                self.MainModel.data = df
                self.MainModel.pre_data = df[["实验室号", "座位号", "姓名", "学号", "导师"]]
                self.MainModel.DataDisplay = self.MainModel.pre_data
                self.MainModel.endResetModel()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 实验室管理之工位视图
    def RoomSeatShow(self):
        if self.RoomSelect.currentText() == "":
            return
        try:
            df = pd.read_excel("data/{}.xlsx".format(self.RoomSelect.currentText()), header=None)
            self.Seats_Update(df, self.MainModel.data)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))
            self.SeatTable1.setModel(self.empty_Model)

    # 实验室管理刷新人员和工位视图
    def RoomShow(self):
        self.RoomStaffShow()
        self.RoomSeatShow()

    # 工牌打印
    def print_Word(self):
        try:
            QR = QRTool()
            res = QR.MakeCard(self.MainModel.data)
            if res:
                choice = QMessageBox.information(self, '文档成功生成', '已保存到Outputs文件夹\n点击Yes打开',
                                                 QMessageBox.Yes | QMessageBox.No)
                if choice == QMessageBox.Yes:
                    os.system("start Outputs/学生工位牌.docx")
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 工位图显示
    def Seats_Update(self, df1, df2):
        try:
            if self.StackView.currentIndex() == 0:
                self.SeatModel = SeatTableModel(df1)
                self.SeatModel.DataShow(df2)
                self.SeatTable.setModel(self.SeatModel)
            elif self.StackView.currentIndex() == 1:
                self.SeatModel = SeatTableModel(df1)
                self.SeatModel.DataShow(df2)
                self.SeatTable1.setModel(self.SeatModel)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 双击修改
    def Seats_modify(self):
        try:
            if self.StackView.currentIndex() == 0:
                QMessageBox.information(self, "无此权限", "学生无此权限", QMessageBox.Yes)
            if self.StackView.currentIndex() == 1:
                row = self.SeatTable1.currentIndex().row()
                col = self.SeatTable1.currentIndex().column()
                if str(self.SeatModel.DataDisplay.iloc[row, col]) == 'nan':
                    self.AddLine_Show()
                    self.add_view.setX(row)
                    self.add_view.setY(col)
                    self.add_view.setLab(self.LabSelect.currentText())
                    self.add_view.setRoom(self.RoomSelect.currentText())
                else:
                    self.Modify_show()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 修改视图展示
    def Modify_show(self):
        try:
            row = self.SeatTable1.currentIndex().row()
            col = self.SeatTable1.currentIndex().column()
            seat = int(self.SeatModel.DataDisplay.iloc[row, col])
            data = self.MainModel.data[self.MainModel.data["座位号"] == seat]
            self.modify_view = modify_view(data)
            self.modify_view.show()
            self.modify_view.signal_Modify.connect(self.AdvancedModify)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 高权限修改
    def AdvancedModify(self, data):
        try:
            Header = data["Header"]
            Input = data["Input"]
            Index = data["Index"]
            row = self.MainModel.data.loc[Index, "序号"]
            loc = self.MainModel.df[self.MainModel.df["序号"] == row].index.values
            self.MainModel.data.loc[Index, Header] = Input
            self.MainModel.pre_data = self.MainModel.data[["实验室号", "座位号", "姓名", "学号", "导师"]]
            self.MainModel.beginResetModel()
            self.MainModel.DataDisplay = self.MainModel.pre_data
            self.MainModel.endResetModel()
            self.MainModel.df.loc[loc, Header] = Input
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 一键更新
    def quick_update(self):
        try:
            self.UpdateView = UpdateTool()
            self.UpdateView.show()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    ##################################################################################################

    # 6、searchView
    # 正则标记
    def Reg(self):
        if self.reg_checkbox.checkState() == 2:
            self.reg_flag = True
        else:
            self.reg_flag = False

    # 查
    def Search(self):
        try:
            col = []
            reg = []
            for i in range(len(self.search_line)):
                if self.search_line[i].text():
                    col.append(self.search_label[i].text())
                    reg.append(self.search_line[i].text())
            if len(col):
                res = self.MainModel.search(col, reg, self.reg_flag)
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # 返回
    def Back(self):
        try:
            self.MainModel.dataUpdate()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    ##################################################################################################
    # 样式
    ##################################################################################################
    def Style(self):
        self.StyleChange()
        self.center(self)

    # 样式改变
    def StyleChange(self):
        try:
            config = configparser.ConfigParser()
            config.read("configs/config.ini")

            main_theme = config.get('CSS', "main_theme")
            main_background = ("True" == config["CSS"]["main_BackGround"])
            self.apply_stylesheet(self, theme=main_theme, invert_secondary=main_background)

            Menu_theme = config.get('CSS', "Menu_theme")
            Menu_background = ("True" == config["CSS"]["Menu_BackGround"])
            self.apply_stylesheet(self.menu, theme=Menu_theme, invert_secondary=Menu_background)

            Modify_theme = config.get('CSS', "Modify_theme")
            Modify_background = ("True" == config["CSS"]["Modify_BackGround"])
            self.apply_stylesheet(self.InputDialog, theme=Modify_theme, invert_secondary=Modify_background)

            icon_path = config.get('Path', 'icon')
            self.InputDialog.setWindowIcon(QIcon(icon_path))

            icon_path = config.get('Path', 'icon')
            self.setWindowIcon(QIcon(icon_path))

            self.MainTableStyle()
            self.SomeButtonStyle()
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # MainTableStyle
    def MainTableStyle(self):
        try:
            self.MainTable.setShowGrid(False)
            self.MainTable.setAlternatingRowColors(True)
            self.MainTable.verticalHeader().hide()
            self.MainTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
            self.MainTable.horizontalHeader().setFont(QFont("Roman times", 10, QFont.Bold, QFont.Black))
        except Exception as e:
            self.logger.debug("{} {}:{}".format(self.__class__.__name__, sys._getframe().f_code.co_name, str(e)))

    # SomeButtonStyle
    def SomeButtonStyle(self):
        self.exit_Button.setProperty('class', 'danger')

    # 居中
    def center(self, widget):
        screen = QDesktopWidget().screenGeometry()
        size = widget.geometry()
        widget.move((screen.width() - size.width()) / 2,
                    (screen.height() - size.height()) / 2)
    ##################################################################################################


if __name__ == '__main__':
    app = QApplication(sys.argv)
    demo = Demo()
    demo.show()
    sys.exit(app.exec_())
