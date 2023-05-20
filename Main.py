import configparser
import re
import sys
import pandas as pd
import numpy as np
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


class Demo(QWidget, QtStyleTools):
    # 初始化
    signal_addLine = pyqtSignal(dict)
    signal_Modify = pyqtSignal(dict)
    signal_UpdateSetting = pyqtSignal()

    def __init__(self):
        super(Demo, self).__init__()
        self.resize(1400, 900)
        self.setWindowTitle("实验室人员管理")
        self.elements()
        self.pop_up_elements()
        self.Layout()
        self.connection()
        self.StackView_connection()
        self.conf_affirm()
        self.Style()
        self.Parameters()

    # 配置文件检查
    def conf_affirm(self):
        setting.ConfigCheck()

    ##############################################################################
    # 参数
    def Parameters(self):
        config = configparser.ConfigParser()
        config.read("configs/config.ini")
        self.OutputPath = config.get("Path", "Outputs")
        self.DataPath = config.get("Path", "Data")

    # 组件库
    def elements(self):
        self.StackView = QStackedWidget(self)
        self.menu = QWidget()
        self.LM_Button = QPushButton("实验室管理")
        self.LMS_Button = QPushButton("实验室人员管理")
        self.setting_Button = QPushButton("设置")
        self.exit_Button = QPushButton("退出")

        # 表格视图
        self.MainTable = QTableView()
        self.MainModel = pandasModel()
        self.MainTable.setModel(self.MainModel)
        self.StackView_Widgets()

        # 详细信息视图
        self.InfoWidget = QWidget()
        self.InfoWidgetUI()
        self.StackView.addWidget(self.InfoWidget)

        # 实验室管理视图
        self.LM_Widget = QWidget()
        self.LM_WidgetUI()
        self.StackView.addWidget(self.LM_Widget)

        # 保存
        self.save_button = QPushButton('保存', self)
        # 另存为
        self.save_as_button = QPushButton('另存为', self)
        # 增
        self.add_line_button = QPushButton('新增', self)
        # 删
        self.remove_line_button = QPushButton('删除', self)
        # 查
        self.search_label = []
        self.search_line = []
        for each in self.MainModel.DataDisplay.columns.values:
            self.search_label.append(QLabel(each))
            self.search_line.append(QLineEdit(self))
        self.reg_flag = False
        self.reg_checkbox = QCheckBox('正则匹配', self)
        self.search_button = QPushButton('搜索', self)
        self.back_button = QPushButton('返回', self)
        self.StackView.setCurrentIndex(0)

    def pop_up_elements(self):
        empty_df = pd.DataFrame(["请添加工位图"], columns=["查询无结果"])
        self.empty_Model = SeatTableModel(empty_df)

        self.InputDialog = QInputDialog()

        self.InputDialog.setWindowTitle("添加实验室门牌号")
        self.InputDialog.setLabelText("请输入实验室门牌号:")
        self.InputDialog.setFixedSize(350, 400)
        self.InputDialog.setOkButtonText(u"确定")
        self.InputDialog.setCancelButtonText(u"取消")
        self.InputDialog.setWindowFlags(Qt.WindowStaysOnTopHint)

    # 总布局
    def Layout(self):
        self.h_layout = QHBoxLayout()
        self.v_layout1 = QVBoxLayout()
        self.v_layout23s = QVBoxLayout()
        self.h_layout23 = QHBoxLayout()
        self.v_layout2 = QVBoxLayout()
        self.v_layout3 = QVBoxLayout()
        self.search_layout = QHBoxLayout()

        self.v_layout1.addWidget(self.LM_Button)
        self.v_layout1.addWidget(self.LMS_Button)
        vSpacer = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.v_layout1.addItem(vSpacer)
        self.v_layout1.addWidget(self.setting_Button)
        self.v_layout1.addWidget(self.exit_Button)
        self.menu.setLayout(self.v_layout1)

        self.v_layout2.addWidget(self.StackView)

        self.h_layout3_1 = QHBoxLayout()
        self.h_layout3_1.addWidget(self.save_button)
        self.h_layout3_1.addWidget(self.save_as_button)
        Label1 = QLabel("实验室人员名单")
        Label1.setFont(QFont("Roman times", 10, QFont.Bold))
        Label1.setAlignment(Qt.AlignHCenter)
        self.v_layout3.addWidget(Label1)
        self.v_layout3.addLayout(self.h_layout3_1)
        self.v_layout3.addWidget(self.MainTable)
        self.h_layout3_2 = QHBoxLayout()
        self.h_layout3_2.addWidget(self.add_line_button)
        self.h_layout3_2.addWidget(self.remove_line_button)
        self.v_layout3.addLayout(self.h_layout3_2)

        for i in range(len(self.search_line)):
            self.search_layout.addWidget(self.search_label[i])
            self.search_layout.addWidget(self.search_line[i])
        self.search_layout.addWidget(self.reg_checkbox)
        self.search_layout.addWidget(self.search_button)
        self.search_layout.addWidget(self.back_button)

        self.h_layout23.addLayout(self.v_layout2)
        self.h_layout23.addLayout(self.v_layout3)
        self.v_layout23s.addLayout(self.h_layout23)
        self.v_layout23s.addLayout(self.search_layout)

        self.h_layout.addWidget(self.menu)
        self.h_layout.addLayout(self.v_layout23s)

        self.setLayout(self.h_layout)

    # 信号槽连接
    def connection(self):
        ########################################################
        # 添加记录
        self.add_line_button.clicked.connect(self.AddLine_Show)
        # 删除一行
        self.remove_line_button.clicked.connect(self.Remove_line)
        # 正则选择
        self.reg_checkbox.clicked.connect(self.Reg)
        # 搜索
        self.search_button.clicked.connect(self.Search)
        # 返回
        self.back_button.clicked.connect(self.Back)
        ########################################################
        # 保存
        self.save_button.clicked.connect(self.Save)
        # 另存为
        self.save_as_button.clicked.connect(self.Save_as)
        # 设置(undone)
        self.setting_Button.clicked.connect(self.Setting)
        # 退出
        self.exit_Button.clicked.connect(self.Exit)
        ########################################################
        # 实验室学生管理
        self.LMS_Button.clicked.connect(self.LMS_Show)
        # 实验室管理
        self.LM_Button.clicked.connect(self.LM_Show)
        ########################################################
        # 展示详细信息
        self.MainTable.clicked.connect(self.ShowInfo)
        # 通过工位图修改
        self.SeatTable.doubleClicked.connect(self.Seats_modify)
        self.SeatTable1.doubleClicked.connect(self.Seats_modify)
        ########################################################

    ##############################################################################
    # StackView组件
    def StackView_Widgets(self):
        header = self.MainModel.data.columns.values
        # InfoWidget
        self.LineEdit = {}
        self.modify_button = {}
        for i in range(len(header)):
            if i == 0:
                self.LineEdit[header[i]] = QLabel("")
            else:
                self.LineEdit[header[i]] = QLineEdit("")
                self.modify_button[header[i]] = QPushButton('修改', self)
        # 工位视图
        self.SeatTable = QTableView()
        self.SeatTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # 打印
        self.print_Word_button = QPushButton('生成工牌打印文件', self)

        # LM_Widget
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

    # StackView信号槽连接
    def StackView_connection(self):
        header = self.MainModel.data.columns.values
        self.modify_button[header[1]].clicked.connect(lambda: self.Modify(header[1]))
        self.modify_button[header[2]].clicked.connect(lambda: self.Modify(header[2]))
        self.modify_button[header[3]].clicked.connect(lambda: self.Modify(header[3]))
        self.modify_button[header[4]].clicked.connect(lambda: self.Modify(header[4]))
        self.modify_button[header[5]].clicked.connect(lambda: self.Modify(header[5]))
        self.modify_button[header[6]].clicked.connect(lambda: self.Modify(header[6]))
        self.modify_button[header[7]].clicked.connect(lambda: self.Modify(header[7]))
        self.modify_button[header[8]].clicked.connect(lambda: self.Modify(header[8]))
        # 实验室选择框
        self.LabSelect.currentTextChanged.connect(self.RoomComboBox)
        # 实验室门牌号选择框
        self.RoomSelect.currentTextChanged.connect(self.RoomShow)
        # 打印
        self.print_Word_button.clicked.connect(self.print_Word)
        self.print_Word_button1.clicked.connect(self.print_Word)
        # 一键更新
        self.quick_update_button.clicked.connect(self.quick_update)
        # 添加实验室
        self.AddLabButton.clicked.connect(self.AddLab)
        # 删除实验室
        self.DelLabButton.clicked.connect(self.DelLab)
        # 添加实验室门牌
        self.AddRoomButton.clicked.connect(self.AddRoom)
        # 删除实验室门牌
        self.DelRoomButton.clicked.connect(self.DelRoom)
        # 保存实验室信息
        self.LMSaveButton.clicked.connect(self.LMSave)

    ##############################################################################
    # 详细信息视图布局
    def InfoWidgetUI(self):
        # 详细信息
        GridLayout = QGridLayout()
        header = self.MainModel.data.columns.values
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
        # 布局
        VLayout = QVBoxLayout()
        Label1 = QLabel("详细信息")
        Label1.setFont(QFont("Roman times", 10, QFont.Bold))
        Label1.setAlignment(Qt.AlignHCenter)
        VLayout.addWidget(Label1)
        VLayout.addLayout(GridLayout)
        Label2 = QLabel("工位图")
        Label2.setFont(QFont("Roman times", 10, QFont.Bold))
        Label2.setAlignment(Qt.AlignHCenter)
        VLayout.addWidget(Label2)
        VLayout.addWidget(self.SeatTable)
        VLayout.addWidget(self.print_Word_button)
        self.InfoWidget.setLayout(VLayout)

    # 详细信息展示
    def ShowInfo(self):
        # InfoWidget = QWidget()
        self.index = self.MainTable.currentIndex().row()
        Info = {}
        try:
            for i in range(len(self.MainModel.data.columns.values)):
                Info[self.MainModel.data.columns.values[i]] = self.MainModel.data.iloc[self.index, i]
        except Exception as e:
            print("ShowInfo1")
            print(str(e))
        try:
            if str(Info['实验室门牌号']) == 'nan':
                print("nan")
                raise Exception("实验室门牌号 == nan")
            else:
                df = pd.read_excel("data/{}.xlsx".format(int(Info['实验室门牌号'])), header=None)
                df2 = self.MainModel.data[self.MainModel.data["实验室门牌号"] == int(Info['实验室门牌号'])]
                self.Seats_Update(df, df2)
        except Exception as e:
            print("ShowInfo2")
            self.SeatTable.setModel(self.empty_Model)
            print(str(e))
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
            print("ShowInfo3")
            print(str(e))

    ##############################################################################
    # 添加实验室
    def AddLab(self):
        # print(self.LabSelect.count())
        num = str(self.LabSelect.count() + 1)
        self.LabSelect.addItem(num)
        QMessageBox.information(self, "提示", "实验室{}已添加\n请添加实验室门牌并保存使此操作生效".format(num), QMessageBox.Yes)

    # 删除实验室
    def DelLab(self):
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

    # 添加实验室门牌号
    def AddRoom(self):
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

    # 删除实验室门牌号
    def DelRoom(self):
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

    # 实验室保存
    def LMSave(self):
        try:
            res = self.MainModel.PrimarySave()
        except Exception as e:
            print(str(e))
        try:
            res = self.MainModel.AdvancedSave()
        except Exception as e:
            print("MainModel")
            print(str(e))
        try:
            path = self.DataPath + self.RoomSelect.currentText()+".xlsx"
            self.SeatModel.df.to_excel(path, sheet_name='Sheet1', index=False, header=None)
        except Exception as e:
            print("MainModel")
            print(str(e))

    ##############################################################################
    # 实验室管理视图布局
    def LM_WidgetUI(self):
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
        self.LM_Widget.setLayout(VLayout)

    # 实验室选择框
    def LabComboBox(self):
        try:
            self.LabSelect.currentTextChanged.disconnect(self.RoomComboBox)
        except:
            pass
        self.LabSelect.clear()
        self.LabSelect.currentTextChanged.connect(self.RoomComboBox)
        LabList = np.unique(self.MainModel.df3["实验室号"].values).tolist()
        self.LabSelect.addItems([str(each) for each in LabList])

    # 实验室门牌号选择框
    def RoomComboBox(self):
        try:
            self.RoomSelect.currentTextChanged.disconnect(self.RoomShow)
        except:
            pass
        self.RoomSelect.clear()
        self.RoomSelect.currentTextChanged.connect(self.RoomShow)
        df = self.MainModel.df3.loc[self.MainModel.df3["实验室号"] == int(self.LabSelect.currentText())]
        RoomList = df["实验室门牌号"].tolist()
        self.RoomSelect.addItems([str(each) for each in RoomList])

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
            print("RoomStaffShow")
            print(str(e))

    # 实验室管理之工位视图
    def RoomSeatShow(self):
        if self.RoomSelect.currentText() == "":
            return
        try:
            print(self.RoomSelect.currentText())
            df = pd.read_excel("data/{}.xlsx".format(self.RoomSelect.currentText()), header=None)
            self.Seats_Update(df, self.MainModel.data)
        except Exception as e:
            print("RoomSeatShow")
            self.SeatTable1.setModel(self.empty_Model)
            print(str(e))

    # 实验室管理刷新人员和工位视图
    def RoomShow(self):
        self.RoomStaffShow()
        self.RoomSeatShow()

    ##############################################################################
    # 添加一行组件视图
    def AddLine_Show(self):
        if self.StackView.currentIndex() == 0:
            QMessageBox.information(self, "提示", "该功能只能由管理人员操作", QMessageBox.Yes)
        elif self.StackView.currentIndex() == 1:
            try:
                self.add_view = add_view(self.MainModel.df, self.MainModel.df3)
                self.add_view.show()
                self.add_view.signal_addLine.connect(self.Insert)
            except Exception as e:
                print("AddLine_Show")
                print(str(e))

    # 增
    def Insert(self, data):
        self.MainModel.temp = pd.DataFrame([data])
        self.MainModel.insertRows(self.MainModel.data.shape[0])
        if self.StackView.currentIndex() == 0:
            self.ShowInfo()
        elif self.StackView.currentIndex() == 1:
            self.RoomSeatShow()

    ##############################################################################
    # 删
    def Remove_line(self):
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
                print("a")
                self.Remove(index)

    def Remove(self, index):
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
                    print(str(e))

    ##############################################################################
    # 普通修改
    def Modify(self, header):
        try:
            if self.index:
                pass
        except Exception as e:
            print(str(e))
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
            print(str(e))

    ##############################################################################
    # 正则标记
    def Reg(self):
        if self.reg_checkbox.checkState() == 2:
            self.reg_flag = True
        else:
            self.reg_flag = False

    # 查
    def Search(self):
        col = []
        reg = []
        for i in range(len(self.search_line)):
            if self.search_line[i].text():
                col.append(self.search_label[i].text())
                reg.append(self.search_line[i].text())
        if len(col):
            res = self.MainModel.search(col, reg, self.reg_flag)

    # 返回
    def Back(self):
        self.MainModel.dataUpdate()

    ##############################################################################
    # 工牌打印
    def print_Word(self):
        QR = QRTool()
        res = QR.MakeCard(self.MainModel.data)
        if res:
            choice = QMessageBox.information(self, '二维码文档已生成', '是否需要一键打印', QMessageBox.Yes | QMessageBox.No)
            if choice == QMessageBox.Yes:
                QMessageBox.information(self, '操作成功', '请等待连接打印机', QMessageBox.Yes)

    ##############################################################################
    # 工位图显示
    def Seats_Update(self, df1, df2):
        if self.StackView.currentIndex() == 0:
            self.SeatModel = SeatTableModel(df1)
            self.SeatModel.DataShow(df2)
            self.SeatTable.setModel(self.SeatModel)
        elif self.StackView.currentIndex() == 1:
            self.SeatModel = SeatTableModel(df1)
            self.SeatModel.DataShow(df2)
            self.SeatTable1.setModel(self.SeatModel)

    # 双击修改
    def Seats_modify(self):
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
            print(str(e))

    # 高权限修改
    def AdvancedModify(self, data):
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

    ##############################################################################
    def quick_update(self):
        pass

    ##############################################################################
    # 保存
    def Save(self):
        try:
            self.MainModel.PrimarySave()
        except Exception as e:
            print(str(e))

    # 另存为
    def Save_as(self):
        path, _ = QFileDialog.getSaveFileName(self, 'Save File', self.OutputPath, 'Files (*.xlsx)')
        self.MainModel.CustomizedSave(path)

    # 设置
    def Setting(self):
        self.settingView = setting()
        self.settingView.show()
        self.settingView.signal_UpdateSetting.connect(self.StyleChange)

    # 退出
    def Exit(self):
        res = QMessageBox.question(self, "确认退出?", "请先保存后退出。", QMessageBox.Yes | QMessageBox.No)
        if res == QMessageBox.Yes:
            self.close()

    ##############################################################################
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
            print(str(e))

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
            print(str(e))

    ##############################################################################

    # MainTableStyle
    def MainTableStyle(self):
        self.MainTable.setShowGrid(False)
        self.MainTable.setAlternatingRowColors(True)
        self.MainTable.verticalHeader().hide()
        self.MainTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.MainTable.horizontalHeader().setFont(QFont("Roman times", 10, QFont.Bold, QFont.Black))

    # 样式改变
    def StyleChange(self):
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

    # SomeButtonStyle
    def SomeButtonStyle(self):
        self.exit_Button.setProperty('class', 'danger')

    # style
    def Style(self):
        self.StyleChange()
        self.center(self)

    # 居中
    def center(self, widget):
        screen = QDesktopWidget().screenGeometry()
        size = widget.geometry()
        widget.move((screen.width() - size.width()) / 2,
                    (screen.height() - size.height()) / 2)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    demo = Demo()
    demo.show()
    sys.exit(app.exec_())

