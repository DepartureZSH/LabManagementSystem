import re
import pandas as pd
import numpy as np
import configparser
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QHBoxLayout, QGridLayout, QVBoxLayout, QFormLayout, QComboBox
from PyQt5.QtWidgets import QLabel, QMessageBox, QPushButton, QLineEdit
from qt_material import QtStyleTools

class modify_view(QWidget, QtStyleTools):
    signal_Modify = pyqtSignal(dict)

    def __init__(self, df):
        super(modify_view, self).__init__()
        self.df = df
        self.widgets()
        self.Layouts()
        self.Style()

    # 组件
    def widgets(self):
        header = self.df.columns.values
        # InfoWidget
        self.LineEdit = {}
        self.modify_button = {}
        for i in range(len(header)):
            if i == 0:
                self.LineEdit[header[i]] = QLabel(str(self.df.iloc[0, i]))
            else:
                self.LineEdit[header[i]] = QLineEdit("")
                self.modify_button[header[i]] = QPushButton('修改', self)
                if str(self.df.iloc[0, i]) == 'nan':
                    continue
                elif type(self.df.iloc[0, i]) == np.float64 or type(self.df.iloc[0, i]) == float:
                    self.LineEdit[header[i]].setPlaceholderText(str(int(self.df.iloc[0, i])))
                else:
                    self.LineEdit[header[i]].setPlaceholderText(str(self.df.iloc[0, i]))
        self.modify_button[header[1]].clicked.connect(lambda: self.Modify(header[1]))
        self.modify_button[header[2]].clicked.connect(lambda: self.Modify(header[2]))
        self.modify_button[header[3]].clicked.connect(lambda: self.Modify(header[3]))
        self.modify_button[header[4]].clicked.connect(lambda: self.Modify(header[4]))
        self.modify_button[header[5]].clicked.connect(lambda: self.Modify(header[5]))
        self.modify_button[header[6]].clicked.connect(lambda: self.Modify(header[6]))
        self.modify_button[header[7]].clicked.connect(lambda: self.Modify(header[7]))
        self.modify_button[header[8]].clicked.connect(lambda: self.Modify(header[8]))
        self.cancer_bottom = QPushButton('关闭')
        self.cancer_bottom.clicked.connect(self.cancer)

    # 添加一行表单视图
    def Layouts(self):
        # 详细信息
        GridLayout = QGridLayout()
        header = self.df.columns.values
        for i in range(len(header)):
            each = header[i]
            if each == "序号":
                GridLayout.addWidget(QLabel(each), i + 1, 0)
                GridLayout.addWidget(self.LineEdit[each], i + 1, 1)
            else:
                GridLayout.addWidget(QLabel(each), i + 1, 0)
                GridLayout.addWidget(self.LineEdit[each], i + 1, 1)
                GridLayout.addWidget(self.modify_button[each], i + 1, 2)
        # 布局
        VLayout = QVBoxLayout()
        Label1 = QLabel("详细信息")
        Label1.setFont(QFont("Roman times", 10, QFont.Bold))
        Label1.setAlignment(Qt.AlignHCenter)
        VLayout.addWidget(Label1)
        VLayout.addLayout(GridLayout)
        VLayout.addWidget(self.cancer_bottom)
        self.setLayout(VLayout)

    # 改
    def Modify(self, header):
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
            res = {
                "Header": header,
                "Input": Input,
                "Index": self.df.index.values[0]
            }
            self.signal_Modify.emit(res)
            self.LineEdit[header].setText("")
            self.LineEdit[header].setPlaceholderText(str(Input))
            return

        except Exception as e:
            QMessageBox.information(self, '提示', '{}'.format(str(e)), QMessageBox.Yes)
            print(str(e))

    # 取消
    def cancer(self):
        self.close()

    # 样式
    def Style(self):
        config = configparser.ConfigParser()
        config.read("configs/config.ini")
        self.setWindowTitle("修改")
        icon_path = config.get('Path', 'icon')
        self.setWindowIcon(QIcon(icon_path))
        Modify_theme = config.get('CSS', "Modify_theme")
        Modify_background = ("True" == config["CSS"]["Modify_BackGround"])
        self.apply_stylesheet(self, theme=Modify_theme, invert_secondary=Modify_background)
