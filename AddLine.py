import re
import pandas as pd
import numpy as np
import configparser
from qt_material import QtStyleTools
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QHBoxLayout, QVBoxLayout, QFormLayout, QComboBox
from PyQt5.QtWidgets import QLabel, QMessageBox, QPushButton, QLineEdit

class add_view(QWidget, QtStyleTools):
    signal_addLine = pyqtSignal(dict)

    def __init__(self, df, df3):
        super(add_view, self).__init__()
        self.df = df
        self.df3 = df3
        self.data_reform()
        self.widgets()
        self.Layouts()
        self.Style()

    def setLab(self, Lab):
        try:
            self.LabSelect.setCurrentText(str(Lab))
            self.LabSelect.setEnabled(False)
        except Exception as e:
            print(str(e))

    def setRoom(self, Room):
        try:
            self.RoomSelect.setCurrentText(str(Room))
            self.RoomSelect.setEnabled(False)
        except Exception as e:
            print(str(e))

    def setX(self, X):
        try:
            self.add_list["工位位置X"].setText(str(X))
            self.add_list["工位位置X"].setEnabled(False)
        except Exception as e:
            print(str(e))

    def setY(self, Y):
        try:
            self.add_list["工位位置Y"].setText(str(Y))
            self.add_list["工位位置Y"].setEnabled(False)
        except Exception as e:
            print(str(e))

    def data_reform(self):
        self.headers = self.df.columns.values[1:]
        self.LabList = np.unique(self.df3["实验室号"].values).tolist()

    # 实验室选择框
    def LabComboBox(self):
        self.LabSelect.addItems([str(each) for each in self.LabList])

    # 实验室门牌号选择框
    def RoomComboBox(self):
        self.RoomSelect.clear()
        df = self.df3.loc[self.df3["实验室号"] == int(self.LabSelect.currentText())]
        RoomList = df["实验室门牌号"].tolist()
        self.RoomSelect.addItems([str(each) for each in RoomList])

    # 组件
    def widgets(self):
        self.add_label_list = {}
        self.add_list = {}
        self.LabSelect = QComboBox()
        self.LabComboBox()
        self.RoomSelect = QComboBox()
        self.RoomComboBox()
        self.LabSelect.currentTextChanged.connect(self.RoomComboBox)
        for each in self.headers:
            self.add_label_list[each] = QLabel(each)
            if each == "实验室号" or each == "实验室门牌号":
                continue
            self.add_list[each] = QLineEdit()
        self.submit_bottom = QPushButton('确定')
        self.submit_bottom.clicked.connect(self.submit)
        self.cancer_bottom = QPushButton('取消')
        self.cancer_bottom.clicked.connect(self.cancer)

    # 添加一行表单视图
    def Layouts(self):
        layout4 = QVBoxLayout()
        layout4_1 = QFormLayout()
        layout4_2 = QHBoxLayout()
        layout4.setAlignment(Qt.AlignCenter | Qt.AlignVCenter)
        for each in self.headers:
            if each == "实验室号":
                layout4_1.addRow(self.add_label_list[each], self.LabSelect)
            elif each == "实验室门牌号":
                layout4_1.addRow(self.add_label_list[each], self.RoomSelect)
            else:
                layout4_1.addRow(self.add_label_list[each], self.add_list[each])
        layout4_2.addWidget(self.submit_bottom)
        layout4_2.addWidget(self.cancer_bottom)
        layout4.addLayout(layout4_1)
        layout4.addLayout(layout4_2)
        self.setLayout(layout4)

    # 新建提交
    def submit(self):
        add_list = {}
        try:
            add_list['序号'] = self.df.shape[0] + 1
            for each in self.df.columns.values[1:]:
                if each == "实验室号":
                    add_list[each] = self.LabSelect.currentText()
                elif each == "实验室门牌号":
                    add_list[each] = self.RoomSelect.currentText()
                elif self.add_list[each].text():
                    add_list[each] = self.add_list[each].text()
                else:
                    add_list[each] = ""
                    raise Exception("请输入{}".format(self.add_label_list[each].text()))
            if add_list['姓名']:
                if re.match(r"^(?:[\u4e00-\u9fa5]+)(?:●[\u4e00-\u9fa5]+)*$|^[a-zA-Z0-9]+\s?[\.·\-()a-zA-Z]*[a-zA-Z]+$", add_list['姓名']):
                    pass
                else:
                    raise Exception("请输入正确的姓名（中文名或英文名）")
            if add_list['学号']:
                if re.match(r"^(S[0-9]{9})*$|^(B[0-9]{9})*$", add_list['学号']):
                    pass
                else:
                    raise Exception("请输入正确的学号（S或B+9位数字）")
            if add_list['导师']:
                if re.match(r"^(?:[\u4e00-\u9fa5]+)(?:●[\u4e00-\u9fa5]+)*$|^[a-zA-Z0-9]+\s?[\.·\-()a-zA-Z]*[a-zA-Z]+$", add_list['导师']):
                    pass
                else:
                    raise Exception("请输入正确的导师姓名（中文名或英文名）")
            if add_list['实验室门牌号']:
                add_list['实验室门牌号'] = int(add_list['实验室门牌号'])
            if add_list['座位号']:
                if re.match(r"^[0-9]{1,2}$", add_list['座位号']):
                    add_list['座位号'] = int(add_list['座位号'])
                    # 被占用异常
                else:
                    raise Exception("请输入正确的座位号（正整数）")
            if add_list['工位位置X'] and add_list['工位位置Y']:
                if re.match(r"^[0-9]{1,2}$", add_list['工位位置X']):
                    if re.match(r"^[0-9]{1,2}$", add_list['工位位置Y']):
                        add_list['工位位置X'] = float(add_list['工位位置X'])
                        add_list['工位位置Y'] = float(add_list['工位位置Y'])
                        # 被占用异常
                    else:
                        raise Exception("请输入正确的工位位置（整数）")
            if add_list['实验室号']:
                add_list['实验室号'] = int(add_list['实验室号'])
            # 发出信号
            self.signal_addLine.emit(add_list)
        except Exception as e:
            QMessageBox.information(self, '提示', '{}'.format(str(e)), QMessageBox.Yes)
            print(str(e))

    # 新建取消
    def cancer(self):
        self.close()

    # 样式改变
    def Style(self):
        config = configparser.ConfigParser()
        config.read("configs/config.ini")
        self.setWindowTitle("新增")
        icon_path = config.get('Path', 'icon')
        self.setWindowIcon(QIcon(icon_path))
        AddLine_theme = config.get('CSS', "AddLine_theme")
        AddLine_background = ("True" == config["CSS"]["AddLine_BackGround"])
        self.apply_stylesheet(self, theme=AddLine_theme, invert_secondary=AddLine_background)

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     demo = add_view(["实验室号", "实验室门牌号", "座位号", "工位位置X", "工位位置Y", "姓名", "学号", "导师"])
#     demo.show()
#     sys.exit(app.exec_())
