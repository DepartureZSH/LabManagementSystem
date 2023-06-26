import re
import pandas as pd
import numpy as np
import os
import cv2
from cnocr import CnOcr
from PIL import Image
import configparser
from threading import Thread
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QColor, QFont, QPixmap
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QHBoxLayout, QVBoxLayout, QFormLayout, QComboBox
from PyQt5.QtWidgets import QLabel, QMessageBox, QPushButton, QLineEdit

from qt_material import QtStyleTools
from SeatModel import SeatModel
from Image_Recognition import TableRecognition
from PyQt5.QtWidgets import *
import sys

class UpdateTool(QWidget, QtStyleTools):
    def __init__(self):
        super(UpdateTool, self).__init__()
        self.setWindowTitle("一键更新")
        self.resize(1200, 400)
        self.Elements()
        self.Layout()
        self.Connection()
        self.Style()

    def Elements(self):
        self.LabLabel = QLabel("实验室门牌号")
        self.LabelLineEdit = QLineEdit()
        self.LabelLineEdit.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)
        self.LoadPicButton = QPushButton("识别图片")
        self.addRowAheadButton = QPushButton("添加一行(上方)")
        self.addRowBelowButton = QPushButton("添加一行(下方)")
        self.delRowButton = QPushButton("删除此行")
        self.addColLeftButton = QPushButton("添加一列(左侧)")
        self.addColRightButton = QPushButton("添加一列(右侧)")
        self.delColButton = QPushButton("删除此列")
        self.saveLabButton = QPushButton("保存")
        self.saveLabButton.setEnabled(False)
        self.addRowAheadButton.setEnabled(False)
        self.addRowBelowButton.setEnabled(False)
        self.delRowButton.setEnabled(False)
        self.addColLeftButton.setEnabled(False)
        self.addColRightButton.setEnabled(False)
        self.delColButton.setEnabled(False)
        self.Pic = QLabel()
        self.stackView = QStackedWidget()
        self.Pic.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.SeatTable = QTableView()
        self.stackView.addWidget(self.SeatTable)
        self.NoResultLabel = QLabel("无法识别该图片")
        self.NoResultLabel.setFont(QFont("Roman times", 10, QFont.Bold))
        self.NoResultLabel.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)
        self.stackView.addWidget(self.NoResultLabel)
        self.stackView.setCurrentIndex(0)

    def Layout(self):
        H_Layout1 = QHBoxLayout()
        H_Layout1.addWidget(self.LabLabel)
        H_Layout1.addWidget(self.LabelLineEdit)
        H_Layout1.addWidget(self.LoadPicButton)
        hSpacer = QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        H_Layout1.addItem(hSpacer)
        H_Layout2 = QHBoxLayout()
        H_Layout2.addItem(hSpacer)
        H_Layout2.addWidget(self.addRowAheadButton)
        H_Layout2.addWidget(self.addRowBelowButton)
        H_Layout2.addWidget(self.delRowButton)
        H_Layout2.addWidget(self.addColLeftButton)
        H_Layout2.addWidget(self.addColRightButton)
        H_Layout2.addWidget(self.delColButton)
        H_Layout2.addWidget(self.saveLabButton)
        GridLayout = QGridLayout()
        GridLayout.addLayout(H_Layout1, 0, 0, 1, 1)
        GridLayout.addItem(H_Layout2, 0, 1, 1, 1)
        GridLayout.addWidget(QLabel(""), 1, 0, 1, 2)
        Label1 = QLabel("被识别图片")
        Label1.setFont(QFont("Roman times", 10, QFont.Bold))
        Label1.setAlignment(Qt.AlignHCenter)
        Label2 = QLabel("生成表格")
        Label2.setFont(QFont("Roman times", 10, QFont.Bold))
        Label2.setAlignment(Qt.AlignHCenter)
        GridLayout.addWidget(Label1, 2, 0, 1, 1)
        GridLayout.addWidget(Label2, 2, 1, 1, 1)
        GridLayout.addWidget(self.Pic, 3, 0, 1, 1)
        GridLayout.addWidget(self.stackView, 3, 1, 1, 1)
        self.setLayout(GridLayout)

    def Connection(self):
        self.addRowBelowButton.clicked.connect(self.addRowBelow)
        self.addRowAheadButton.clicked.connect(self.addRowAhead)
        self.delRowButton.clicked.connect(self.delRow)
        self.addColRightButton.clicked.connect(self.addColRight)
        self.addColLeftButton.clicked.connect(self.addColLeft)
        self.delColButton.clicked.connect(self.delCol)
        self.LoadPicButton.clicked.connect(self.LoadPic)
        self.saveLabButton.clicked.connect(self.SaveLab)

    def addRowBelow(self):
        if self.SeatTable.currentIndex().row() == -1:
            self.TableModel.InsertRowBelow()
        else:
            self.TableModel.InsertRowBelow(self.SeatTable.currentIndex().row())

    def addRowAhead(self):
        if self.SeatTable.currentIndex().row() == -1:
            self.TableModel.InsertRowAhead()
        else:
            self.TableModel.InsertRowAhead(self.SeatTable.currentIndex().row())

    def delRow(self):
        if self.SeatTable.currentIndex().row() == -1:
            self.TableModel.delRow()
        else:
            self.TableModel.delRow(self.SeatTable.currentIndex().row())

    def addColRight(self):
        if self.SeatTable.currentIndex().column() == -1:
            self.TableModel.InsertColRight()
        else:
            self.TableModel.InsertColRight(self.SeatTable.currentIndex().column())

    def addColLeft(self):
        if self.SeatTable.currentIndex().column() == -1:
            self.TableModel.InsertColLeft()
        else:
            self.TableModel.InsertColLeft(self.SeatTable.currentIndex().column())

    def delCol(self):
        if self.SeatTable.currentIndex().column() == -1:
            self.TableModel.delCol()
        else:
            self.TableModel.delCol(self.SeatTable.currentIndex().column())

    def LoadPic(self):
        self.stackView.setCurrentIndex(0)
        try:
            self.imageName, imgType = QFileDialog.getOpenFileName(self, "识别图片", "imgs", "*.png;;*.jpg;;All Files(*)")
        except Exception as e:
            print(str(e))
            return
        t = Thread(target=self.RecognitionThread)
        t.start()
        try:
            jpg = QPixmap(self.imageName)
            self.Pic.setPixmap(jpg)
            self.Pic.setScaledContents(True)
            self.Pic.setMaximumSize(600, 400)
        except Exception as e:
            print(str(e))
            return

    def RecognitionThread(self):
        try:
            TR = TableRecognition()
            res = TR.Recognition(self.imageName)
            res = pd.DataFrame(res)
            self.TableModel = SeatModel(res)
            self.SeatTable.setModel(self.TableModel)
            self.saveLabButton.setEnabled(True)
            self.addRowAheadButton.setEnabled(True)
            self.addRowBelowButton.setEnabled(True)
            self.delRowButton.setEnabled(True)
            self.addColLeftButton.setEnabled(True)
            self.addColRightButton.setEnabled(True)
            self.delColButton.setEnabled(True)
            self.stackView.setCurrentIndex(0)
        except Exception as e:
            print(str(e))
            self.stackView.setCurrentIndex(1)

    def SaveLab(self):
        if self.LabelLineEdit.text() == '':
            QMessageBox.information(self, "提示", "请输入该实验室门牌号", QMessageBox.Yes)
            return
        try:
            Data = self.TableModel.DataDisplay.copy()
            for i in range(Data.shape[0]):
                for j in range(Data.shape[1]):
                    if str(Data.iloc[i, j]) == '':
                        pass
                    else:
                        Data.iloc[i, j] = 'T'
            Data.to_excel(self.path + self.LabelLineEdit.text() + ".xlsx", sheet_name='Sheet1', index=False, header=None)
            choice = QMessageBox.information(self, "添加成功", "该工位图掩码已保存到data文件夹\n点击Yes打开",
                                             QMessageBox.Yes | QMessageBox.No)
            if choice == QMessageBox.Yes:
                os.system("start {}".format(self.path + self.LabelLineEdit.text() + ".xlsx"))
        except Exception as e:
            print(str(e))

    def Style(self):
        config = configparser.ConfigParser()
        config.read("configs/config.ini")
        UpdateView_theme = config["CSS"]["Main_Theme"]
        UpdateView_background = ("True" == config["CSS"]["Modify_BackGround"])
        self.apply_stylesheet(self, theme=UpdateView_theme, invert_secondary=UpdateView_background)
        icon_path = config.get('Path', 'icon')
        self.setWindowIcon(QIcon(icon_path))
        self.path = config.get('Path', 'data')
        self.SeatTable.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.SeatTable.verticalHeader().setSectionResizeMode(QHeaderView.Stretch)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    demo = UpdateTool()
    demo.show()
    sys.exit(app.exec_())
