import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QApplication, QTableView, QWidget, QHeaderView
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt, pyqtSignal
from PyQt5.QtGui import QColor

# 数据模型
class SeatTableModel(QAbstractTableModel):
    """
    QAbstractTableModel实例化需要：

    - rowCount() 返回行数
    - columnCount() 返回列数
    - data() 返回表格数据
    - headerData() 返回表格标题

    """

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self.df = data.copy()
        self.pre_df = data.copy()
        self.DataDisplay = data.copy()
        self.currentX = -1
        self.currentY = -1

    def setcurrentXY(self, x, y):
        self.currentX = x
        self.currentY = y

    def DataShow(self, data):
        try:
            data = data[["座位号", "工位位置X", "工位位置Y"]]
            for i in range(data.shape[0]):
                try:
                    # 座位号
                    num = data.iloc[i, 0]
                    if str(num) == 'nan':
                        continue
                    else:
                        num = int(num)
                    # 工位位置X
                    X = data.iloc[i, 1]
                    if str(X) == 'nan':
                        continue
                    else:
                        X = int(X)
                    # 工位位置Y
                    Y = data.iloc[i, 2]
                    if str(Y) == 'nan':
                        continue
                    else:
                        Y = int(Y)
                    if self.pre_df.iloc[X, Y] == 'T':
                        self.pre_df.iloc[X, Y] = num
                    else:
                        self.df.iloc[X, Y] = 'T'
                        self.pre_df.iloc[X, Y] = num
                except Exception as e:
                    print(str(e))
            self.beginResetModel()
            self.DataDisplay = self.pre_df
            self.endResetModel()
        except Exception as e:
            print(str(e))

    def rowCount(self, parent=None):
        """
        表格行数

        :param parent: 主要是给树形列表用的，忽略
        :return: 返回值：表格行数
        """
        return self.DataDisplay.shape[0]

    def columnCount(self, parent=None):
        """
        表格列数

        :param parent: 主要是给树形列表用的，忽略
        :return: 返回值：表格列数
        """
        return self.DataDisplay.shape[1]

    def data(self, index, role=Qt.DisplayRole):
        """
        表格数据

        :param index: 模型索引，用于通过模型检索或修改数据
        :param role: 表格的状态：Qt.DisplayRole——视图的单元格一般状态下（例如初始化显示时，编辑完成时），要显示的数据。
        :return: 各datas数据
        """
        if index.isValid():
            if role == Qt.TextAlignmentRole:
                return Qt.AlignCenter
            if role == Qt.DisplayRole or role == Qt.EditRole:
                if str(self.DataDisplay.iloc[index.row(), index.column()]) == "nan":
                    return ""
                else:
                    try:
                        return str(int(self.DataDisplay.iloc[index.row(), index.column()]))
                    except:
                        return str(self.DataDisplay.iloc[index.row(), index.column()])
        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        """
        表格标题

        :param section: 表格列数
        :param orientation: 标题朝向
        :param role: 表格的状态：Qt.DisplayRole——视图的单元格一般状态下（例如初始化显示时，编辑完成时），要显示的数据。
        :return: 各datas横向标题
        """
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return int(section)
        elif orientation == Qt.Vertical and role == Qt.DisplayRole:
            return int(section)
        #     return self.DataDisplay.columns[col]
        return None

    # 改
    def setData(self, index, value, role=Qt.EditRole):
        """
        修改并更新模型中的数据

        :param index: 修改的位置
        :param value: 修改的值
        :param role: 表格的状态：Qt.DisplayRole——当单元格（item）进入编辑态时（一般双击会进入编辑态），要显示的数据，
        :return: 修改成功返回True，修改失败返回False
        """
        # 编辑后更新模型中的数据 View中编辑后，View会调用这个方法修改Model中的数据
        # print(self.data.shape[0])
        if index.isValid() and 0 <= index.row() < self.DataDisplay.shape[0] and value:
            col = index.column()
        if 1 <= col < self.DataDisplay.shape[1]:
            self.beginResetModel()
            self.endResetModel()
            return True
        return False

    def flags(self, index):
        """
        设置索引对应表格元素的标志

        - Qt.ItemIsEnabled：用户可以与项目交互
        - Qt.ItemIsEditable：可以被编辑
        - Qt.ItemIsSelectable：可被选择

        :param index: 模型索引
        :return: 标志位
        """
        # 必须实现的接口方法，不实现，则View中数据不可编辑
        if not index.isValid():
            return Qt.ItemIsEnabled
        return Qt.ItemFlags(QAbstractTableModel.flags(self, index) | Qt.ItemIsSelectable)
