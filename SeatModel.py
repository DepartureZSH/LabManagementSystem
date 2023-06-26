import sys
import pandas as pd
import numpy as np
from PyQt5.QtWidgets import QApplication, QTableView, QWidget, QHeaderView
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt, pyqtSignal
from PyQt5.QtGui import QColor


# 数据模型
class SeatModel(QAbstractTableModel):
    """
    QAbstractTableModel实例化需要：

    - rowCount() 返回行数
    - columnCount() 返回列数
    - data() 返回表格数据
    - headerData() 返回表格标题

    """

    def __init__(self, data):
        QAbstractTableModel.__init__(self)
        self.DataDisplay = data.copy()

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
        if 0 <= col < self.DataDisplay.shape[1]:
            col = index.column()
            row = index.row()
            self.beginResetModel()
            self.DataDisplay.iloc[row, col] = value
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
        return Qt.ItemFlags(QAbstractTableModel.flags(self, index) | Qt.ItemIsEditable | Qt.ItemIsSelectable)

    def InsertRowBelow(self, position=-1):
        self.beginResetModel()
        if position == -1:
            try:
                self.DataDisplay = self.DataDisplay.append(pd.DataFrame([[""] * self.DataDisplay.shape[1]], columns=self.DataDisplay.columns))
                self.DataDisplay = self.DataDisplay.reset_index(drop=True)
            except Exception as e:
                print(str(e))
        else:
            try:
                self.DataDisplay = pd.DataFrame(np.insert(self.DataDisplay.values, position+1, values=[[""] * self.DataDisplay.shape[1]], axis=0), columns=self.DataDisplay.columns)
                self.DataDisplay = self.DataDisplay.reset_index(drop=True)
            except Exception as e:
                print(str(e))
        self.endResetModel()

    def InsertRowAhead(self, position=0):
        self.beginResetModel()
        try:
            self.DataDisplay = pd.DataFrame(np.insert(self.DataDisplay.values, position, values=[[""] * self.DataDisplay.shape[1]], axis=0), columns=self.DataDisplay.columns)
            self.DataDisplay = self.DataDisplay.reset_index(drop=True)
        except Exception as e:
            print(str(e))
        self.endResetModel()

    def delRow(self, position=-1):
        self.beginResetModel()
        if position == -1:
            self.DataDisplay = self.DataDisplay.drop(labels=self.DataDisplay.shape[0]-1, axis=0)
            self.DataDisplay = self.DataDisplay.reset_index(drop=True)
        else:
            self.DataDisplay = self.DataDisplay.drop(labels=position, axis=0)
            self.DataDisplay = self.DataDisplay.reset_index(drop=True)
        self.endResetModel()

    def InsertColRight(self, position=-1):
        self.beginResetModel()
        if position == -1:
            try:
                self.DataDisplay.loc[:, self.DataDisplay.shape[1]] = ""
            except Exception as e:
                print(str(e))
        else:
            try:
                self.DataDisplay = pd.DataFrame(np.insert(self.DataDisplay.values, position+1, values=[[""] * self.DataDisplay.shape[0]], axis=1))
                self.DataDisplay.columns = [i for i in range(self.DataDisplay.shape[1])]
            except Exception as e:
                print(str(e))
        self.endResetModel()

    def InsertColLeft(self, position=0):
        self.beginResetModel()
        try:
            self.DataDisplay = pd.DataFrame(np.insert(self.DataDisplay.values, position, values=[[""] * self.DataDisplay.shape[0]], axis=1))
            self.DataDisplay.columns = [i for i in range(self.DataDisplay.shape[1])]
        except Exception as e:
            print(str(e))
        self.endResetModel()

    def delCol(self, position=-1):
        self.beginResetModel()
        try:
            if position == -1:
                print(self.DataDisplay.columns)
                self.DataDisplay = self.DataDisplay.drop(columns=self.DataDisplay.shape[1]-1, axis=1)
                self.DataDisplay.columns = [i for i in range(self.DataDisplay.shape[1])]
            else:
                print(self.DataDisplay.columns)
                self.DataDisplay = self.DataDisplay.drop(columns=position, axis=1)
                self.DataDisplay.columns = [i for i in range(self.DataDisplay.shape[1])]
        except Exception as e:
            print(str(e))
        self.endResetModel()
