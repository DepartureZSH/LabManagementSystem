import sys
import pandas as pd
import numpy as np
import configparser
from PyQt5.QtWidgets import QApplication, QTableView, QWidget, QHeaderView
from PyQt5.QtCore import QAbstractTableModel, QModelIndex, Qt, pyqtSignal
from xlsx_pandas import xlsxToPandas


# 数据模型
class pandasModel(QAbstractTableModel):
    """
    QAbstractTableModel实例化需要：

    - rowCount() 返回行数
    - columnCount() 返回列数
    - data() 返回表格数据
    - headerData() 返回表格标题

    """

    def __init__(self):
        QAbstractTableModel.__init__(self)
        config = configparser.ConfigParser()
        config.read("configs/config.ini")
        Path = config.get('Path', 'Data')
        self.path = Path
        self.xp = xlsxToPandas()
        self.df1 = self.xp.getData(self.path+"实验室人员.xlsx")
        self.df2 = self.xp.getData(self.path+"实验室安排.xlsx")
        self.df3 = self.xp.getData(self.path+"实验室.xlsx")
        self.AdvancedUpdate()
        self.dataUpdate()

    def AdvancedUpdate(self):
        df4 = pd.merge(self.df2, self.df3, how="left")
        self.df = pd.merge(self.df1, df4, how="left")
        self.df = self.df[["序号", "实验室号", "实验室门牌号", "座位号", "工位位置X", "工位位置Y", "姓名", "学号", "导师"]]

    def setPath(self, Path):
        self.path = Path

    def setDf(self):
        self.df1 = self.xp.getData(self.path+"实验室人员.xlsx")
        self.df2 = self.xp.getData(self.path+"实验室安排.xlsx")
        self.df3 = self.xp.getData(self.path+"实验室.xlsx")
        df4 = pd.merge(self.df2, self.df3, how="left")
        self.df = pd.merge(self.df1, df4, how="left")
        self.df = self.df[["序号", "实验室号", "实验室门牌号", "座位号", "工位位置X", "工位位置Y", "姓名", "学号", "导师"]]

    def dataUpdate(self):
        self.beginResetModel()
        self.data = self.df
        self.pre_data = self.df[["实验室号", "座位号", "姓名", "学号", "导师"]]
        self.DataDisplay = self.pre_data
        self.endResetModel()

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
                Type = type(self.DataDisplay.iloc[index.row(), index.column()])
                if str(self.DataDisplay.iloc[index.row(), index.column()]) == "nan":
                    return ""
                elif Type == str:
                    return self.DataDisplay.iloc[index.row(), index.column()]
                elif Type == int or Type == np.int64 or Type == np.int32:
                    return str(self.DataDisplay.iloc[index.row(), index.column()])
                else:
                    # print(type(self.DataDisplay.iloc[index.row(), index.column()]))
                    return str(int(self.DataDisplay.iloc[index.row(), index.column()]))
        return None

    def headerData(self, col, orientation, role=Qt.DisplayRole):
        """
        表格标题

        :param col: 表格列数
        :param orientation: 标题朝向
        :param role: 表格的状态：Qt.DisplayRole——视图的单元格一般状态下（例如初始化显示时，编辑完成时），要显示的数据。
        :return: 各datas横向标题
        """
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.DataDisplay.columns[col]
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
        try:
            if index.isValid() and 0 <= index.row() < self.DataDisplay.shape[0] and value:
                col = index.column()
                if 1 <= col < self.DataDisplay.shape[1]:
                    self.beginResetModel()
                    col = self.DataDisplay.columns.values[col]
                    row = self.data.loc[index.row(), "序号"]
                    loc = self.df[self.df["序号"] == row].index.values
                    self.data.loc[index.row(), col] = value
                    self.pre_data = self.data[["实验室号", "座位号", "姓名", "学号", "导师"]]
                    self.DataDisplay = self.pre_data
                    self.df.loc[loc, col] = value
                    self.endResetModel()
                    return True
        except Exception as e:
            print("Exception")
            print(str(e))
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

    # 增
    def insertRows(self, position, rows=1, index=QModelIndex()):
        """
        插入一行
        :param position: 插入位置
        :param rows: 插入行数
        :param index:
        :return:
        """
        # 基础表格插入
        # position 插入位置；rows 插入行数
        self.beginResetModel()
        self.data = self.data.append(self.temp)
        self.data = self.data.reset_index(drop=True)
        self.pre_data = self.data[["实验室号", "座位号", "姓名", "学号", "导师"]]
        self.DataDisplay = self.pre_data
        self.df = self.df.append(self.temp)
        self.df = self.df.sort_values(by=['实验室号', '座位号'])
        self.endResetModel()

    # 删
    def removeRows(self, position, rows=1, index=QModelIndex):
        """
        删除一行
        :param position: 删除位置
        :param rows: 删除行数
        :param index:
        :return:
        """
        # position 删除位置；rows 删除行数
        self.beginRemoveRows(QModelIndex(), position, position + rows - 1)
        pass
        self.endRemoveRows()
        return True

    # 查
    def search(self, col, reg, reg_flag):
        if type(col) == str and type(reg) == str:
            df = self.data.fillna('')
            if reg_flag:
                if df[col].dtype == object:
                    res = df[df[col].str.contains(reg, na=False)]
                else:
                    res = df[df[col].astype(str).str.contains(str(reg), na=False) | (df[col].astype(str).str.contains(str(int(reg)), na=False))]
            else:
                if df[col].dtype == object:
                    res = df.loc[df[col] == reg]
                else:
                    res = df.loc[(df[col] == reg) | (df[col] == int(reg))]
            return res
        if type(col) == list and type(reg) == list:
            df = self.data
            for i in range(len(col)):
                if col[i] != "" and reg != "":
                    try:
                        temp = self.search(col[i], reg[i], reg_flag)
                        if i == 0:
                            df = temp
                        else:
                            df = pd.merge(df, temp)
                    except Exception as e:
                        print(str(e))
            self.beginResetModel()
            df = df.reset_index(drop=True)
            self.data = df
            self.pre_data = df[["实验室号", "座位号", "姓名", "学号", "导师"]]
            self.DataDisplay = self.pre_data
            self.endResetModel()
            return True
        return False

    # 保存人员名单
    def PrimarySave(self):
        try:
            self.df.to_excel(self.path + "实验室人员.xlsx", sheet_name='Sheet1', index=False, columns=self.df1.columns.values)
            self.df.to_excel(self.path + "实验室安排.xlsx", sheet_name='Sheet1', index=False, columns=self.df2.columns.values)
            return True
        except Exception as e:
            print(str(e))
            return False

    # 保存实验室名单
    def AdvancedSave(self):
        try:
            self.df3.to_excel(self.path + "实验室.xlsx", sheet_name='Sheet1', index=False, columns=self.df3.columns.values)
            return True
        except Exception as e:
            print("AdvancedSave")
            print(str(e))
            return False

    # 保存显示的名单
    def CustomizedSave(self, filename):
        try:
            col = ["实验室号", "实验室门牌号", "座位号", "姓名", "学号", "导师", "工位位置X", "工位位置Y"]
            self.data.to_excel(filename, sheet_name='Sheet1', index=False, columns=col)
            return True
        except Exception as e:
            print(str(e))
            return False
