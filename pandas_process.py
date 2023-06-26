import pandas as pd
import numpy as np
import re
import xlrd

class ORM:
    # 数据文件加载、新增一行、新增一列、重设“序号”索引、删除一行、删除一列、修改值
    # 加载数据
    def excel_load(self, filename):
        data = pd.read_excel(filename, header=0)
        return data

    # 添加一列
    def add_col(self, data, name, value):
        data[name] = value
        return data

    def index_set(self, data, col):
        data.insert(0, "序号", "0", allow_duplicates=False)
        data = data.set_index(col, drop=False)
        data = self.reset_num(data, col)
        data = data.reset_index(drop=True)
        return data

    # 重设序号
    def reset_num(self, data, index_col):
        data[index_col] = range(1, len(data) + 1)
        return data

    # 增添一行
    def add_line(self, data, new_line):
        """

        - new_line: DataFrame、dict、list类型，类型错误raise故障
        return: self.df对象->DataFrame
        """
        # new_line = pd.DataFrame([[]], columns=data.columns)
        data = data.append(new_line)
        # data = data.append(pd.Series(), ignore_index=True)
        data = data.reset_index(drop=True)
        # print(data.dtypes)
        return data

    # 删除一行
    def del_line(self, data, row_index):
        """

        :param row_index: 行序号，可为list
        :return: 已删除序列
        """
        res = data.loc[row_index]
        data = data.drop(labels=row_index, axis=0)
        data = data.reset_index(drop=True)
        # dataB = dataB.reset_index(drop=True)
        # dataA = self.reset_num(dataA)
        # data = self.reset_num(data, "序号")
        return True, data, res

    # 修改值
    def modify(self, data, row, col, contents):
        """

        - row: 行序号
        - col: 列序号
        - contents: 待修改的内容

        返回值：修改的行对象->DataFrame
        """
        data.iloc[row, col] = contents
        return data

    # 搜索
    def search(self, data, col, reg, reg_flag=False):
        # try:
        df = data.fillna('暂缺')
        # for each in df.columns.values:
        #     df[each] = df[each].apply(str)
        if reg_flag:
            # res = df[df[col].str.contains(reg, na=False)]
            if df[col].dtype == object:
                res = df[df[col].str.contains(reg, na=False)]
            else:
                res = df[df[col].str.contains(reg, na=False) | (df[col].str.contains(int(reg), na=False))]
        else:
            if df[col].dtype == object:
                res = df.loc[df[col] == reg]
            # res = df.loc[df[col] == reg]
            else:
                res = df.loc[(df[col] == reg) | (df[col] == int(reg))]
        return res
        # except Exception as e:
        #     return ""


# if __name__ == '__main__':
#     orm = ORM()
#     data, excel_info = orm.excel_load("data/人员名单.xlsx")
#     print(data.head())
#     data = orm.index_set(data, "序号")
#     print(data.head())
    # for each in excel_info:
    #     print(excel_info[each])
    # data = orm.add_col(data, "性别", "无")
    # print(data.head())
    # res = orm.search(data, "姓名", "张", reg_flag=True)
    # print(res)
    # res = orm.search(data, "实验室号", "2", reg_flag=False)
    # print(res)
