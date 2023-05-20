from xlsx_pandas import xlsxToPandas
import pandas as pd
import numpy as np
import re
import xlrd

class ORM:
    def __init__(self):
        self.xp = xlsxToPandas()

    # 加载数据
    def excel_load(self, filename):
        data = self.xp.getData(filename)
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
        try:
            data[index_col] = range(1, len(data) + 1)
            return data
        except Exception as e:
            print("data_process_pd-reset_num")
            print(str(e))
            return False

    # 增添一行
    def add_line(self, data, new_line):
        """

        - new_line: DataFrame、dict、list类型，类型错误raise故障
        return: self.df对象->DataFrame
        """
        try:
            # new_line = pd.DataFrame([[]], columns=data.columns)
            data = data.append(new_line)
            # data = data.append(pd.Series(), ignore_index=True)
            data = data.reset_index(drop=True)
            # print(data.dtypes)
            return data
        except Exception as e:
            print("data_process_pd-add_line")
            print(str(e))
            return False

    # 删除一行
    def del_line(self, data, row_index):
        """

        :param row_index: 行序号，可为list
        :return: 已删除序列
        """
        try:
            res = data.loc[row_index]
            data = data.drop(labels=row_index, axis=0)
            data = data.reset_index(drop=True)
            # dataB = dataB.reset_index(drop=True)
            # dataA = self.reset_num(dataA)
            # data = self.reset_num(data, "序号")
            return True, data, res
        except Exception as e:
            print("data_process_pd-del_line")
            print(str(e))
            return False, False, False

    # 修改值
    def modify(self, data, row, col, contents):
        """

        - row: 行序号
        - col: 列序号
        - contents: 待修改的内容

        返回值：修改的行对象->DataFrame
        """
        try:
            data.iloc[row, col] = contents
        except Exception as e:
            print(str(e))
        return data

    # 搜索
    def search(self, data, col, reg, reg_flag=False):
        try:
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
        except Exception as e:
            print("data_process-search")
            print(str(e))
            return ""


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
