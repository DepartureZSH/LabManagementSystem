import pandas as pd

# 数据表的pandas实现
class xlsxToPandas:

    def getData(self, OpenFileName):
        df = pd.read_excel(OpenFileName, header=0)
        return df
