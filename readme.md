## 文件结构
本系统包含十个py模块和四个文件夹如下所示：
```cmd
│ Main.py //主界面模块
│ StudentTableModel.py //学生信息数据抽象表格模块
│ SeatTableModel.py //实验室工位图抽象表格模块
│ PandasProcess.py //Pandas ORM 模块
│ SettingWidget.py //设置界面模块
│ AddLineWidget.py //新增学生信息界面模块
│ ModifyWidget.py //修改学生信息界面模块
│ QuickUpdateWidget.py //一键更新界面模块
│ QRCode_Word.py //工位打印文档生成模块
│ ImageRecognition.py //表格识别模块
├─configs //配置文件夹
├─data //数据文件夹
├─imgs //图片文件夹
└─Outputs //输出文件夹
```
## 数据存储
本系统采用excel存储数据，其E-R图如下所示：
![数据实体图](https://github.com/DepartureZSH/LabManagementSystem/assets/91520014/43af47ae-d6fa-4153-9be6-0e10488d6082)

