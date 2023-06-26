import re
import pandas as pd
import numpy as np
import os
import configparser
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtWidgets import *

from qt_material import QtStyleTools


class setting(QWidget, QtStyleTools):
    signal_UpdateSetting = pyqtSignal()
    configsPath = "configs/config.ini"

    def __init__(self):
        super(setting, self).__init__()
        self.setWindowTitle("设置")
        self.resize(600, 600)
        self.ReadConfig()
        self.elements()
        self.Layout()
        self.Connection()
        self.StyleChange()

    @staticmethod
    def ConfigCheck():
        config = configparser.ConfigParser()
        SettingConfig = configparser.ConfigParser()
        SettingConfig["CSS"] = {
            'Main_Theme': "light_cyan_500.xml",
            'Main_BackGround': "True",
            'Menu_Theme': "light_blue.xml",
            'Menu_BackGround': "False",
            'AddLine_Theme': "light_cyan_500.xml",
            'AddLine_BackGround': "True",
            'Modify_Theme': 'light_cyan_500.xml',
            'Modify_Background': 'True'
        }
        SettingConfig['Path'] = {
            "Icon": "imgs/heu.png",
            "QR_temp": "imgs/temp.jpg",
            "Data": "data/",
            "Outputs": "OutPuts/",
            "Configs": "configs/config.ini"
        }
        path = "configs"
        folder = os.path.exists(path)
        if not folder:
            os.makedirs(path)
            with open(path + "/config.ini", 'w') as configfile:
                SettingConfig.write(configfile)
        else:
            config.read("configs/config.ini")
            if not config.sections():
                with open("configs/config.ini", 'w') as configfile:
                    SettingConfig.write(configfile)

    # 读取配置文件
    def ReadConfig(self):
        self.config = configparser.ConfigParser()
        self.SettingConfig = configparser.ConfigParser()
        try:
            self.config.read("configs/config.ini")
            if self.config.sections():
                self.SettingConfig["CSS"] = {
                    'Main_Theme': self.config.get('CSS', 'Main_Theme'),
                    'Main_BackGround': self.config.get('CSS', 'Main_BackGround'),
                    'Menu_Theme': self.config.get('CSS', 'Menu_Theme'),
                    'Menu_BackGround': self.config.get('CSS', 'Menu_BackGround'),
                    'AddLine_Theme': self.config.get('CSS', 'AddLine_Theme'),
                    'AddLine_BackGround': self.config.get('CSS', 'AddLine_BackGround'),
                    'Modify_Theme': self.config.get('CSS', 'Modify_Theme'),
                    'Modify_Background': self.config.get('CSS', 'Modify_Background')
                }
                self.SettingConfig['Path'] = {
                    "Icon": self.config.get('Path', 'Icon'),
                    "QR_temp": self.config.get('Path', 'QR_temp'),
                    "Data": self.config.get('Path', 'Data'),
                    "Outputs": self.config.get('Path', 'Outputs'),
                    "Configs": self.config.get('Path', 'Configs')
                }
            else:
                raise Exception("")
        except:
            self.SettingConfig["CSS"] = {
                'Main_Theme': "light_cyan_500.xml",
                'Main_BackGround': "True",
                'Menu_Theme': "light_blue.xml",
                'Menu_BackGround': "False",
                'AddLine_Theme': "light_cyan_500.xml",
                'AddLine_BackGround': "True",
                'Modify_Theme': 'light_cyan_500.xml',
                'Modify_Background': 'True'
            }
            self.SettingConfig['Path'] = {
                "Icon": "imgs/heu.png",
                "QR_temp": "imgs/temp.jpg",
                "Data": "data/",
                "Outputs": "OutPuts/",
                "Configs": "configs/config.ini"
            }
            self.save()

    # 组件库
    def elements(self):
        self.StackView = QStackedWidget()

        self.MenuView = QWidget()
        self.StyleButton = QPushButton("样式")
        self.PathButton = QPushButton("路径")
        self.DefaultButton = QPushButton("恢复默认")
        self.MenuViewUI()

        Style_Choices = ['dark_amber.xml', 'dark_blue.xml', 'dark_cyan.xml', 'dark_lightgreen.xml',
                         'dark_pink.xml', 'dark_purple.xml', 'dark_red.xml', 'dark_teal.xml',
                         'dark_yellow.xml', 'light_amber.xml', 'light_blue.xml', 'light_cyan.xml',
                         'light_cyan_500.xml', 'light_lightgreen.xml', 'light_pink.xml', 'light_purple.xml',
                         'light_red.xml', 'light_teal.xml', 'light_yellow.xml']
        self.StyleView = QWidget()
        self.ViewList = [QLabel("主界面"), QLabel("菜单界面"), QLabel("添加界面"), QLabel("修改界面")]
        self.ViewComboBox = {}
        self.BackgroundCheckbox = {}
        item = self.config.items('CSS')
        for key, value in item:
            if value == "True":
                self.BackgroundCheckbox[key] = QCheckBox("背景反转", self)
                self.BackgroundCheckbox[key].setCheckState(2)
            elif value == "False":
                self.BackgroundCheckbox[key] = QCheckBox("背景反转", self)
                self.BackgroundCheckbox[key].setCheckState(0)
            else:
                self.ViewComboBox[key] = QComboBox()
                self.ViewComboBox[key].addItems(Style_Choices)
                self.ViewComboBox[key].setCurrentText(value)
        self.StyleApplyButton = QPushButton("应用")
        self.StyleApplyButton.setEnabled(False)
        self.StyleSaveButton = QPushButton("保存")
        self.StyleViewUI()

        self.StackView.addWidget(self.StyleView)

        self.PathView = QWidget()
        item = self.config.items('Path')
        self.PathLabels = [QLabel("图标路径"), QLabel("二维码暂存路径"), QLabel("数据文件夹路径"), QLabel("输出文件夹路径"), QLabel("配置文件路径")]
        self.PathEdit = {}
        for key, value in item:
            self.PathEdit[key] = QLineEdit()
            self.PathEdit[key].setText(value)
        self.PathApplyButton = QPushButton("应用")
        self.PathApplyButton.setEnabled(False)
        self.PathSaveButton = QPushButton("保存")
        self.PathViewUI()
        self.StackView.addWidget(self.PathView)

        self.StackView.setCurrentIndex(0)

    # 布局
    def Layout(self):
        HBoxLayout = QHBoxLayout()
        HBoxLayout.addWidget(self.MenuView)
        HBoxLayout.addWidget(self.StackView)
        self.setLayout(HBoxLayout)

    # 信号槽
    def Connection(self):
        self.StyleButton.clicked.connect(self.StackViewChange)
        self.PathButton.clicked.connect(self.StackViewChange)
        self.DefaultButton.clicked.connect(self.Default)
        self.BackgroundCheckbox["main_background"].clicked.connect(lambda: self.Background("Main_BackGround"))
        self.BackgroundCheckbox["menu_background"].clicked.connect(lambda: self.Background("Menu_BackGround"))
        self.BackgroundCheckbox["addline_background"].clicked.connect(lambda: self.Background("AddLine_BackGround"))
        self.BackgroundCheckbox["modify_background"].clicked.connect(lambda: self.Background("Modify_Background"))
        self.ViewComboBox["main_theme"].currentIndexChanged.connect(lambda: self.View("Main_Theme"))
        self.ViewComboBox["menu_theme"].currentIndexChanged.connect(lambda: self.View("Menu_Theme"))
        self.ViewComboBox["addline_theme"].currentIndexChanged.connect(lambda: self.View("AddLine_Theme"))
        self.ViewComboBox["modify_theme"].currentIndexChanged.connect(lambda: self.View("Modify_Theme"))
        self.PathEdit["icon"].textChanged.connect(lambda: self.Path("Icon"))
        self.PathEdit["qr_temp"].textChanged.connect(lambda: self.Path("QR_temp"))
        self.PathEdit["data"].textChanged.connect(lambda: self.Path("Data"))
        self.PathEdit["outputs"].textChanged.connect(lambda: self.Path("Outputs"))
        self.PathEdit["configs"].textChanged.connect(lambda: self.Path("Configs"))
        self.StyleApplyButton.clicked.connect(self.Apply)
        self.PathApplyButton.clicked.connect(self.Apply)
        self.StyleSaveButton.clicked.connect(self.save)
        self.PathSaveButton.clicked.connect(self.save)

    # 背景虚化
    def Background(self, slot):
        if self.BackgroundCheckbox[slot.lower()].isChecked():
            self.SettingConfig['CSS'][slot] = 'True'
        else:
            self.SettingConfig['CSS'][slot] = 'False'
        self.StyleApplyButton.setEnabled(True)

    # 背景改变
    def View(self, slot):
        self.SettingConfig['CSS'][slot] = self.ViewComboBox[slot.lower()].currentText()
        self.StyleApplyButton.setEnabled(True)

    # 路径改变
    def Path(self, slot):
        self.SettingConfig['Path'][slot] = self.PathEdit[slot.lower()].text()
        self.PathApplyButton.setEnabled(True)

    # 保存
    def save(self):
        with open('configs/config.ini', 'w') as configfile:
            self.SettingConfig.write(configfile)

    # 应用
    def Apply(self):
        self.save()
        self.signal_UpdateSetting.emit()
        self.StyleChange()

    # 样式改变
    def StyleChange(self):
        config = configparser.ConfigParser()
        config.read("configs/config.ini")
        StackView_theme = config["CSS"]["Main_Theme"]
        StackView_background = ("True" == config["CSS"]["Modify_BackGround"])
        self.apply_stylesheet(self.StackView, theme=StackView_theme, invert_secondary=StackView_background)
        MenuView_theme = config["CSS"]["Menu_Theme"]
        MenuView_background = ("True" == config["CSS"]["Menu_BackGround"])
        self.apply_stylesheet(self.MenuView, theme=MenuView_theme, invert_secondary=MenuView_background)
        self.setWindowIcon(QIcon(config["Path"]["Icon"]))

    # 菜单栏界面
    def MenuViewUI(self):
        VBoxLayout = QVBoxLayout()
        VBoxLayout.addWidget(self.StyleButton)
        VBoxLayout.addWidget(self.PathButton)
        vSpacer = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        VBoxLayout.addItem(vSpacer)
        VBoxLayout.addWidget(self.DefaultButton)
        self.MenuView.setLayout(VBoxLayout)

    # 风格界面
    def StyleViewUI(self):
        option = self.config.options('CSS')
        GridLayout = QGridLayout()
        for i in range(len(self.ViewList)):
            GridLayout.addWidget(self.ViewList[i], i, 0)
            key = option.pop(0)
            GridLayout.addWidget(self.ViewComboBox[key], i, 1)
            key = option.pop(0)
            GridLayout.addWidget(self.BackgroundCheckbox[key], i, 2)
        HBoxLayout = QHBoxLayout()
        hSpacer = QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        HBoxLayout.addItem(hSpacer)
        HBoxLayout.addWidget(self.StyleApplyButton)
        HBoxLayout.addWidget(self.StyleSaveButton)
        vSpacer = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        VBoxLayout = QVBoxLayout()
        VBoxLayout.addLayout(GridLayout)
        VBoxLayout.addItem(vSpacer)
        VBoxLayout.addLayout(HBoxLayout)
        self.StyleView.setLayout(VBoxLayout)

    # 路径界面
    def PathViewUI(self):
        item = self.config.items('Path')
        GridLayout = QGridLayout()
        for i in range(len(self.PathLabels)):
            GridLayout.addWidget(self.PathLabels[i], i, 0)
            key, value = item[i]
            GridLayout.addWidget(self.PathEdit[key], i, 1)
        HBoxLayout = QHBoxLayout()
        hSpacer = QSpacerItem(20, 20, QSizePolicy.Expanding, QSizePolicy.Minimum)
        HBoxLayout.addItem(hSpacer)
        HBoxLayout.addWidget(self.PathApplyButton)
        HBoxLayout.addWidget(self.PathSaveButton)
        vSpacer = QSpacerItem(20, 20, QSizePolicy.Minimum, QSizePolicy.Expanding)
        VBoxLayout = QVBoxLayout()
        VBoxLayout.addLayout(GridLayout)
        VBoxLayout.addItem(vSpacer)
        VBoxLayout.addLayout(HBoxLayout)
        self.PathView.setLayout(VBoxLayout)

    # 页面变换
    def StackViewChange(self):
        if self.StackView.currentIndex() == 0:
            self.StackView.setCurrentIndex(1)
        elif self.StackView.currentIndex() == 1:
            self.StackView.setCurrentIndex(0)

    # 恢复默认
    def Default(self):
        self.SettingConfig["CSS"] = {
            'Main_Theme': "light_cyan_500.xml",
            'Main_BackGround': "True",
            'Menu_Theme': "light_blue.xml",
            'Menu_BackGround': "False",
            'AddLine_Theme': "light_cyan_500.xml",
            'AddLine_BackGround': "True",
            'Modify_Theme': 'light_cyan_500.xml',
            'Modify_Background': 'True'
        }
        self.SettingConfig['Path'] = {
            "Icon": "imgs/heu.png",
            "QR_temp": "imgs/temp.jpg",
            "Data": "data/",
            "Outputs": "OutPuts/",
            "Configs": "configs/config.ini"
        }
        with open('configs/config.ini', 'w') as configfile:
            self.SettingConfig.write(configfile)
        self.signal_UpdateSetting.emit()
        self.StyleChange()

# if __name__ == '__main__':
#     app = QApplication(sys.argv)
#     demo = setting()
#     demo.show()
#     sys.exit(app.exec_())
