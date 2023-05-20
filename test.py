import os
import configparser

path = "configs"
folder = os.path.exists(path)
if not folder:
    os.makedirs(path)
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
    with open(path + "/config.ini", 'w') as configfile:
        SettingConfig.write(configfile)
