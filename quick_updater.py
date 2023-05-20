import re
import pandas as pd
import numpy as np
import configparser
from qt_material import QtStyleTools
from PyQt5.QtCore import Qt, pyqtSignal
from PyQt5.QtGui import QIcon, QColor, QFont
from PyQt5.QtWidgets import QWidget
from PyQt5.QtWidgets import QHBoxLayout, QVBoxLayout, QFormLayout, QComboBox
from PyQt5.QtWidgets import QLabel, QMessageBox, QPushButton, QLineEdit

class UpdateTool(QWidget, QtStyleTools):
    def __init__(self, df, df3):
        super(UpdateTool, self).__init__()