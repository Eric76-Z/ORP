# coding=utf-8
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication

from history import orp_demo

####################### 全局变量#########################
app = QApplication(sys.argv)


class MyWindows(orp_demo.Ui_Form, QMainWindow):
    def __init__(self):
        super(MyWindows, self).__init__()
        self.setupUi(self)


my_windows = MyWindows()  # 实例化对象
my_windows.show()  # 显示窗口

sys.exit(app.exec_())
