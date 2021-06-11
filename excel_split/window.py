import sys, os, time, logging
from PyQt5 import QtCore, QtWidgets
import PyQt5
from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QMainWindow, QFileDialog, QMessageBox, 
QAction, QPushButton, QVBoxLayout, QGridLayout, QHBoxLayout, QWidget, QListView, QLineEdit, QLabel)

from win32com import client as wc
from openpyxl import Workbook
from utility import *



class MainWindow(QMainWindow):
    

    def __init__(self):
        super().__init__()
        self.init_ui()
        self.connect_action()

    def init_ui(self):
        self.setWindowTitle('excel 拆分工具')
        self.setStyleSheet('#fileTypeTrans{background-color:white}')
        self.setWindowIcon(QIcon('../static/img/logo-circle.png'))
        self.setMinimumSize(480, 300)
        self.mainLayout = QGridLayout()

        # main content
        self.centralWidget = QWidget(self)
        self.centralWidget.setObjectName('centralWidget')
        self.centralWidget.setLayout(self.mainLayout)
        self.setCentralWidget(self.centralWidget)



class SplitWidget(QWidget):
    inputFileName = ''
    inputFilePath = ''

    def __init__(self):
        super().__init__()
        self.init_ui()
        self.connect_action()

    def init_ui(self):
        # self.centralWidget = QtWidgets.QWidget(self)
        # self.centralWidget.setObjectName("centralWidget")
        self.setMinimumSize(400,300)
        self.mainLayout = QGridLayout()

        # input file path button
        self.inputFileBtn = QPushButton()
        self.inputFileBtn.setText('选择需要拆分的 xlsx 文件')


        # input file box
        self.filePicLabel = QLabel()
        self.filePic = QPixmap("../static/img/excel.png")
        self.filePic_scaled = self.filePic.scaled(100, 100)
        self.filePicLabel.setPixmap(self.filePic_scaled)
        self.filePicLabel.setAlignment(QtCore.Qt.AlignCenter)

        # file name label
        self.fileNameLabel = QLabel()
        self.fileNameLabel.setText('未选择文件')
        self.fileNameLabel.setWordWrap(True)
        self.fileNameLabel.setAlignment(QtCore.Qt.AlignCenter)

        # start split btn
        self.startBtn = QPushButton()
        self.startBtn.setText('开始')

        # reset btn
        self.resetBtn = QPushButton()
        self.resetBtn.setText('重置')

        self.mainLayout.addWidget(self.inputFileBtn,0,0,1,1)
        self.mainLayout.addWidget(self.filePicLabel,1,0)
        self.mainLayout.addWidget(self.fileNameLabel,2,0)
        self.mainLayout.addWidget(self.startBtn,3,0)

        self.setLayout(self.mainLayout)

    def connect_action(self):
        self.inputFileBtn.clicked.connect(self.set_inputfile)
        self.startBtn.clicked.connect(self.start_split)

    def set_inputfile(self):
        self.inputFilePath, filter = QFileDialog.getOpenFileName(
            self, '选择待处理文件', r'C:\Users\Administrator\Desktop', 'excel 文件(*.xlsx)')
        self.inputFileName = os.path.basename(self.inputFilePath)

        self.fileNameLabel.setText(self.inputFilePath)
        print(self.inputFileName)


    def start_split(self):
        if self.inputFilePath:
            excel  = wc.Dispatch('Excel.Application')
            excel.DisplayAlerts = True
            excel.visible = True
            wb = excel.Workbooks.Open(self.inputFilePath)
            QMessageBox.about(self, '请确认', '1111111111')
        else:
            QMessageBox.warning(self, '提示','请选择待处理文件')





if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = SplitWidget()
    main.show()
    sys.exit(app.exec_())