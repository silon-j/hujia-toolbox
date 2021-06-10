from math import ceil
import sys, os, time, logging
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QMainWindow, QFileDialog, QMessageBox, 
QAction, QVBoxLayout, QGridLayout, QHBoxLayout, QWidget, QListView, QLineEdit)

from docx import Document
from openpyxl import Workbook
from utility import *

class MainWindow(QMainWindow):
    inputDirPath = ''
    inputDirFiles = []
    todoFiles = []
    progressBarValue = 0

    def __init__(self):
        super().__init__()
        self.init_ui()
        self.connect_action()

    def init_ui(self):
        self.setWindowTitle('word表格批量提取')
        self.setStyleSheet('#fileTypeTrans{background-color:white}')
        self.setWindowIcon(QIcon('../static/img/logo-circle.png'))
        self.setMinimumSize(480, 300)
        self.mainLayout = QGridLayout()

        # main content
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.centralwidget.setLayout(self.mainLayout)
        self.setCentralWidget(self.centralwidget)

        # select input dir btn
        self.inputDirBtn = QtWidgets.QPushButton(self.centralwidget)
        self.inputDirBtn.setObjectName("inputDirBtn")
        self.inputDirBtn.setText('选择输入文件夹')

        # input dir path
        self.inputDirEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.inputDirEdit.setObjectName("inputDirEdit")

        # start btn
        self.startBtn = QtWidgets.QPushButton(self.centralwidget)
        self.startBtn.setObjectName("startBtn")
        self.startBtn.setText('开始')

        # process bar
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")

        self.mainLayout.addWidget(self.inputDirBtn, 0 ,0)
        self.mainLayout.addWidget(self.inputDirEdit, 0, 2)
        self.mainLayout.addWidget(self.startBtn, 1, 2)
        self.mainLayout.addWidget(self.progressBar, 2, 2)

    def connect_action(self):
        # select input dir
        self.inputDirBtn.clicked.connect(self.set_inputdir)
        self.startBtn.clicked.connect(self.start_process)


    def set_inputdir(self):
        self.reset_process_setting()
        # clear line edit & set process to 0 when reselect input/output dir
        dirPath = QFileDialog.getExistingDirectory(
            self, "选择待处理文件所在目录", r"C:\Users\Administrator\Desktop")
        self.inputDirEdit.setText(dirPath)
        self.inputDirPath = dirPath
        self.inputDirFiles = os.listdir(dirPath)
        print(self.inputDirFiles)
        self.scan_inputdir()

    def scan_inputdir(self):
        '''检测输入文件夹，给出docx文件数量'''
        docx_cnt = 0
        for filename in self.inputDirFiles:
            if filename.endswith('.docx') and not filename.startswith('~$'):
                docx_cnt += 1
                self.todoFiles.append(filename)
        hint = QMessageBox.about(self, '请确认', '该文件夹中共发现docx文件{}个'.format(docx_cnt))

    def reset_process_setting(self):
        self.todoFiles = []
        self.progressBarValue = 0

    def start_process(self):
        startTime = time.localtime()
        check_dir('output')

        success_cnt = 0
        fail_cnt = 0

        wb_log = Workbook()
        ws_log = wb_log.worksheets[0]
        ws_log.append(['文件名','表格数','处理状态','备注'])

        wb = Workbook()
        dirPath = self.inputDirPath
        for todoName in self.todoFiles:
            filePath = os.path.join(os.path.abspath(dirPath), todoName)
            try:
                docx = Document(filePath)
                tables = docx.tables
                for index, table in enumerate(tables):
                    print('--------table---------')
                    table_data = []
                    for row in table.rows:
                        # 每行数据
                        # row_list = [cell.text for cell in row.cells]
                        # 每格数据
                        for cell in row.cells:
                            print(cell.text, '\t')
                            table_data.append(cell.text)
                            print(table_data)
                    wb.worksheets[index].append(table_data)
                    success_cnt += 1
                ws_log.append([todoName, len(tables), '成功', ''])
            except Exception as e:
                fail_cnt += 1
                ws_log.append([todoName, 0, '失败', e])
            
            self.progressBarValue = (success_cnt + fail_cnt) * 1.0 / len(self.todoFiles) * 100
            self.progressBar.setProperty("value", self.progressBarValue)


        outputName = '{}.xlsx'.format(time.strftime("%m-%d__%H-%M-%S", startTime))
        wb.save('./output/{}.xlsx'.format(outputName))
        wb_log.save('./output/{}_log.xlsx'.format(outputName))
        QMessageBox.about(self, '请确认', '''处理完成。<br>成功处理{}个文件，失败{}个文件\
            <br>已生成输出文件：{}<br>处理记录文件：{}'''.format(success_cnt, fail_cnt, outputName, outputName+'_log'))


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())