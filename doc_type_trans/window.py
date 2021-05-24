import sys, os, time
from PyQt5 import QtCore, QtWidgets ,QtGui
from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QIcon
from PyQt5.QtWidgets import (QApplication, QComboBox, QMainWindow, QFileDialog, QMessageBox, 
QAction, QWidget, QListView)

from openpyxl import Workbook, worksheet
from win32com import client as wc


def check_dir(dir_path):
    if not os.path.exists(dir_path):
        os.makedirs(dir_path)


class WindowFileConvert(QWidget):

    # 文件夹内所有文件
    inputDirFiles = []
    # 应处理文件
    todoFiles = []
    # 未处理文件
    passedFiles = []
    # 转换正确的文件总数
    successFiles = []
    # 装换错误的文件总数
    errorFiles = []
    # 生成的文件名
    logFileName = 'log.txt'
    # 当前功能编号
    currentFnIndex = 0

    def __init__(self) -> None:
        super().__init__()
        self.init_ui()
        self.connect_action()

    def init_ui(self):
        # window setting
        self.setObjectName("fileTypeTrans")
        self.setStyleSheet('#fileTypeTrans{background-color:white}')
        self.setFixedSize(500, 500)
        self.setWindowTitle('word文档类型转换')
        self.setWindowIcon(QIcon('static/img/logo-circle.png'))

        # main content
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")

        # function select box
        self.fnSelector = QComboBox(self.centralwidget)
        self.fnSelector.setView(QListView())
        self.fnSelector.setObjectName('transTypeSelector')
        self.fnSelector.setGeometry(QtCore.QRect(60, 60, 260, 28))
        self.fnSelector.addItems(['doc 转 docx', 'docx 转 doc', 'xls 转 xlsx', 'xlsx 转 xls'])
        self.fnSelector.setStyleSheet('''
            #transTypeSelector
            {
                background: #eeeeee;
                padding-left:5px ;
                border-radius:3px;
                border: 1px solid gray;
            }
            #transTypeSelector QAbstractItemView::item
            {
                height:40px;
            }
        ''')
        # log dir open btn
        self.logDirBtn = QtWidgets.QPushButton(self.centralwidget)
        self.logDirBtn.setGeometry(QtCore.QRect(340, 60, 100, 23))
        self.logDirBtn.setObjectName("logDirBtn")
        self.logDirBtn.setText('处理记录')

        # input dir path
        self.inputDirEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.inputDirEdit.setGeometry(QtCore.QRect(60, 100, 260, 20))
        self.inputDirEdit.setObjectName("inputDirEdit")
        # select input dir btn
        self.inputDirBtn = QtWidgets.QPushButton(self.centralwidget)
        self.inputDirBtn.setGeometry(QtCore.QRect(340, 100, 100, 23))
        self.inputDirBtn.setObjectName("inputDirBtn")
        self.inputDirBtn.setText('选择输入文件夹')
        # output dir
        self.outputDirEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.outputDirEdit.setGeometry(QtCore.QRect(60, 130, 260, 20))
        self.outputDirEdit.setObjectName("outputDirEdit")
        # select output dir btn
        self.outputDirBtn = QtWidgets.QPushButton(self.centralwidget)
        self.outputDirBtn.setGeometry(QtCore.QRect(340, 130, 100, 23))
        self.outputDirBtn.setObjectName("outputDirBtn")
        self.outputDirBtn.setText('选择输出文件夹')
        # process log
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setGeometry(QtCore.QRect(60, 180, 380, 190))
        self.textEdit.setObjectName("textEdit")
        # process bar
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(60, 400, 270, 23))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        # start btn
        self.startBtn = QtWidgets.QPushButton(self.centralwidget)
        self.startBtn.setGeometry(QtCore.QRect(340, 400, 100, 23))
        self.startBtn.setObjectName("startBtn")
        self.startBtn.setText('开始转换')
        # 
        # self.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(self)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 500, 23))
        self.menubar.setObjectName("menubar")
        QtCore.QMetaObject.connectSlotsByName(self)

    def connect_action(self):
        # select current function
        self.fnSelector.currentIndexChanged[int].connect(self.select_function)
        # open log dir
        self.logDirBtn.clicked.connect(self.open_logdir)
        # select input dir
        self.inputDirBtn.clicked.connect(self.set_inputdir)
        # select output dir
        self.outputDirBtn.clicked.connect(self.set_outputdir)
        # if input or output dir edit content change, set Gui(process .etc) default
        self.inputDirEdit.textChanged.connect(self.init_ui_status)
        self.outputDirEdit.textChanged.connect(self.init_ui_status)
        # start function
        self.startBtn.clicked.connect(self.start_convert)

    def log_process_msg(self):
        self.textEdit.append("正在生成处理记录excel文件\n")
        logExcelPath = os.path.join('./log', '{}.xlsx'.format(self.logFileName.split('.')[0]))

        # gen log excel
        wb = Workbook()
        ws = wb.worksheets[0]
        ws.append(['文件名', '处理状态', '备注'])

        for filename in self.errorFiles:
            ws.append([filename, '处理失败', ''])
        for filename in self.successFiles:
            ws.append([filename, '处理成功', ''])
        for filename in self.passedFiles:
            ws.append([filename, '未处理', '不符合处理文件拓展名'])

        wb.save(logExcelPath)
        self.textEdit.append("转换完成，请到程序目录下log文件夹中 (log_处理时间.xlsx) 查看详细信息")

    def set_ui_enabled(self, flag):
        self.inputDirEdit.setEnabled(flag)
        self.inputDirBtn.setEnabled(flag)
        self.outputDirEdit.setEnabled(flag)
        self.outputDirBtn.setEnabled(flag)
        self.startBtn.setEnabled(flag)

    def open_logdir(self):
        path = './log'
        os.startfile(os.path.abspath(path))

    def set_inputdir(self):
        # clear line edit & set process to 0 when reselect input/output dir
        dirPath = QFileDialog.getExistingDirectory(
            self, "选择待处理文件所在目录", r"C:\Users\Administrator\Desktop")
        self.inputDirEdit.setText(dirPath)

    def set_outputdir(self):
        dirPath = QFileDialog.getExistingDirectory(
            self, "选择待处理文件所在目录", r"C:\Users\Administrator\Desktop")
        self.outputDirEdit.setText(dirPath)

    def init_ui_status(self):
        """ reset all process status and log box
        """
        self.progressBar.setProperty("value", 0)
        self.textEdit.setText("")

    def init_data(self):
        # input dir all file count
        self.inputDirFiles = []
        # 应处理文件
        self.todoFiles = []
        # 未处理文件
        self.passedFiles = []
        # 转换正确的文件总数
        self.successFiles = []
        # 装换错误的文件总数
        self.errorFiles = []
        # clear 处理记录
        self.textEdit.clear()


    def select_function(self, functionIndex):
        # function select
        self.currentFnIndex=functionIndex

    def start_convert(self):
        fnMap = [
            {'index':0, 'function':self.convert_to_docx, 'desc':'doc 转 docx'},
            {'index':1, 'function':self.convert_to_doc, 'desc':'docx 转 doc'},
            {'index':2, 'function':self.convert_to_xlsx, 'desc':'xls 转 xlsx'},
            {'index':3, 'function':self.convert_to_xls, 'desc':'xlsx 转 xls'},
        ]
        # check dir
        inputUrl = self.inputDirEdit.text().strip()
        outputUrl = self.outputDirEdit.text().strip()
        if inputUrl == '' or outputUrl == '':
            msgBox = QMessageBox(QMessageBox.Warning, "警告", "请选中要转换的目录或目标目录")
            msgBox.exec()
            return
        # set all button and input disabled, reset data
        self.set_ui_enabled(False)
        self.init_data()

        # gen process log text file
        check_dir('./log')
        startTime = time.localtime()
        self.logFileName = 'log__{}.txt'.format(time.strftime("%m-%d__%H-%M-%S", startTime))
        logFilePath = os.path.join('./log', self.logFileName)
        # Additional written to log text
        with open(logFilePath, "a") as logFile:
            # wont process files in sub dirs
            self.inputDirFiles = os.listdir(inputUrl)
            startHint='{} \n开始转换： \n 当前转换模式为：{} \n'.format(
                time.strftime("%Y-%m-%d__%H:%M:%S", startTime),
                fnMap[self.currentFnIndex]["desc"])
            logFile.write(startHint)
            self.textEdit.append(startHint)
            # function processing
            fnMap[self.currentFnIndex]["function"](inputUrl, outputUrl, logFile)
            self.log_process_msg()
            finishHint = """\n\n转换完成！\n 文件夹中文件：{}个 \n 应处理文件数：{} \n
            成功处理数：{} \n 失败数：{}
            """.format(len(self.inputDirFiles), len(self.todoFiles), len(self.successFiles), len(self.errorFiles))
            self.textEdit.append(finishHint)
            logFile.write(finishHint)
        self.set_ui_enabled(True)


    def convert_to_docx(self, inputUrl, outputUrl, logFile):
        totalCnt = len(self.inputDirFiles)
        processCnt = 0

        word = wc.Dispatch("Word.Application")
        word.DisplayAlerts = 0
        word.visible = 0

        for filename in self.inputDirFiles:
            processCnt += 1
            # processbar fresh render
            processBarValue = processCnt * 1.0 / totalCnt * 100
            self.progressBar.setValue(int(processBarValue))
            if filename.endswith('.doc') and not filename.startswith('~$'):
                self.todoFiles.append(filename)

                filePath = os.path.join(os.path.abspath(inputUrl), filename)
                convertHint = '正在处理doc文件：{} \n'.format(filename)
                self.textEdit.append(convertHint)
                logFile.write(convertHint)
                try:
                    doc = word.Documents.Open(filePath)
                    # split filename
                    rename = os.path.splitext(filename)
                    # gen a .docx file, 12 means docx format
                    targetPath = os.path.join(os.path.abspath(outputUrl), '{}.docx'.format(rename[0]))
                    doc.SaveAs(targetPath, 12)
                    doc.Close()
                    self.successFiles.append(filename)
                except Exception as e:
                    # error log
                    errorHint = '转换doc至docx失败，请检查该文件：{}\n'.format(filename)
                    logFile.write(errorHint)
                    logFile.write('error detail: \n {}'.format(e))
                    # error show in app
                    self.textEdit.append(errorHint)
                    # record
                    self.errorFiles.append(filename)
            else:
                self.passedFiles.append(filename)
        word.Quit()



    def convert_to_doc(self, inputUrl, outputUrl, logFile):
        word = wc.Dispatch("Word.Application")
        word.DisplayAlerts = 0
        word.visible = 0

        totalCnt = len(self.inputDirFiles)
        processCnt = 0

        for filename in self.inputDirFiles:
            processCnt += 1
            # processbar fresh render
            processBarValue = processCnt * 1.0 / totalCnt * 100
            self.progressBar.setValue(int(processBarValue))
            if filename.endswith('.docx') and not filename.startswith('~$'):
                self.todoFiles.append(filename)

                filePath = os.path.join(os.path.abspath(inputUrl), filename)
                convertHint = '正在处理docx文件：{} \n'.format(filename)
                self.textEdit.append(convertHint)
                logFile.write(convertHint)
                try:
                    docx = word.Documents.Open(filePath)
                    # split filename
                    rename = os.path.splitext(filename)
                    # gen a .doc file, 0 means doc format
                    targetPath = os.path.join(os.path.abspath(outputUrl), '{}.doc'.format(rename[0]))
                    docx.SaveAs(targetPath, 0)
                    docx.Close()
                    self.successFiles.append(filename)
                except Exception as e:
                    # error log
                    errorHint = '转换docx至doc失败，请检查该文件：{}\n'.format(filename)
                    logFile.write(errorHint)
                    logFile.write('error detail: \n {}'.format(e))
                    # error show in app
                    self.textEdit.append(errorHint)
                    # record
                    self.errorFiles.append(filename)
            else:
                self.passedFiles.append(filename)
        word.Quit()

    def convert_to_xlsx(self, inputUrl, outputUrl, logFile):
        excel  = wc.Dispatch('Excel.Application')
        # excel.DisplayAlerts = 0
        # excel.visible = 0

        totalCnt = len(self.inputDirFiles)
        processCnt = 0

        for filename in self.inputDirFiles:
            processCnt += 1
            # processbar fresh render
            processBarValue = processCnt * 1.0 / totalCnt * 100
            self.progressBar.setValue(int(processBarValue))

            if filename.endswith('.xls') and not filename.startswith('~$'):
                self.todoFiles.append(filename)

                filePath = os.path.join(os.path.abspath(inputUrl), filename)
                convertHint = '正在转换xls文件至xlsx：{} \n'.format(filename)
                self.textEdit.append(convertHint)
                logFile.write(convertHint)

                try:
                    wb = excel.Workbooks.Open(filePath)
                    # split filename
                    rename = os.path.splitext(filename)
                    # gen a .xlsx file
                    # FileFormat = 51 is for .xlsx extension
                    # FileFormat = 56 is for .xls extension
                    targetPath = os.path.join(os.path.abspath(outputUrl), '{}.xlsx'.format(rename[0]))
                    wb.SaveAs(targetPath, 51)
                    wb.Close()
                    self.successFiles.append(filename)
                except Exception as e:
                    # error log
                    errorHint = '转换xls至xlsx失败，请检查该文件：{}\n'.format(filename)
                    logFile.write(errorHint)
                    logFile.write('error detail: \n {}'.format(e))
                    # error show in app
                    self.textEdit.append(errorHint)
                    # record
                    self.errorFiles.append(filename)
            else:
                self.passedFiles.append(filename)
        excel.Application.Quit()

    def convert_to_xls(self, inputUrl, outputUrl, logFile):
        excel  = wc.Dispatch('Excel.Application')
        # excel.DisplayAlerts = 0
        # excel.visible = 0

        totalCnt = len(self.inputDirFiles)
        processCnt = 0

        for filename in self.inputDirFiles:
            processCnt += 1
            # processbar fresh render
            processBarValue = processCnt * 1.0 / totalCnt * 100
            self.progressBar.setValue(int(processBarValue))

            if filename.endswith('.xlsx') and not filename.startswith('~$'):
                self.todoFiles.append(filename)

                filePath = os.path.join(os.path.abspath(inputUrl), filename)
                convertHint = '正在转换xlsx文件至xls：{} \n'.format(filename)
                self.textEdit.append(convertHint)
                logFile.write(convertHint)
                try:
                    wb = excel.Workbooks.Open(filePath)
                    # split filename
                    rename = os.path.splitext(filename)
                    # gen a .xlsx file
                    # FileFormat = 51 is for .xlsx extension
                    # FileFormat = 56 is for .xls extension
                    targetPath = os.path.join(os.path.abspath(outputUrl), '{}.xls'.format(rename[0]))
                    print(targetPath)
                    wb.SaveAs(targetPath, 56)
                    wb.Close()
                    self.successFiles.append(filename)
                except Exception as e:
                    # error log
                    errorHint = '转换xlsx至xls失败，请检查该文件：{}\n'.format(filename)
                    logFile.write(errorHint)
                    logFile.write('error detail: \n {}'.format(e))
                    # error show in app
                    self.textEdit.append(errorHint)
                    # record
                    self.errorFiles.append(filename)
            else:
                self.passedFiles.append(filename)
        excel.Application.Quit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = WindowFileConvert()
    main.show()
    sys.exit(app.exec_())