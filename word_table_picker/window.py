import sys, os, time, logging
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtWidgets import (QApplication, QComboBox, QMainWindow, QFileDialog, QMessageBox, 
QAction, QWidget, QListView, QLineEdit)


class MainWindow(QMainWindow):
    inputDirPath = ''

    def __init__(self):
        super().__init__()
        self.init_ui()
        self.connect_action()
    

    def init_ui(self):
        self.setMinimumSize(300, 500)

        # main content
        self.centralwidget = QWidget(self)
        self.centralwidget.setObjectName("centralwidget")

        # select input dir btn
        self.inputDirBtn = QtWidgets.QPushButton(self.centralwidget)
        self.inputDirBtn.setGeometry(QtCore.QRect(340, 100, 100, 23))
        self.inputDirBtn.setObjectName("inputDirBtn")
        self.inputDirBtn.setText('选择输入文件夹')

        # input dir path
        self.inputDirEdit = QtWidgets.QLineEdit(self.centralwidget)
        self.inputDirEdit.setGeometry(QtCore.QRect(60, 100, 260, 20))
        self.inputDirEdit.setObjectName("inputDirEdit")

    def connect_action(self):
        # select input dir
        self.inputDirBtn.clicked.connect(self.set_inputdir)


    def set_inputdir(self):
        # clear line edit & set process to 0 when reselect input/output dir
        dirPath = QFileDialog.getExistingDirectory(
            self, "选择待处理文件所在目录", r"C:\Users\Administrator\Desktop")
        self.inputDirEdit.setText(dirPath)



if __name__ == "__main__":
    app = QApplication(sys.argv)
    main = MainWindow()
    main.show()
    sys.exit(app.exec_())