import sys
import os
import math
import exchangelib.services
import psutil
import time
from exchangelib import (
    DELEGATE, Credentials, Account, EWSTimeZone, EWSDateTime, FileAttachment,
)
from exchangelib.errors import UnauthorizedError, EWSError

from openpyxl import load_workbook, Workbook

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QPainter, QColor, QFont, QIcon
from PyQt5.QtWidgets import QMainWindow, QTextEdit, QWidget, QVBoxLayout, QApplication, QLabel, QDesktopWidget, QHBoxLayout, QFormLayout, \
    QPushButton, QLineEdit, QMessageBox, QAction, qApp
from PyQt5.QtWebEngineWidgets import *

ACCOUNT = None

class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # 设置窗口大小
        self.setObjectName("loginWindow")
        self.setStyleSheet('#loginWindow{background-color:white}')
        self.setFixedSize(650, 440)
        self.setWindowTitle("登录邮箱")
        self.setWindowIcon(QIcon('static/img/logo-circle.png'))

        # banner
        banner = QPixmap("static/img/window_banner2.png")
        rape_banner = banner.scaled(650, 140)
        label = QLabel(self)
        label.setPixmap(rape_banner)

        # 绘制 banner 文字
        text = "exchange邮箱用户登录"
        lbl_logo = QLabel(self)
        lbl_logo.setText(text)
        lbl_logo.setStyleSheet("QWidget{color:white;font-weight:600;background: transparent;font-size:30px;}")
        lbl_logo.setFont(QFont("Microsoft YaHei"))
        lbl_logo.move(150, 50)
        lbl_logo.setAlignment(Qt.AlignCenter)
        lbl_logo.raise_()

        # 登录表单内容部分
        login_widget = QWidget(self)
        login_widget.move(0, 140)
        login_widget.setGeometry(0, 140, 650, 260)

        hbox = QHBoxLayout()
        # 添加左侧logo
        logolb = QLabel(self)
        logopix = QPixmap("static/img/logo-circle.png")
        logopix_scared = logopix.scaled(100, 100)
        logolb.setPixmap(logopix_scared)
        logolb.setAlignment(Qt.AlignCenter)
        hbox.addWidget(logolb, 1)
        # 添加右侧表单
        fmlayout = QFormLayout()
        lbl_email = QLabel("邮箱")
        lbl_email.setFont(QFont("Microsoft YaHei"))
        self.input_email = QLineEdit()
        self.input_email.setFixedWidth(270)
        self.input_email.setFixedHeight(38)

        lbl_pwd = QLabel("密码")
        lbl_pwd.setFont(QFont("Microsoft YaHei"))
        self.input_pwd = QLineEdit()
        self.input_pwd.setEchoMode(QLineEdit.Password)
        self.input_pwd.setFixedWidth(270)
        self.input_pwd.setFixedHeight(38)

        self.btn_login = QPushButton("登录")
        self.btn_login.setFixedWidth(270)
        self.btn_login.setFixedHeight(40)
        self.btn_login.setFont(QFont("Microsoft YaHei"))
        self.btn_login.setObjectName("login_btn")
        self.btn_login.setStyleSheet("#login_btn{background-color:#2c7adf;color:#fff;border:none;border-radius:4px;}")
        self.btn_login.clicked.connect(self.login)

        fmlayout.addRow(lbl_email, self.input_email)
        fmlayout.addRow(lbl_pwd, self.input_pwd)
        fmlayout.addWidget(self.btn_login)

        hbox.setAlignment(Qt.AlignCenter)
        # 调整间距
        fmlayout.setHorizontalSpacing(20)
        fmlayout.setVerticalSpacing(12)

        hbox.addLayout(fmlayout, 2)
        login_widget.setLayout(hbox)

        # 底部状态提示
        self.status_widget = QWidget(self)
        self.status_widget.move(0, 380)
        self.status_widget.setGeometry(0, 380, 650, 40)

        self.statusbar = QLabel(self)
        self.statusbar.setText('')
        self.statusbar.setStyleSheet("QWidget{color:white;font-weight:36;background: white;font-size:16px;}")
        self.statusbar.setAlignment(Qt.AlignCenter)
        self.statusbar.raise_()

        status_box = QVBoxLayout()
        status_box.addWidget(self.statusbar)
        self.status_widget.setLayout(status_box)
        self.center()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def login(self):
        em_ad = self.input_email.text()
        em_psw =  self.input_pwd.text()
        em_username = em_ad.split('@')[0]
        em_domain = em_ad.split('@')[-1].split('.')[0]
        cred = Credentials(r'{}\{}'.format(em_domain, em_username), em_psw)
        try:
            a = Account(
                primary_smtp_address=em_ad,
                credentials=cred,
                autodiscover=True,
                access_type=DELEGATE
            )
            if a:
                global ACCOUNT
                ACCOUNT = a
                self.statusbar.setStyleSheet('QWidget{color:white; background: green}')
                self.statusbar.setText('>>>你的邮箱已连接,即将进入<<<')
                time.sleep(3)
                self.close()
                mainWindow.show()
            else:
                print('未能连接到该账户。请重试...')
        except EWSError:
            self.statusbar.setStyleSheet('QWidget{color:white; background: orange}')
            self.statusbar.setText('账号密码错误，请重试')
        except Exception as e:
            print('连接邮箱失败,请检查账户或网络：\n{e}\n'.format(e=e))
            self.statusbar.setStyleSheet('QWidget{color:white; background: orange}')
            self.statusbar.setText('连接邮箱失败,请检查网络或账户')


class MainWindow(QMainWindow):

    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        # 设置窗口大小
        # self.resize(1280, 720)
        self.setMinimumWidth(1280)
        self.setMinimumHeight(960)
        # 窗口居中
        self.center()
        # 设置窗口的标题
        self.setWindowTitle('我的邮箱')
        # 设置窗口的图标，引用当前目录下的web.png图片
        self.setWindowIcon(QIcon('static/img/logo192.png'))
        # 初始化菜单，状态栏
        self.init_statusbar()
        self.init_menubar()
        # 初始化layout
        self.init_layout()


    def closeEvent(self, event):
        # 第一个参数是父控件指针
        # 第二个参数是标题
        # 第三个参数是内容
        # 第四个参数是窗口里面要多少个按钮（默认为OK）
        # 第五个参数指定默认焦点。（默认为NoButton，此时QMessageBox会自动选择合适的默认值。）
        reply = QMessageBox.question(
            self, 'Message', "确认退出吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if reply == QMessageBox.Yes:
            event.accept()
        else:
            event.ignore()

    def init_statusbar(self):
        self.statusBar().showMessage('Ready')

    def init_menubar(self):
        # 创建菜单，并创建快捷键，以及菜单hover提示
        exit_action = QAction(QIcon('static/img/logo192.png'), '退出', self)
        exit_action.setShortcut('Ctrl+Q')
        exit_action.setStatusTip('退出应用')
        exit_action.triggered.connect(qApp.quit)  # 点击菜单中止应用程序

    def init_layout(self):
        self.layout_hbox = QHBoxLayout()

        self.guide_vbox = QVBoxLayout()
        self.mail_vbox = QVBoxLayout()
        self.content_vbox = QVBoxLayout()
        self.terminal_vbox = QVBoxLayout()

        self.guide_widget = QWidget()
        self.mail_widget = QWidget()
        self.content_widget = QWidget()
        self.terminal_widget = QWidget()

        self.guide_widget.setLayout(self.guide_vbox)
        self.mail_widget.setLayout(self.mail_vbox)
        self.content_widget.setLayout(self.content_vbox)
        self.terminal_widget.setLayout(self.terminal_vbox)

        self.layout_hbox.addWidget(self.guide_widget)
        self.layout_hbox.addWidget(self.mail_widget)
        self.layout_hbox.addWidget(self.content_widget)
        self.layout_hbox.addWidget(self.terminal_widget)

        main_widget = QWidget()
        main_widget.setLayout(self.layout_hbox)
        self.setCentralWidget(main_widget)

        self.guide_vbox.addWidget(QLabel('账户'))
        self.guide_vbox.addStretch()

        self.mail_vbox.addWidget(QLabel('收件箱'))
        self.mail_vbox.addStretch()

        self.content_vbox.addWidget(QLabel('详情'))
        self.content_detail = QWebEngineView()
        self.content_detail.setHtml('12312312313<strong>23422</strong>')
        self.content_detail.setMinimumHeight(880)

        self.content_detail.setMaximumWidth(600)
        self.content_vbox.addWidget(self.content_detail)
        self.content_vbox.addStretch()

        self.terminal_vbox.addWidget(QLabel('处理记录'))
        self.process_log = QTextEdit('')
        self.terminal_vbox.addWidget(self.process_log)
        self.terminal_widget.setMaximumWidth(200)

    # 控制窗口显示在屏幕中心
    def center(self):
        # 获得窗口
        qr = self.frameGeometry()
        # 获得屏幕中心点
        cp = QDesktopWidget().availableGeometry().center()
        # 显示到屏幕中心
        qr.moveCenter(cp)
        self.move(qr.topLeft())


if __name__ == "__main__":
    app = QApplication(sys.argv)
    loginWindow = LoginWindow()
    mainWindow = MainWindow()
    loginWindow.show()
    # mainWindow.show()
    sys.exit(app.exec_())