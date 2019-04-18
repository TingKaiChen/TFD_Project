# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '三級加班.ui'
#
# Created by: PyQt5 UI code generator 5.11.3
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from PyQt5.QtWidgets import QWidget

class Ui_MainWindow(QWidget):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(800, 396)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 80, 151, 16))
        self.label.setObjectName("label")
        self.filename_btn = QtWidgets.QPushButton(self.centralwidget)
        self.filename_btn.setGeometry(QtCore.QRect(690, 80, 93, 28))
        self.filename_btn.setObjectName("filename_btn")
        self.start_btn = QtWidgets.QPushButton(self.centralwidget)
        self.start_btn.setGeometry(QtCore.QRect(660, 300, 93, 28))
        self.start_btn.setObjectName("start_btn")
        self.filename = QtWidgets.QTextBrowser(self.centralwidget)
        self.filename.setGeometry(QtCore.QRect(179, 71, 491, 41))
        self.filename.setObjectName("filename")
        self.filesht_cb = QtWidgets.QComboBox(self.centralwidget)
        self.filesht_cb.setGeometry(QtCore.QRect(223, 130, 131, 21))
        self.filesht_cb.setObjectName("filesht_cb")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(130, 130, 79, 16))
        self.label_2.setObjectName("label_2")
        self.splitter = QtWidgets.QSplitter(self.centralwidget)
        self.splitter.setGeometry(QtCore.QRect(150, 180, 220, 21))
        self.splitter.setOrientation(QtCore.Qt.Horizontal)
        self.splitter.setObjectName("splitter")
        self.label_6 = QtWidgets.QLabel(self.splitter)
        self.label_6.setObjectName("label_6")
        self.year_box = QtWidgets.QComboBox(self.splitter)
        self.year_box.setObjectName("year_box")
        self.label_7 = QtWidgets.QLabel(self.splitter)
        self.label_7.setObjectName("label_7")
        self.month_box = QtWidgets.QComboBox(self.splitter)
        self.month_box.setObjectName("month_box")
        self.label_8 = QtWidgets.QLabel(self.splitter)
        self.label_8.setObjectName("label_8")
        self.splitter_2 = QtWidgets.QSplitter(self.centralwidget)
        self.splitter_2.setGeometry(QtCore.QRect(80, 240, 301, 21))
        self.splitter_2.setOrientation(QtCore.Qt.Horizontal)
        self.splitter_2.setObjectName("splitter_2")
        self.label_9 = QtWidgets.QLabel(self.splitter_2)
        self.label_9.setObjectName("label_9")
        self.account = QtWidgets.QLineEdit(self.splitter_2)
        self.account.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.account.setObjectName("account")
        self.splitter_3 = QtWidgets.QSplitter(self.centralwidget)
        self.splitter_3.setGeometry(QtCore.QRect(80, 290, 301, 21))
        self.splitter_3.setOrientation(QtCore.Qt.Horizontal)
        self.splitter_3.setObjectName("splitter_3")
        self.label_10 = QtWidgets.QLabel(self.splitter_3)
        self.label_10.setObjectName("label_10")
        self.password = QtWidgets.QLineEdit(self.splitter_3)
        self.password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.password.setObjectName("password")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 25))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.start_btn.clicked.connect(self.start_click)
        self.filename_btn.clicked.connect(self.filename_click)
        self.year_box.currentTextChanged['QString'].connect(self.year_click)
        self.month_box.currentTextChanged['QString'].connect(self.month_click)
        self.filesht_cb.currentTextChanged['QString'].connect(self.filesht_cb_click)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "三級開設加班"))
        self.label.setText(_translate("MainWindow", "三級專案加班統計表:"))
        self.filename_btn.setText(_translate("MainWindow", "瀏覽"))
        self.start_btn.setText(_translate("MainWindow", "開始"))
        self.label_2.setText(_translate("MainWindow", "統計表分頁:"))
        self.label_6.setText(_translate("MainWindow", "西元"))
        self.label_7.setText(_translate("MainWindow", "年"))
        self.label_8.setText(_translate("MainWindow", "月"))
        self.label_9.setText(_translate("MainWindow", "WebITR 差勤系統帳號:"))
        self.label_10.setText(_translate("MainWindow", "WebITR 差勤系統密碼:"))

    @pyqtSlot( )
    def start_click(self):
        pass
    
    @pyqtSlot( )
    def filename_click(self):
        pass
    
    @pyqtSlot( )
    def year_click(self):
        pass
    
    @pyqtSlot( )
    def month_click(self):
        pass
    
    @pyqtSlot( )
    def filesht_cb_click(self):
        pass
    
    