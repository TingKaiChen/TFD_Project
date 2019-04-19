# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'C:\Users\TFD\救指調假\GUI\救指調假.ui'
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
        MainWindow.resize(800, 552)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.start_btn = QtWidgets.QPushButton(self.centralwidget)
        self.start_btn.setGeometry(QtCore.QRect(670, 460, 93, 28))
        self.start_btn.setObjectName("start_btn")
        self.splitter_2 = QtWidgets.QSplitter(self.centralwidget)
        self.splitter_2.setGeometry(QtCore.QRect(420, 30, 301, 21))
        self.splitter_2.setOrientation(QtCore.Qt.Horizontal)
        self.splitter_2.setObjectName("splitter_2")
        self.label_9 = QtWidgets.QLabel(self.splitter_2)
        self.label_9.setObjectName("label_9")
        self.account = QtWidgets.QLineEdit(self.splitter_2)
        self.account.setEchoMode(QtWidgets.QLineEdit.Normal)
        self.account.setObjectName("account")
        self.splitter_3 = QtWidgets.QSplitter(self.centralwidget)
        self.splitter_3.setGeometry(QtCore.QRect(420, 80, 301, 21))
        self.splitter_3.setOrientation(QtCore.Qt.Horizontal)
        self.splitter_3.setObjectName("splitter_3")
        self.label_10 = QtWidgets.QLabel(self.splitter_3)
        self.label_10.setObjectName("label_10")
        self.password = QtWidgets.QLineEdit(self.splitter_3)
        self.password.setEchoMode(QtWidgets.QLineEdit.Password)
        self.password.setObjectName("password")
        self.start_dt_cal = QtWidgets.QCalendarWidget(self.centralwidget)
        self.start_dt_cal.setGeometry(QtCore.QRect(50, 210, 301, 231))
        self.start_dt_cal.setGridVisible(True)
        self.start_dt_cal.setObjectName("start_dt_cal")
        self.end_dt_cal = QtWidgets.QCalendarWidget(self.centralwidget)
        self.end_dt_cal.setGeometry(QtCore.QRect(400, 210, 301, 231))
        self.end_dt_cal.setGridVisible(True)
        self.end_dt_cal.setObjectName("end_dt_cal")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 40, 121, 21))
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(60, 180, 71, 21))
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(410, 180, 71, 21))
        self.label_3.setObjectName("label_3")
        self.start_dt_label = QtWidgets.QLabel(self.centralwidget)
        self.start_dt_label.setGeometry(QtCore.QRect(130, 180, 121, 21))
        self.start_dt_label.setText("")
        self.start_dt_label.setObjectName("start_dt_label")
        self.end_dt_label = QtWidgets.QLabel(self.centralwidget)
        self.end_dt_label.setGeometry(QtCore.QRect(480, 180, 121, 21))
        self.end_dt_label.setText("")
        self.end_dt_label.setObjectName("end_dt_label")
        self.year_box = QtWidgets.QComboBox(self.centralwidget)
        self.year_box.setGeometry(QtCore.QRect(30, 80, 80, 21))
        self.year_box.setObjectName("year_box")
        self.label_7 = QtWidgets.QLabel(self.centralwidget)
        self.label_7.setGeometry(QtCore.QRect(116, 80, 16, 21))
        self.label_7.setObjectName("label_7")
        self.month_box = QtWidgets.QComboBox(self.centralwidget)
        self.month_box.setGeometry(QtCore.QRect(137, 80, 80, 21))
        self.month_box.setObjectName("month_box")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(223, 80, 141, 21))
        self.label_8.setObjectName("label_8")
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
        self.year_box.currentTextChanged['QString'].connect(self.year_click)
        self.month_box.currentTextChanged['QString'].connect(self.month_click)
        self.start_dt_cal.selectionChanged.connect(self.start_dt)
        self.end_dt_cal.selectionChanged.connect(self.end_dt)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "救指調假"))
        self.start_btn.setText(_translate("MainWindow", "開始"))
        self.label_9.setText(_translate("MainWindow", "WebITR 差勤系統帳號:"))
        self.label_10.setText(_translate("MainWindow", "WebITR 差勤系統密碼:"))
        self.label.setText(_translate("MainWindow", "文件輸出標籤:"))
        self.label_2.setText(_translate("MainWindow", "開始日期:"))
        self.label_3.setText(_translate("MainWindow", "結束日期:"))
        self.label_7.setText(_translate("MainWindow", "年"))
        self.label_8.setText(_translate("MainWindow", "月救指調假資訊.txt"))

    @pyqtSlot( )
    def start_click(self):
        pass
    
    @pyqtSlot( )
    def year_click(self):
        pass
    
    @pyqtSlot( )
    def month_click(self):
        pass
    
    @pyqtSlot( )
    def start_dt(self):
        pass

    @pyqtSlot( )
    def end_dt(self):
        pass
    
    