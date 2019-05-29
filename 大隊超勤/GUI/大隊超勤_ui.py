# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file '大隊超勤.ui'
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
        MainWindow.resize(800, 496)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(30, 80, 79, 16))
        self.label.setObjectName("label")
        self.dirpath_btn = QtWidgets.QPushButton(self.centralwidget)
        self.dirpath_btn.setGeometry(QtCore.QRect(690, 80, 93, 28))
        self.dirpath_btn.setObjectName("dirpath_btn")
        self.pmfile_cb = QtWidgets.QComboBox(self.centralwidget)
        self.pmfile_cb.setGeometry(QtCore.QRect(260, 210, 211, 31))
        self.pmfile_cb.setObjectName("pmfile_cb")
        self.namecorrection_cb = QtWidgets.QComboBox(self.centralwidget)
        self.namecorrection_cb.setGeometry(QtCore.QRect(260, 330, 211, 31))
        self.namecorrection_cb.setObjectName("namecorrection_cb")
        self.start_btn = QtWidgets.QPushButton(self.centralwidget)
        self.start_btn.setGeometry(QtCore.QRect(660, 410, 93, 28))
        self.start_btn.setObjectName("start_btn")
        self.label_6 = QtWidgets.QLabel(self.centralwidget)
        self.label_6.setGeometry(QtCore.QRect(130, 220, 91, 16))
        self.label_6.setObjectName("label_6")
        self.label_8 = QtWidgets.QLabel(self.centralwidget)
        self.label_8.setGeometry(QtCore.QRect(130, 340, 111, 20))
        self.label_8.setObjectName("label_8")
        self.namecorrection_btn = QtWidgets.QPushButton(self.centralwidget)
        self.namecorrection_btn.setGeometry(QtCore.QRect(690, 280, 93, 28))
        self.namecorrection_btn.setObjectName("namecorrection_btn")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(18, 280, 81, 20))
        self.label_5.setObjectName("label_5")
        self.pmfile_btn = QtWidgets.QPushButton(self.centralwidget)
        self.pmfile_btn.setGeometry(QtCore.QRect(687, 156, 93, 28))
        self.pmfile_btn.setObjectName("pmfile_btn")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(50, 160, 49, 16))
        self.label_3.setObjectName("label_3")
        self.dirpath = QtWidgets.QTextBrowser(self.centralwidget)
        self.dirpath.setGeometry(QtCore.QRect(117, 71, 553, 41))
        self.dirpath.setObjectName("dirpath")
        self.namecorrection = QtWidgets.QTextBrowser(self.centralwidget)
        self.namecorrection.setGeometry(QtCore.QRect(120, 270, 551, 41))
        self.namecorrection.setObjectName("namecorrection")
        self.pmfile = QtWidgets.QTextBrowser(self.centralwidget)
        self.pmfile.setGeometry(QtCore.QRect(119, 151, 551, 39))
        self.pmfile.setObjectName("pmfile")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 800, 25))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        self.dirpath_btn.clicked.connect(self.dirpath_click)
        self.pmfile_btn.clicked.connect(self.pmfile_click)
        self.pmfile_cb.currentTextChanged['QString'].connect(self.pmfile_cb_click)
        self.namecorrection_btn.clicked.connect(self.namecorrection_click)
        self.namecorrection_cb.currentTextChanged['QString'].connect(self.namecorrection_cb_click)
        self.start_btn.clicked.connect(self.start_click)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "大隊超勤"))
        self.label.setText(_translate("MainWindow", "超勤檔案夾:"))
        self.dirpath_btn.setText(_translate("MainWindow", "瀏覽"))
        self.start_btn.setText(_translate("MainWindow", "開始"))
        self.label_6.setText(_translate("MainWindow", "薪資表分頁:"))
        self.label_8.setText(_translate("MainWindow", "姓名更正表分頁:"))
        self.namecorrection_btn.setText(_translate("MainWindow", "瀏覽"))
        self.label_5.setText(_translate("MainWindow", "姓名更正表:"))
        self.pmfile_btn.setText(_translate("MainWindow", "瀏覽"))
        self.label_3.setText(_translate("MainWindow", "薪資表:"))

    @pyqtSlot( )
    def dirpath_click(self):
        pass
    
    @pyqtSlot( )
    def pmfile_click(self):
        pass
    
    @pyqtSlot( )
    def pmfile_cb_click(self):
        pass
    
    @pyqtSlot( )
    def namecorrection_click(self):
        pass
    
    @pyqtSlot( )
    def namecorrection_cb_click(self):
        pass
    
    @pyqtSlot( )
    def start_click(self):
        pass
    
    