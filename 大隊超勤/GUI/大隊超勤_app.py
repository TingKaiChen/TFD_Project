#!/usr/bin/env python
# coding: utf-8

# In[1]:


from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtCore import QObject, pyqtSlot
from PyQt5.QtWidgets import QApplication, QWidget, QInputDialog, QLineEdit, QFileDialog
from .大隊超勤_ui import Ui_MainWindow
import sys, os
import xlwings as xw
import configparser as cfgparser
import ctypes

# In[ ]:


class MainWindowUIClass( Ui_MainWindow ):
    def __init__( self ):
        '''Initialize the super class
        '''
        super().__init__()
        self.app = xw.App(add_book = False, visible = False)
        self.app.display_alerts = False

        self._dirpath = None
        self._pmfile = None
        self._pmfilesht = None
        self._personinfo = None
        self._personinfosht = None
        self._namecorrection = None
        self._namecorrectionsht = None

        self.cfg = cfgparser.ConfigParser()
        self.cfg.read('config.ini')
        
    def setupUi( self, MW ):
        ''' Setup the UI of the super class, and add here code
        that relates to the way we want our UI to operate.
        '''
        super().setupUi( MW )
        self.MW = MW
        self.MW.setWindowIcon(QtGui.QIcon('./GUI/icon.png'))
        # Set different PID to show the icon
        # ref: https://stackoverflow.com/questions/1551605/how-to-set-applications-taskbar-icon-in-windows-7
        myappid = 'myapp' 
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        
    def dirpath_click(self):
        parpath = self.cfg.get('DEFAULT', 'dirpath_parpath', fallback = '')
        dirName = QFileDialog.getExistingDirectory(
            None,"選擇含有超勤檔案的資料夾", parpath + '\sdf')
        dirName += '/'
        if dirName:
            self.dirpath.setText(dirName)
            self._dirpath = dirName
    
    def pmfile_click(self):
        parpath = self.cfg.get('DEFAULT', 'pmfile_parpath', fallback = '')
        fileName, _ = QFileDialog.getOpenFileName(
            None,"選擇薪資表", parpath, "All Files (*)")
        if fileName:
            self.pmfile.setText(fileName)
            self._pmfile = fileName
            self.app.books.api.Open(fileName, UpdateLinks=False)
            self.wb = self.app.books[-1]
            self.pmfile_cb.clear()
            self.pmfile_cb.addItems([sht.name for sht in self.wb.sheets])
            self._pmfilesht = self.pmfile_cb.currentText()
    
    def pmfile_cb_click(self):
        self._pmfilesht = self.pmfile_cb.currentText()
    
    def personinfo_click(self):
        parpath = self.cfg.get('DEFAULT', 'personinfo_parpath', fallback = '')
        fileName, _ = QFileDialog.getOpenFileName(
            None,"選擇人資表", parpath, "All Files (*)")
        if fileName:
            self.personinfo.setText(fileName)
            self._personinfo = fileName
            self.app.books.api.Open(fileName, UpdateLinks=False)
            self.wb = self.app.books[-1]
            self.personinfo_cb.clear()
            self.personinfo_cb.addItems([sht.name for sht in self.wb.sheets])
            self._personinfosht = self.personinfo_cb.currentText()
    
    def personinfo_cb_click(self):
        self._personinfosht = self.personinfo_cb.currentText()
    
    def namecorrection_click(self):
        parpath = self.cfg.get('DEFAULT', 'namecorrection_parpath', fallback = '')
        fileName, _ = QFileDialog.getOpenFileName(
            None,"選擇姓名更正表", parpath, "All Files (*)")
        if fileName:
            self.namecorrection.setText(fileName)
            self._namecorrection = fileName
            self.app.books.api.Open(fileName, UpdateLinks=False)
            self.wb = self.app.books[-1]
            self.namecorrection_cb.clear()
            self.namecorrection_cb.addItems([sht.name for sht in self.wb.sheets])
            self._namecorrectionsht = self.namecorrection_cb.currentText()
    
    def namecorrection_cb_click(self):
        self._namecorrectionsht = self.namecorrection_cb.currentText()
    
    def start_click(self):
        self.MW.close() 
        for wb in self.app.books:
            wb.close()
        self.app.quit()
        self.app.kill()

        if (self._dirpath and self._pmfile and self._pmfilesht and 
           self._personinfo and self._personinfosht and self._namecorrection and 
           self._namecorrectionsht):
            # Update the configuration
            self.cfg['DEFAULT']['dirpath_parpath'] = \
                os.path.abspath(os.path.join(self._dirpath, os.pardir))
            self.cfg['DEFAULT']['pmfile_parpath'] = \
                os.path.abspath(os.path.join(self._pmfile, os.pardir))
            self.cfg['DEFAULT']['personinfo_parpath'] = \
                os.path.abspath(os.path.join(self._personinfo, os.pardir))
            self.cfg['DEFAULT']['namecorrection_parpath'] = \
                os.path.abspath(os.path.join(self._namecorrection, os.pardir))
            with open('config.ini', 'w') as configfile:
                self.cfg.write(configfile)
        
    def getParam(self):
        return (self._dirpath, self._pmfile, self._pmfilesht, self._personinfo, 
                self._personinfosht, self._namecorrection, self._namecorrectionsht)
            