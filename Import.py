__author__ = 'Administrator'

from Core import XLParser
from PySide2.QtGui import *
from PySide2.QtWidgets import *
import os

class Import:

    def __init__(self,frame,dbname):
        self.frame = frame
        self.dbname = dbname

    def createNewAttendance(self,frame):

        # open save dialog
        r = QFileDialog.getOpenFileName(parent=frame,caption='New Attendance',filter='*.xlsx')
        path = r[0]

        # if saved pressed then only
        if path != '':
            if os.path.exists(self.dbname):
                os.remove(self.dbname)

            # path not null
            #successful correct biometric file selection
            # parse
            parser = XLParser(path,self.dbname,self.frame)
            parser.parse()
            return True

        else:
            return False

