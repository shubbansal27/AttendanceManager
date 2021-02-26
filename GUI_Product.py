__author__ = 'Administrator'

from PySide2.QtGui import *
from PySide2.QtWidgets import *
from PySide2.QtCore import Slot,Qt,QSize
import sys

class GUI_Product:

    def show(self,frame):

        self.returnVal = -1

        # creating modal window
        self.window = QDialog(frame)
        self.window.setWindowTitle('Welcome')
        # window.setWindowFlags(Qt.Popup)
        self.setColor(self.window,240,240,240)
        self.window.setMinimumSize(500,300)
        self.window.setMaximumSize(500,300)

        vlayout = QVBoxLayout()
        #1 topPanel
        topPanel = QGroupBox('')
        self.setColor(topPanel,255,255,220)
        formLayout = QFormLayout()
        #1.1
        formLayout.setVerticalSpacing(20)
        lb1 = QLabel('Attendance Manager')
        lb1.setFont(QFont('Arial',9,QFont.Bold))
        lb2 = QLabel('1.0')
        lb2.setFont(QFont('Arial',9,QFont.Bold))
        lb3 = QLabel('XYZ Pvt Ltd')
        lb3.setFont(QFont('Arial',9,QFont.Bold))
        formLayout.addRow('Project Title:  ',lb1)
        formLayout.addRow('Version:  ',lb2)
        formLayout.addRow('Oragnisation:  ',lb3)
        topPanel.setLayout(formLayout)
        vlayout.addWidget(topPanel)

        #2
        bottomPanel = QGroupBox('')
        hlayout = QHBoxLayout()
        hlayout.setAlignment(Qt.AlignLeft)
        hlayout.setSpacing(40)
        #2.1
        descLayout = QHBoxLayout()
        descLayout.setAlignment(Qt.AlignTop)
        descLabel = QLabel('\nBrowse Biometric excel report from your system.\n'+
                           'Note: Please choose valid file\n\n' +
                           '-> click on import button to continue\n'
                           )
        descLabel.setFont(QFont('Arial',9,QFont.Normal))
        descLayout.addWidget(descLabel)
        hlayout.addLayout(descLayout)
        #2.2
        btnLoad = QPushButton('Import')
        btnLoad.clicked.connect(self.load)
        self.setColor(btnLoad,255,255,255)
        btnLoad.setFlat(True)
        btnLoad.setIcon(QIcon('icons/excel.png'))
        btnLoad.setIconSize(QSize(100,100))
        btnLoad.setMinimumSize(100,100)
        hlayout.addWidget(btnLoad)
        bottomPanel.setLayout(hlayout)
        vlayout.addWidget(bottomPanel)

        self.window.setLayout(vlayout)
        self.window.exec_()
        # check exit status
        if self.returnVal != 1:
            sys.exit(1)
        else:
            return 1


####################################################
    def setColor(self,widget,R,G,B):

         p = widget.palette()
         p.setColor(widget.backgroundRole(),QColor(R,G,B))
         widget.setAutoFillBackground(True)
         widget.setPalette(p)

###################################################

    @Slot()
    def load(self):

        self.returnVal = 1
        self.window.close()

    @Slot()
    def exit(self):
        # exit application
        # set returnVal to zero, put exit code after window.exec()
        self.returnVal = 0
        self.window.close()








