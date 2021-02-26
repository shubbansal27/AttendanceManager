__author__ = 'Administrator'

from PySide2.QtWidgets import *
from PySide2.QtCore import Slot,QTime,SIGNAL
import calendar
import sqlite3

class DateFrame(QFrame):

    def __init__(self,num, MainGui_Ref):
        QFrame.__init__(self)
        self.num = num
	self.MainGui_Ref = MainGui_Ref

    def mousePressEvent(self, evt):
        print "mouse pressed !!"
        self.emit(SIGNAL('clicked'),self.num)
	self.MainGui_Ref.openGuiEdit(self.num)


###################################################################################

class GUI_Edit:

    def openDialog(self,frame,dbname,date,month,year,empID,empName):

        self.dbname = dbname
        self.date = date
        self.empID = empID
        self.returnVal = 0

        #dialog
        self.dialog = QDialog(frame)
        self.dialog.setWindowTitle(str(date)+'/'+str(month)+'/'+str(year)+' | '+empName)
        self.dialog.setMinimumSize(600,420)

        vlayout = QVBoxLayout()
        # Duty On
        self.createDutyOnBox()
        vlayout.addWidget(self.dutyOnBox)

        # Duty Off
        self.createDutyOffBox()
        vlayout.addWidget(self.dutyOffBox)

        #adding Save/Cancel button
        hlayout = QHBoxLayout()
        #1
        btnSave = QPushButton('Save')
        btnSave.clicked.connect(self.saveData)
        btnSave.setMaximumSize(50,40)
        hlayout.addWidget(btnSave)
        #2
        btnCancel = QPushButton('Cancel')
        btnCancel.clicked.connect(self.cancelEdit)
        btnCancel.setMaximumSize(50,40)
        hlayout.addWidget(btnCancel)
        vlayout.addLayout(hlayout)

        self.dialog.setLayout(vlayout)

        #fill data
        self.fillInitialData();

        #show
        self.dialog.exec_()
        return self.returnVal


#################
    def fillInitialData(self):

        # opening connection
        con = sqlite3.connect(self.dbname)

        #select query
        queryIn = "select d" + str(self.date) + " from AttendanceIn where empID = '" + self.empID + "'"
        cursorIn = con.execute(queryIn)
        queryOut = "select d" + str(self.date) + " from AttendanceOut where empID = '" + self.empID + "'"
        cursorOut = con.execute(queryOut)

        dataIn = cursorIn.fetchone()[0];
        dataOut = cursorOut.fetchone()[0];

        ## disable time-edit
        self.timeEditIn.setEnabled(False)
        self.timeEditOut.setEnabled(False)

        if dataIn == None or 'NA' in dataIn :
            self.rNAIn.setChecked(True)
        elif dataIn == 'Leave':
            self.rLeaveIn.setChecked(True)
        elif dataIn == 'OD':
            self.rODIn.setChecked(True)
        elif dataIn == 'WFH':
            self.rWFHIn.setChecked(True)
        elif dataIn == 'Holiday':
            self.rHolidayIn.setChecked(True)
        else:
            self.timeEditIn.setEnabled(True)
            hour = int(dataIn.split(':')[0])
            min = int(dataIn.split(':')[1])
            self.timeEditIn.setTime(QTime(hour,min))
            self.rPresentIn.setChecked(True)

        if dataOut == None or 'NA' in dataOut:
            self.rNAOut.setChecked(True)
        elif dataOut == 'Leave':
            self.rLeaveOut.setChecked(True)
        elif dataOut == 'OD':
            self.rODOut.setChecked(True)
        elif dataOut == 'WFH':
            self.rWFHOut.setChecked(True)
        elif dataOut == 'Holiday':
            self.rHolidayOut.setChecked(True)
        else:
            self.timeEditOut.setEnabled(True)
            hour = int(dataOut.split(':')[0])
            min = int(dataOut.split(':')[1])
            self.timeEditOut.setTime(QTime(hour,min))
            self.rPresentOut.setChecked(True)


        ###get remarks: for on and off both
        queryIn = "select d" + str(self.date) + " from RemarkIn where empID = '" + self.empID + "'"
        cursorIn = con.execute(queryIn)
        queryOut = "select d" + str(self.date) + " from RemarkOut where empID = '" + self.empID + "'"
        cursorOut = con.execute(queryOut)
        remarkDataIn = cursorIn.fetchone()[0];
        remarkDataOut = cursorOut.fetchone()[0];
        self.remarkIn.setText(remarkDataIn)
        self.remarkOut.setText(remarkDataOut)

        #close connection
        con.close()



####################

    def saveData(self):

        inData = ''
        outData = ''

        #dutyOn
        if self.rPresentIn.isChecked():
            time = self.timeEditIn.time().toString()
            inData = time.split(':')[0] + ':' + time.split(':')[1]
        elif self.rLeaveIn.isChecked():
            inData = 'Leave'
        elif self.rODIn.isChecked():
            inData = 'OD'
        elif self.rWFHIn.isChecked():
            inData = 'WFH'
        elif self.rNAIn.isChecked():
            inData = 'NA'
        elif self.rHolidayIn.isChecked():
            inData = 'Holiday'


        #dutyOff
        if self.rPresentOut.isChecked():
            time = self.timeEditOut.time().toString()
            outData = time.split(':')[0] + ':' + time.split(':')[1]
        elif self.rLeaveOut.isChecked():
            outData = 'Leave'
        elif self.rODOut.isChecked():
            outData = 'OD'
        elif self.rWFHOut.isChecked():
            outData = 'WFH'
        elif self.rNAOut.isChecked():
            outData = 'NA'
        elif self.rHolidayOut.isChecked():
            outData = 'Holiday'


        # save data
        # opening connection
        con = sqlite3.connect(self.dbname)

        #update attendanceIn and out
        queryIn = "update AttendanceIn set d" + str(self.date) + " = '" +inData+"' where empID = '" + self.empID + "'"
        con.execute(queryIn)
        queryOut = "update AttendanceOut set d" + str(self.date) + " = '" +outData+"' where empID = '" + self.empID + "'"
        con.execute(queryOut)
        ##update remarksIn and out
        queryIn = "update RemarkIn set d" + str(self.date) + " = '" +self.remarkIn.toPlainText()+"' where empID = '" + self.empID + "'"
        con.execute(queryIn)
        queryOut = "update RemarkOut set d" + str(self.date) + " = '" +self.remarkOut.toPlainText()+"' where empID = '" + self.empID + "'"
        con.execute(queryOut)

        #commit
        con.commit()
        #closing connection
        con.close()

        #close dialog
        self.dialog.close()
        self.returnVal = 1

################################
    def cancelEdit(self):

        #close dialog
        self.dialog.close()
        self.returnVal = 0



#######################
    def createDutyOnBox(self):
        self.dutyOnBox = QGroupBox('Duty On')
        hlayout = QHBoxLayout()
        #left part
        vlayoutLeft = QVBoxLayout()
        self.rPresentIn = QRadioButton('Present')
        self.rPresentIn.clicked.connect(self.dutyOnRadioClicked)
        self.rLeaveIn = QRadioButton('Leave')
        self.rLeaveIn.clicked.connect(self.dutyOnRadioClicked)
        self.rODIn = QRadioButton('OD')
        self.rODIn.clicked.connect(self.dutyOnRadioClicked)
        self.rWFHIn = QRadioButton('WFH')
        self.rWFHIn.clicked.connect(self.dutyOnRadioClicked)
        self.rNAIn = QRadioButton('NA')
        self.rNAIn.clicked.connect(self.dutyOnRadioClicked)
        self.rHolidayIn = QRadioButton('Holiday')
        self.rHolidayIn.clicked.connect(self.dutyOnRadioClicked)

        vlayoutLeft.addWidget(self.rPresentIn)
        vlayoutLeft.addWidget(self.rLeaveIn)
        vlayoutLeft.addWidget(self.rODIn)
        vlayoutLeft.addWidget(self.rWFHIn)
        vlayoutLeft.addWidget(self.rNAIn)
        vlayoutLeft.addWidget(self.rHolidayIn)
        hlayout.addLayout(vlayoutLeft)

        #right part
        vlayoutRight = QVBoxLayout()
        #1
        groupInTime = QGroupBox('In Time')
        vlayout = QVBoxLayout()
        self.timeEditIn = QTimeEdit()
        self.timeEditIn.setMinimumSize(100,35)
        self.timeEditIn.setDisplayFormat('h:mm a')
        vlayout.addWidget(self.timeEditIn)
        groupInTime.setLayout(vlayout)
        vlayoutRight.addWidget(groupInTime)
        #2
        #Remark
        groupRemark = QGroupBox('Remarks')
        vlayout = QVBoxLayout()
        self.remarkIn = QTextEdit()
        vlayout.addWidget(self.remarkIn)
        groupRemark.setLayout(vlayout)
        vlayoutRight.addWidget(groupRemark)

        hlayout.addLayout(vlayoutRight)
        self.dutyOnBox.setLayout(hlayout)



    def createDutyOffBox(self):
        self.dutyOffBox = QGroupBox('Duty Off')
        hlayout = QHBoxLayout()
        #left part
        vlayoutLeft = QVBoxLayout()
        self.rPresentOut = QRadioButton('Present')
        self.rPresentOut.clicked.connect(self.dutyOffRadioClicked)
        self.rLeaveOut = QRadioButton('Leave')
        self.rLeaveOut.clicked.connect(self.dutyOffRadioClicked)
        self.rODOut = QRadioButton('OD')
        self.rODOut.clicked.connect(self.dutyOffRadioClicked)
        self.rWFHOut = QRadioButton('WFH')
        self.rWFHOut.clicked.connect(self.dutyOffRadioClicked)
        self.rNAOut = QRadioButton('NA')
        self.rNAOut.clicked.connect(self.dutyOffRadioClicked)
        self.rHolidayOut = QRadioButton('Holiday')
        self.rHolidayOut.clicked.connect(self.dutyOffRadioClicked)

        vlayoutLeft.addWidget(self.rPresentOut)
        vlayoutLeft.addWidget(self.rLeaveOut)
        vlayoutLeft.addWidget(self.rODOut)
        vlayoutLeft.addWidget(self.rWFHOut)
        vlayoutLeft.addWidget(self.rNAOut)
        vlayoutLeft.addWidget(self.rHolidayOut)
        hlayout.addLayout(vlayoutLeft)

        #right part
        vlayoutRight = QVBoxLayout()
        #1
        groupInTime = QGroupBox('Out Time')
        vlayout = QVBoxLayout()
        self.timeEditOut = QTimeEdit()
        self.timeEditOut.setMinimumSize(100,35)
        self.timeEditOut.setDisplayFormat('h:mm a')
        vlayout.addWidget(self.timeEditOut)
        groupInTime.setLayout(vlayout)
        vlayoutRight.addWidget(groupInTime)
        #2
        #Remark
        groupRemark = QGroupBox('Remarks')
        vlayout = QVBoxLayout()
        self.remarkOut = QTextEdit()
        vlayout.addWidget(self.remarkOut)
        groupRemark.setLayout(vlayout)
        vlayoutRight.addWidget(groupRemark)

        hlayout.addLayout(vlayoutRight)
        self.dutyOffBox.setLayout(hlayout)


#############################

    @Slot()
    def dutyOnRadioClicked(self):

        ##if present is checked, enable timeEdit
        if self.rPresentIn.isChecked():
            self.timeEditIn.setEnabled(True)
        else:
            self.timeEditIn.setEnabled(False)

        #clear remarks
        self.remarkIn.setText('')


#############################
    @Slot()
    def dutyOffRadioClicked(self):

        ##if present is checked, enable timeEdit
        if self.rPresentOut.isChecked():
            self.timeEditOut.setEnabled(True)
        else:
            self.timeEditOut.setEnabled(False)


        #clear remarks
        self.remarkOut.setText('')
