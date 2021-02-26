__author__ = 'Administrator'

import sys,os
import calendar
import sqlite3
from PySide2.QtWidgets import *
from PySide2.QtGui import *
from PySide2.QtCore import Slot,SIGNAL,Qt, QObject
from GUI_Holiday import GUI_Holiday
from GUI_Edit import GUI_Edit
from GUI_Edit import DateFrame
from GUI_Info import GUI_Info
from Export import Export
from Import import Import
from GUI_Product import GUI_Product

class MainGui:

    def __init__(self,dbname):
        self.dbname = dbname


    def open(self):

        # check 'localdb.db' exists
        # set flagDB
        if not os.path.exists(self.dbname):
            self.flagDB = False
        else:
            self.flagDB = True


        # create app
        app = QApplication(sys.argv)

        # create frame object
        self.frame = QWidget()
        self.frame.setWindowTitle('Attendance Manager')
        self.frame.setWindowIcon(QIcon('icons/icon.ico'))
        self.frame.setMinimumSize(1250,650)

        #create mainLayout
        mainLayout = QVBoxLayout()
        self.frame.setLayout(mainLayout)

        # create ButtonGroupe
        self.createButtonPanel()
        mainLayout.addWidget(self.controls)

        #create center Panel
        #center panel -> left panel, empListPanel
        self.createCenterPanel()
        mainLayout.addWidget(self.centerPanel)

        ###load data from localdb
        ## load employee data after creating buttonGroup and centerPanel
        ## this function will load class variables like year,month,orgName that are needed in bottomPanel:calender
        if self.flagDB:
            self.loadEmployeeList()

        #create bottom panel
        #infoPanel: info,calender,timeline details of selected employee
        self.createBottomPanel()
        mainLayout.addWidget(self.tabbedPane)

        # execute app
        self.frame.showMaximized()

        # check dbname is Empty i.e first time use
        # implementing first time use behaviour
        # display product details and give option to choose biometric data
        if not self.flagDB:

            #set processing start
            self.lbProcess.setHidden(False)

            guiPro = GUI_Product()
            loopVar = False
            while not loopVar:
                t = guiPro.show(self.frame)
                if t ==1:
                    loopVar = Import(self.frame,self.dbname).createNewAttendance(self.frame)

            #here, excel successfully parsed, restart application
            #create remaining gui
            self.flagDB = True
            self.loadEmployeeList()
            self.createCalenderTab(self.tabCalender)
            self.createInfoTab(self.tabInfo)
            self.lbProcess.setHidden(True)

        app.exec_()




#################################################################3

    def createBottomPanel(self):

         ###tabbedPane
        self.tabbedPane = QTabWidget()

        #tab2:calender
        self.tabCalender = QWidget()
        self.createCalenderTab(self.tabCalender)
        self.tabbedPane.addTab(self.tabCalender,'Calender')

        #tab1: Info
        self.tabInfo = QWidget()
        self.createInfoTab(self.tabInfo)
        self.tabbedPane.addTab(self.tabInfo,'Info')

        # #tab3:timeline
        # not implemented in this version of project
        # tab3 = QWidget()
        # self.tabbedPane.addTab(tab3,'Timeline')




    def createInfoTab(self,tab):

        if self.flagDB:
            self.guiInfo = GUI_Info(self.dbname,self.month,self.year)
            infoLayout = self.guiInfo.createInfoPanel()
            tab.setLayout(infoLayout)



    def createCalenderTab(self,tab):

        if self.flagDB:
            #1 - leftPopup
            hlayout = QHBoxLayout()
            hlayout.setContentsMargins(0,0,0,0)
            self.leftPopupPanel = QWidget()
            self.leftPopupPanel.setMaximumWidth(30)
            self.setColor(self.leftPopupPanel,250,250,250)
            #addButtons
            #button-1:
            vlayout = QVBoxLayout()
            vlayout.setAlignment(Qt.AlignTop)
            vlayout.setContentsMargins(0,15,0,0)
            self.btnFilter = QPushButton('')
            self.btnFilter.setFlat(True)
            self.setColor(self.btnFilter,200,220,220)
            self.btnFilter.setIcon(QIcon('icons/filter.png'))
            self.btnFilter.clicked.connect(self.setFilter)
            vlayout.addWidget(self.btnFilter)
            #button-2:
            self.btnClearFilter = QPushButton('')
            self.setColor(self.btnClearFilter,190,250,190)
            self.btnClearFilter.setFlat(True)
            self.btnClearFilter.setIcon(QIcon('icons/clear.png'))
            self.btnClearFilter.clicked.connect(self.clearFilter)
            vlayout.addWidget(self.btnClearFilter)
            self.leftPopupPanel.setLayout(vlayout)
            hlayout.addWidget(self.leftPopupPanel)
            #2 popup
            calendarPanel = QWidget()
            #check dbname is not empty
            #createCalendar method access local.db, so before calling ensure its existence
            self.createCalender(calendarPanel)
            hlayout.addWidget(calendarPanel)
            tab.setLayout(hlayout)



    def createCalender(self,panel):

        # setting popup to none
        self.popup = None

        # filters boolean global variables
        self.boolNA = True
        self.boolOD = False
        self.boolLeave = False
        self.boolWFH = False
        self.boolHoliday = False
        self.boolExtra = False
        self.boolLate = False

        # filter colors
        self.colorNA = [150,240,240]
        self.colorOD = [200,50,50]
        self.colorLeave = [255,243,13]
        self.colorWFH = [57,217,4]
        self.colorHoliday = [231,53,244]
        self.colorExtra = [155,155,0]
        self.colorLate = [128,64,185]

        #days number to  name mapping
        days = calendar.weekheader(3).split(' ')
        dictDays = dict()
        i = 0;
        for item in days:
            dictDays[i] = item
            i = i+1

        ### GUI Component ###
        #calender list
        self.calenderList = list()
        self.calenderListOut = list()
        self.holidayIconList = list()

        # font
        dateFont = QFont('Arial',10,QFont.Bold)

        vlayout = QVBoxLayout()
        #row1
        hlayout1 = QHBoxLayout()
        for i in range(1,11,1):
            box = DateFrame(i, self)
            box.connect(SIGNAL('clicked'),self.openGuiEdit)
            #color even
            if i % 2 == 0:
                self.setColor(box,240,240,240)

            box.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
            boxLayout = QVBoxLayout()
            #1
            hlayoutDate = QVBoxLayout()
            #1.1
            hlayoutIcon = QHBoxLayout()
            hlayoutIcon.setAlignment(Qt.AlignLeft)
            hlayoutIcon.setSpacing(0)
            #1.1.1
            date = QLabel('%d'%i)
            date.setFont(dateFont)
            hlayoutIcon.addWidget(date)
            #1.1.2
            btnHoliday = QPushButton('')
            btnHoliday.setIcon(QIcon('icons/bell.png'))
            btnHoliday.setFlat(True)
            btnHoliday.setHidden(True)
            hlayoutIcon.addWidget(btnHoliday)
            self.holidayIconList.append(btnHoliday)    #add button reference in list
            hlayoutDate.addLayout(hlayoutIcon)
            #1.2
            strDay = QLabel('%s'%dictDays[calendar.weekday(self.year,self.month,i)])
            hlayoutDate.addWidget(strDay)
            boxLayout.addLayout(hlayoutDate)
            #2
            hlayoutData = QHBoxLayout()
            #2.1
            data1 = QLabel('')
            data1.setAlignment(Qt.AlignCenter)
            self.calenderList.append(data1)
            hlayoutData.addWidget(data1)
            #2.2
            dataSep = QLabel('-')
            dataSep.setAlignment(Qt.AlignCenter)
            hlayoutData.addWidget(dataSep)
            #2.3
            data2 = QLabel('')
            data2.setAlignment(Qt.AlignCenter)
            self.calenderListOut.append(data2)
            hlayoutData.addWidget(data2)
            boxLayout.addLayout(hlayoutData)
            box.setLayout(boxLayout)
            hlayout1.addWidget(box)

        vlayout.addLayout(hlayout1)

        #row2
        hlayout2 = QHBoxLayout()
        for i in range(11,21,1):
            box = DateFrame(i, self)
            box.connect(SIGNAL('clicked'),self.openGuiEdit)
            #QObject.connect(box,SIGNAL('clicked()'),self.openGuiEdit)
            #color even
            if i % 2 == 0:
                self.setColor(box,240,240,240)

            box.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
            boxLayout = QVBoxLayout()
            #1
            hlayoutDate = QVBoxLayout()
            #1.1
            hlayoutIcon = QHBoxLayout()
            hlayoutIcon.setAlignment(Qt.AlignLeft)
            hlayoutIcon.setSpacing(0)
            #1.1.1
            date = QLabel('%d'%i)
            date.setFont(dateFont)
            hlayoutIcon.addWidget(date)
            #1.1.2
            btnHoliday = QPushButton('')
            btnHoliday.setIcon(QIcon('icons/bell.png'))
            btnHoliday.setFlat(True)
            btnHoliday.setHidden(True)
            hlayoutIcon.addWidget(btnHoliday)
            self.holidayIconList.append(btnHoliday)    #add button reference in list
            hlayoutDate.addLayout(hlayoutIcon)
            #1.2
            strDay = QLabel('%s'%dictDays[calendar.weekday(self.year,self.month,i)])
            hlayoutDate.addWidget(strDay)
            boxLayout.addLayout(hlayoutDate)
            #2
            hlayoutData = QHBoxLayout()
            #2.1
            data1 = QLabel('')
            data1.setAlignment(Qt.AlignCenter)
            self.calenderList.append(data1)
            hlayoutData.addWidget(data1)
            #2.2
            dataSep = QLabel('-')
            dataSep.setAlignment(Qt.AlignCenter)
            hlayoutData.addWidget(dataSep)
            #2.3
            data2 = QLabel('')
            data2.setAlignment(Qt.AlignCenter)
            self.calenderListOut.append(data2)
            hlayoutData.addWidget(data2)
            boxLayout.addLayout(hlayoutData)
            box.setLayout(boxLayout)
            hlayout2.addWidget(box)

        vlayout.addLayout(hlayout2)


        #row3
        hlayout3 = QHBoxLayout()
        for i in range(21,calendar.monthrange(self.year,self.month)[1]+1,1):
            box = DateFrame(i, self)
            box.connect(SIGNAL('clicked'),self.openGuiEdit)
            #QObject.connect(box,SIGNAL('clicked()'),self.openGuiEdit)
            #color even
            if i % 2 == 0:
                self.setColor(box,240,240,240)

            box.setFrameStyle(QFrame.StyledPanel | QFrame.Sunken)
            boxLayout = QVBoxLayout()
           #1
            hlayoutDate = QVBoxLayout()
            #1.1
            hlayoutIcon = QHBoxLayout()
            hlayoutIcon.setAlignment(Qt.AlignLeft)
            hlayoutIcon.setSpacing(0)
            #1.1.1
            date = QLabel('%d'%i)
            date.setFont(dateFont)
            hlayoutIcon.addWidget(date)
            #1.1.2
            btnHoliday = QPushButton('')
            btnHoliday.setIcon(QIcon('icons/bell.png'))
            btnHoliday.setFlat(True)
            btnHoliday.setHidden(True)
            hlayoutIcon.addWidget(btnHoliday)
            self.holidayIconList.append(btnHoliday)    #add button reference in list
            hlayoutDate.addLayout(hlayoutIcon)
            #1.2
            strDay = QLabel('%s'%dictDays[calendar.weekday(self.year,self.month,i)])
            hlayoutDate.addWidget(strDay)
            boxLayout.addLayout(hlayoutDate)
            #2
            hlayoutData = QHBoxLayout()
            #2.1
            data1 = QLabel('')
            data1.setAlignment(Qt.AlignCenter)
            self.calenderList.append(data1)
            hlayoutData.addWidget(data1)
            #2.2
            dataSep = QLabel('-')
            dataSep.setAlignment(Qt.AlignCenter)
            hlayoutData.addWidget(dataSep)
            #2.3
            data2 = QLabel('')
            data2.setAlignment(Qt.AlignCenter)
            self.calenderListOut.append(data2)
            hlayoutData.addWidget(data2)
            boxLayout.addLayout(hlayoutData)
            box.setLayout(boxLayout)
            hlayout3.addWidget(box)

        vlayout.addLayout(hlayout3)
        panel.setLayout(vlayout)


    def createCenterPanel(self):

        self.centerPanel = QWidget()
        gridLayout = QGridLayout()

        #rightPanel: empListPanel
        #table
        self.tableWidget = QTableWidget(0, 2)
        #self.tableWidget.setEditTriggers(False)
        self.tableWidget.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        colHeader1 = QTableWidgetItem('Employee ID')
        colHeader2 = QTableWidgetItem('Name')
        self.tableWidget.setColumnWidth(0,250)
        self.tableWidget.setColumnWidth(1,680)
        self.tableWidget.setHorizontalHeaderItem(0,colHeader1)
        self.tableWidget.setHorizontalHeaderItem(1,colHeader2)
        self.tableWidget.itemSelectionChanged.connect(self.employeeSelectionChanged)
        gridLayout.addWidget(self.tableWidget,0,1)

        #leftPanel -> topPanel,bottomPanel
        vlayout = QVBoxLayout()
        #topPanel
        detailPanel = QFrame()
        self.setColor(detailPanel,240,220,200)
        detailPanel.setMinimumWidth(360)
        detailPanel.setMinimumHeight(180)
        detailPanel.setMaximumWidth(360)
        detailPanel.setMaximumHeight(180)
        formLayout = QFormLayout()
        formLayout.setVerticalSpacing(15)
        self.orgLabel = QLabel('NA')
        self.monthLabel = QLabel('NA')
        self.monthLabel.setFont(QFont('Arial',9,QFont.Bold))
        self.totalEmpLabel = QLabel('NA')
        self.totalEmpLabel.setFont(QFont('Arial',9,QFont.Bold))
        formLayout.addRow('Oragnisation:  ',self.orgLabel)
        formLayout.addRow('Month:  ',self.monthLabel)
        formLayout.addRow('Total Employee:  ',self.totalEmpLabel)
        detailPanel.setLayout(formLayout)
        vlayout.addWidget(detailPanel)

        ## bottomPanel
        selectionDisplayPanel = QFrame()
        self.setColor(selectionDisplayPanel,240,220,200)
        selectionDisplayPanel.setMinimumWidth(290)
        selectionDisplayPanel.setMinimumHeight(100)
        selectionDisplayPanel.setMaximumWidth(290)
        selectionDisplayPanel.setMaximumHeight(100)
        formSelectedDisplay = QFormLayout()
        formSelectedDisplay.setVerticalSpacing(10)
        self.labelSelectedEmpID = QLabel('Not Selected')
        self.labelSelectedEmpID.setFont(QFont('Arial',9,QFont.Bold))
        self.labelSelectedEmpName = QLabel('Not Selected')
        self.labelSelectedEmpName.setFont(QFont('Arial',9,QFont.Bold))
        # formSelectedDisplay.addRow('',QLabel(''))   #empty row
        formSelectedDisplay.addRow('EmpID:   ',self.labelSelectedEmpID)
        formSelectedDisplay.addRow('EmpName: ',self.labelSelectedEmpName)
        selectionDisplayPanel.setLayout(formSelectedDisplay)
        vlayout.addWidget(selectionDisplayPanel)

        gridLayout.addLayout(vlayout,0,0)

        self.centerPanel.setLayout(gridLayout)


    def createButtonPanel(self):

        self.controls = QFrame()
        self.setColor(self.controls,150,150,150)
        hl = QHBoxLayout()

        #controls-1
        buttonContainer1 = QGroupBox('')
        buttonContainer1.setMaximumWidth(600)
        hlayout = QHBoxLayout()
        #button-1
        button1 = QPushButton('Set Holidays')
        button1.setMinimumHeight(40)
        button1.clicked.connect(self.openGuiHolidays)
        hlayout.addWidget(button1)
        #button-2
        button2 = QPushButton('Edit Policy')
        button2.setEnabled(False)
        button2.setMinimumHeight(40)
        hlayout.addWidget(button2)
        #button-3
        button3 = QPushButton('Export')
        button3.setMinimumHeight(40)
        button3.clicked.connect(self.exportData)
        hlayout.addWidget(button3)
        buttonContainer1.setLayout(hlayout)
        hl.addWidget(buttonContainer1)

        #controls-2
        buttonContainer2 = QGroupBox('')
        buttonContainer2.setMaximumWidth(600)
        hlayout = QHBoxLayout()
        #button-1
        button1 = QPushButton('New Attendance')
        button1.setMinimumHeight(40)
        button1.clicked.connect(self.ImportData)
        hlayout.addWidget(button1)
        #button-2
        button2 = QPushButton('Exit')
        button2.setMinimumHeight(40)
        button2.clicked.connect(self.exit)
        hlayout.addWidget(button2)
        buttonContainer2.setLayout(hlayout)
        hl.addWidget(buttonContainer2)

        #controls-3
        buttonContainer2 = QGroupBox('')
        buttonContainer2.setMaximumWidth(600)
        hlayout = QHBoxLayout()
        hlayout.setAlignment(Qt.AlignRight)
        #1
        self.lbProcess = QLabel('Processing... Please wait...')
        self.lbProcess.setStyleSheet('color:white')
        self.lbProcess.setHidden(True)
        hlayout.addWidget(self.lbProcess)
        buttonContainer2.setLayout(hlayout)
        hl.addWidget(buttonContainer2)
        self.controls.setLayout(hl)



####################################################3
    def setColor(self,widget,R,G,B):

         p = widget.palette()
         p.setColor(widget.backgroundRole(),QColor(R,G,B))
         widget.setAutoFillBackground(True)
         widget.setPalette(p)



    def loadEmployeeList(self):

         # opening connection
        con = sqlite3.connect(self.dbname)

        #select query
        #get Employee List
        query = 'select empID,empName from AttendanceIn'
        cursor = con.execute(query)
        r = 0
        for row in cursor:
            self.tableWidget.insertRow(r)
            self.tableWidget.setItem(r,0,QTableWidgetItem(row[0]))
            self.tableWidget.setItem(r,1,QTableWidgetItem(row[1]))
            r = r + 1


        # set basic details like month,year,orgName
        query = 'select orgName,year,month from BasicDetails'
        cursor = con.execute(query)
        data = cursor.fetchone()
        self.oragnisation = data[0]
        self.year = data[1]
        self.month = data[2]
        #set to gui
        self.orgLabel.setText(self.oragnisation)

        #monthNumbertoName
        mapMonth = ['January','February','March','April','May','June','July','August','September','October','November','December']
        self.monthLabel.setText(mapMonth[self.month-1] + ' ' + str(self.year))
        self.totalEmpLabel.setText(str(r))

        #closing connection
        con.close()


##########################################

    def reloadCalenderData(self):

        # opening connection
        con = sqlite3.connect(self.dbname)

        ##getting holidays list
        query = 'select date from Holidays'
        holidayList = list()
        cursor = con.execute(query)
        for row in cursor:
            date = int(row[0].split('/')[1])
            holidayList.append(date)


        ##employee list
        empList = list()
        query = 'select empID from AttendanceIn'
        cursor = con.execute(query)
        for row in cursor:
            empList.append(row[0])


        # iterate over each empID
        for empID in empList:
            #select query
            queryIn = "select * from AttendanceIn where empID = '" + empID + "'"
            cursorIn = con.execute(queryIn)
            queryOut = "select * from AttendanceOut where empID = '" + empID + "'"
            cursorOut = con.execute(queryOut)

            dataIn = cursorIn.fetchone()
            dataOut = cursorOut.fetchone()

            count = 0
            for dateItem in self.calenderList:
                dutyOn = dataIn[count + 2]
                dutyOff = dataOut[count + 2]

                if dutyOn == None:
                    dutyOn = 'NA'
                if dutyOff == None:
                    dutyOff = 'NA'

                ##mark holiday
                if 'NA' in dutyOn and (count+1) in holidayList:
                    #update attendanceIn
                    subquery = "update AttendanceIn set d" + str(count+1) + " = '" + 'Holiday' +"' where empID = '" + empID + "'"
                    con.execute(subquery)


                if 'NA' in dutyOff and (count+1) in holidayList:
                    #update attendanceIn
                    subquery = "update AttendanceOut set d" + str(count+1) + " = '" + 'Holiday' +"' where empID = '" + empID + "'"
                    con.execute(subquery)


                ##reset to NA mark
                ##In case of any holiday removed from the controls
                if 'Holiday' in dutyOn and (count+1) not in holidayList:
                    #update attendanceIn
                    subquery = "update AttendanceIn set d" + str(count+1) + " = '" + 'NA' +"' where empID = '" + empID + "'"
                    con.execute(subquery)


                if 'Holiday' in dutyOff and (count+1) not in holidayList:
                    #update attendanceIn
                    subquery = "update AttendanceOut set d" + str(count+1) + " = '" + 'NA' +"' where empID = '" + empID + "'"
                    con.execute(subquery)


                count = count + 1

        #closing connection
        con.commit()
        con.close()


###############################################


    def loadCalenderData(self,empID):

        if 'Not Selected' not in empID:
             # opening connection
            con = sqlite3.connect(self.dbname)

            ##getting holidays list
            query = 'select date from Holidays'
            holidayList = list()
            cursor = con.execute(query)
            for row in cursor:
                date = int(row[0].split('/')[1])
                holidayList.append(date)


            #select query
            queryIn = "select * from AttendanceIn where empID = '" + empID + "'"
            cursorIn = con.execute(queryIn)
            queryOut = "select * from AttendanceOut where empID = '" + empID + "'"
            cursorOut = con.execute(queryOut)

            dataIn = cursorIn.fetchone()
            dataOut = cursorOut.fetchone()

            count = 0
            for dateItem in self.calenderList:

                # set/unset official holiday icon
                refBtnHoliday = self.holidayIconList[count]
                if (count + 1) in holidayList:
                    refBtnHoliday.setHidden(False)
                else:
                    refBtnHoliday.setHidden(True)


                #dutyIn QLabel already in dateItem
                #for dutyOff need to create one more: dateItemOut
                dateItemOut = self.calenderListOut[count]

                dutyOn = dataIn[count + 2]
                dutyOff = dataOut[count + 2]

                if dutyOn == None:
                    dutyOn = 'NA'
                if dutyOff == None:
                    dutyOff = 'NA'


                # apply conditional color formatting
                # in general : dutyOn
                if not count % 2 == 0:
                    self.setColor(dateItem,240,240,240)
                else:
                    self.setColor(dateItem,255,255,255)

                # in general : dutyOff
                if not count % 2 == 0:
                    self.setColor(dateItemOut,240,240,240)
                else:
                    self.setColor(dateItemOut,255,255,255)


                # for specific : dutyOn
                if 'NA' in dutyOn:
                    if self.boolNA:
                        self.setColor(dateItem,self.colorNA[0],self.colorNA[1],self.colorNA[2])
                elif 'OD' in dutyOn:
                    if self.boolOD:
                        self.setColor(dateItem,self.colorOD[0],self.colorOD[1],self.colorOD[2])
                elif 'Leave' in dutyOn:
                    if self.boolLeave:
                        self.setColor(dateItem,self.colorLeave[0],self.colorLeave[1],self.colorLeave[2])
                elif 'WFH' in dutyOn:
                    if self.boolWFH:
                        self.setColor(dateItem,self.colorWFH[0],self.colorWFH[1],self.colorWFH[2])
                elif 'Holiday' in dutyOn:
                    if self.boolHoliday:
                        self.setColor(dateItem,self.colorHoliday[0],self.colorHoliday[1],self.colorHoliday[2])
                elif 'Holiday' not in dutyOn and (count+1) in holidayList:
                    if self.boolExtra:
                        self.setColor(dateItem,self.colorExtra[0],self.colorExtra[1],self.colorExtra[2])
                else:
                    if self.boolLate:
                        tmp = float(dutyOn.split(':')[0] + '.' + dutyOn.split(':')[1])
                        if(tmp > 10 and tmp <= 11):
                            self.setColor(dateItem,self.colorLate[0],self.colorLate[1],self.colorLate[2])
                        elif(tmp > 11):
                            self.setColor(dateItem,self.colorLate[0],self.colorLate[1],self.colorLate[2])


                # for specific : dutyOff
                if 'NA' in dutyOff:
                    if self.boolNA:
                        self.setColor(dateItemOut,self.colorNA[0],self.colorNA[1],self.colorNA[2])
                elif 'OD' in dutyOff:
                    if self.boolOD:
                        self.setColor(dateItemOut,self.colorOD[0],self.colorOD[1],self.colorOD[2])
                elif 'Leave' in dutyOff:
                    if self.boolLeave:
                        self.setColor(dateItemOut,self.colorLeave[0],self.colorLeave[1],self.colorLeave[2])
                elif 'WFH' in dutyOff:
                    if self.boolWFH:
                        self.setColor(dateItemOut,self.colorWFH[0],self.colorWFH[1],self.colorWFH[2])
                elif 'Holiday' in dutyOff:
                    if self.boolHoliday:
                        self.setColor(dateItemOut,self.colorHoliday[0],self.colorHoliday[1],self.colorHoliday[2])
                elif 'Holiday' not in dutyOff and (count+1) in holidayList:
                    if self.boolExtra:
                        self.setColor(dateItemOut,self.colorExtra[0],self.colorExtra[1],self.colorExtra[2])


                dateItem.setText(dutyOn)
                dateItemOut.setText(dutyOff)
                count = count + 1

            #closing connection
            con.commit()
            con.close()


    # table: event Handler
    @Slot()
    def employeeSelectionChanged(self):

        row = self.tableWidget.currentRow()
        empID = self.tableWidget.item(row,0).text()
        empName = self.tableWidget.item(row,1).text()

        #display selection
        self.labelSelectedEmpID.setText(empID)
        self.labelSelectedEmpName.setText(empName)

        # load calender
        self.loadCalenderData(empID)

        # refresh info
        self.guiInfo.refreshInfo(empID)


    # button: event Handler
    @Slot()
    def openGuiHolidays(self):

        #set processing label
        self.lbProcess.setHidden(False)

        #open gui_holiday
        gui = GUI_Holiday(self.dbname,self.year,self.month)
        t = gui.openDialog(self.frame)
        if t == 1:

            # refresh system
            # refresh calendar
            self.reloadCalenderData()
            self.loadCalenderData(self.labelSelectedEmpID.text())
            # refresh info
            # recalculate for all employee
            self.guiInfo.calculateAll()
            self.guiInfo.refreshInfo(self.labelSelectedEmpID.text())

        # reloading done, remove processing label
        self.lbProcess.setHidden(True)



    @Slot()
    def openGuiEdit(self,num):
	print 'opened !!'
        if 'Not Selected' not in self.labelSelectedEmpID.text():
            #open gui_holiday
            gui = GUI_Edit()
            t = gui.openDialog(self.frame,self.dbname,num,self.month,self.year,self.labelSelectedEmpID.text(),self.labelSelectedEmpName.text())
            if t == 1:
                # refresh calendar
                self.loadCalenderData(self.labelSelectedEmpID.text())
                # refresh info
                # recalculate info for empID
                self.guiInfo.calculate(self.labelSelectedEmpID.text())
                self.guiInfo.refreshInfo(self.labelSelectedEmpID.text())



    @Slot()
    def exportData(self):

            # export data to excel
            # open save dialog
            expObj = Export(self.dbname,self.month,self.year)
            expObj.exportToExcel(self.frame)
            QMessageBox.information(self.frame, 'Export' , 'Report successfully generated.    ')

    @Slot()
    def ImportData(self):

            # show confirmation dialog
            t = QMessageBox.warning (self.frame, 'Create New Attendance' ,
                                 'Are you sure you want to remove current attendance and want to load new one.\n\n'+
                                 '1. If yes, then press Ok. Application will close immediately. then Restart it.\n'+
                                 '2. Otherwise cancel this prompt to continue with same attendance.',
                                  QMessageBox.Yes | QMessageBox.Cancel )

            if t == QMessageBox.StandardButton.Yes:
                # remove db file and close the application
                os.remove(self.dbname)
                self.frame.close()



    @Slot()
    def setFilter(self):

            if self.popup == None:
                self.popup = QWidget(self.frame)
                self.popup.setMinimumSize(300,200)
                self.popup.setWindowFlags(Qt.Popup)
                vlayout = QVBoxLayout()
                #1
                hlayout = QHBoxLayout()
                self.radNA = QCheckBox('Show NA')
                self.radNA.setChecked(True)
                self.radNA.stateChanged.connect(self.filterSelected)
                hlayout.addWidget(self.radNA)
                colorLb = QLabel('          ')
                self.setColor(colorLb,self.colorNA[0],self.colorNA[1],self.colorNA[2])
                hlayout.addWidget(colorLb)
                vlayout.addLayout(hlayout)
                #2
                hlayout = QHBoxLayout()
                self.radOD = QCheckBox('Show OD')
                self.radOD.stateChanged.connect(self.filterSelected)
                hlayout.addWidget(self.radOD)
                colorLb = QLabel('          ')
                self.setColor(colorLb,self.colorOD[0],self.colorOD[1],self.colorOD[2])
                hlayout.addWidget(colorLb)
                vlayout.addLayout(hlayout)
                #3
                hlayout = QHBoxLayout()
                self.radLeave = QCheckBox('Show Leave')
                self.radLeave.stateChanged.connect(self.filterSelected)
                hlayout.addWidget(self.radLeave)
                colorLb = QLabel('          ')
                self.setColor(colorLb,self.colorLeave[0],self.colorLeave[1],self.colorLeave[2])
                hlayout.addWidget(colorLb)
                vlayout.addLayout(hlayout)
                #4
                hlayout = QHBoxLayout()
                self.radWFH = QCheckBox('Show WFH')
                self.radWFH.stateChanged.connect(self.filterSelected)
                hlayout.addWidget(self.radWFH)
                colorLb = QLabel('          ')
                self.setColor(colorLb,self.colorWFH[0],self.colorWFH[1],self.colorWFH[2])
                hlayout.addWidget(colorLb)
                vlayout.addLayout(hlayout)
                #5
                hlayout = QHBoxLayout()
                self.radHoliday = QCheckBox('Show Holidays')
                self.radHoliday.stateChanged.connect(self.filterSelected)
                hlayout.addWidget(self.radHoliday)
                colorLb = QLabel('          ')
                self.setColor(colorLb,self.colorHoliday[0],self.colorHoliday[1],self.colorHoliday[2])
                hlayout.addWidget(colorLb)
                vlayout.addLayout(hlayout)
                #6
                hlayout = QHBoxLayout()
                self.radExtra = QCheckBox('Show Extras')
                self.radExtra.stateChanged.connect(self.filterSelected)
                hlayout.addWidget(self.radExtra)
                colorLb = QLabel('          ')
                self.setColor(colorLb,self.colorExtra[0],self.colorExtra[1],self.colorExtra[2])
                hlayout.addWidget(colorLb)
                vlayout.addLayout(hlayout)
                #7
                hlayout = QHBoxLayout()
                self.radLate = QCheckBox('Show Lates')
                self.radLate.stateChanged.connect(self.filterSelected)
                hlayout.addWidget(self.radLate)
                colorLb = QLabel('          ')
                self.setColor(colorLb,self.colorLate[0],self.colorLate[1],self.colorLate[2])
                hlayout.addWidget(colorLb)
                vlayout.addLayout(hlayout)
                self.popup.setLayout(vlayout)

            #positioning and show
            pos = self.btnFilter.rect().topLeft()
            gpos = self.btnFilter.mapToGlobal(pos)
            self.popup.move(gpos.x(),gpos.y())
            self.popup.show()


    @Slot()
    def clearFilter(self):

         #uncheck all except NA
         self.radNA.setChecked(True)
         self.radOD.setChecked(False)
         self.radLeave.setChecked(False)
         self.radWFH.setChecked(False)
         self.radHoliday.setChecked(False)
         self.radExtra.setChecked(False)
         self.radLate.setChecked(False)

         #now refresh filter
         self.filterSelected()



    @Slot()
    def filterSelected(self):
         #change bool variables
         try:
             self.boolNA = self.radNA.isChecked()
             self.boolOD = self.radOD.isChecked()
             self.boolLeave = self.radLeave.isChecked()
             self.boolWFH = self.radWFH.isChecked()
             self.boolHoliday = self.radHoliday.isChecked()
             self.boolExtra = self.radExtra.isChecked()
             self.boolLate = self.radLate.isChecked()
         except:
                pass    #do nothing

         # refresh calendar
         self.loadCalenderData(self.labelSelectedEmpID.text())


    @Slot()
    def exit(self):
        # exit the application
        sys.exit(1)
