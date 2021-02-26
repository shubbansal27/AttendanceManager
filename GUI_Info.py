__author__ = 'admin'

import math
import calendar
import sqlite3
from PySide2.QtGui import *
from PySide2.QtWidgets import *
from PySide2.QtCore import *

class GUI_Info:

    def __init__(self,dbname,month,year):
        self.dbname = dbname
        self.month = month
        self.year = year


    def createInfoPanel(self):

        hlayout = QHBoxLayout()

        #1
        form1 = QFrame()
        form1.setFont(QFont('Arial',10,QFont.StyleItalic))
        vlayout = QVBoxLayout()
        #1.1
        partition1 = QFrame()
        self.setColor(partition1,240,210,190)
        hboxLayout = QHBoxLayout()
        lb = QLabel('Total Present')
        lb.setFont(QFont('Arial',12,QFont.Bold))
        lb.setAlignment(Qt.AlignCenter)
        hboxLayout.addWidget(lb)
        self.lbTotalPresent = QLabel('0')
        self.lbTotalPresent.setFont(QFont('Arial',12,QFont.Bold))
        self.lbTotalPresent.setAlignment(Qt.AlignCenter)
        hboxLayout.addWidget(self.lbTotalPresent)
        partition1.setLayout(hboxLayout)
        vlayout.addWidget(partition1)
        #1.2
        partition2 = QFrame()
        self.setColor(partition2,240,210,190)
        hboxLayout = QHBoxLayout()
        lb = QLabel('Total Days    ')
        lb.setFont(QFont('Arial',12,QFont.Bold))
        lb.setAlignment(Qt.AlignCenter)
        hboxLayout.addWidget(lb)
        self.lbTotalDays = QLabel('0')
        self.lbTotalDays.setFont(QFont('Arial',12,QFont.Bold))
        self.lbTotalDays.setAlignment(Qt.AlignCenter)
        hboxLayout.addWidget(self.lbTotalDays)
        partition2.setLayout(hboxLayout)
        vlayout.addWidget(partition2)

        form1.setLayout(vlayout)
        hlayout.addWidget(form1)

        #2
        form2 = QFrame()
        form2.setFont(QFont('Arial',10,QFont.StyleItalic))
        vlayout = QVBoxLayout()
        #2.1
        partition1 = QFrame()
        self.setColor(partition1,190,210,210)
        formLayout = QFormLayout()
        self.lbNumberNA = QLabel('0')
        self.lbNumberNA.setAlignment(Qt.AlignCenter)
        self.lbNumberNA.setFont(QFont('Arial',10,QFont.Bold))
        formLayout.addRow('  Total NA:        ',self.lbNumberNA)
        partition1.setLayout(formLayout)
        #2.2
        partition2 = QFrame()
        self.setColor(partition2,190,210,210)
        formLayout = QFormLayout()
        formLayout.setVerticalSpacing(25)
        self.lbNumberOD = QLabel('0')
        self.lbNumberOD.setAlignment(Qt.AlignCenter)
        self.lbNumberOD.setFont(QFont('Arial',10,QFont.Bold))
        self.lbNumberWFH = QLabel('0')
        self.lbNumberWFH.setAlignment(Qt.AlignCenter)
        self.lbNumberWFH.setFont(QFont('Arial',10,QFont.Bold))
        self.lbNumberLeave = QLabel('0')
        self.lbNumberLeave.setAlignment(Qt.AlignCenter)
        self.lbNumberLeave.setFont(QFont('Arial',10,QFont.Bold))
        formLayout.addRow('  Total OD:       ',self.lbNumberOD)
        formLayout.addRow('  Total WFH:      ',self.lbNumberWFH)
        formLayout.addRow('  Total Leaves:   ',self.lbNumberLeave)
        partition2.setLayout(formLayout)
        #2.3
        partition3 = QFrame()
        self.setColor(partition3,190,210,210)
        formLayout = QFormLayout()
        formLayout.setVerticalSpacing(25)
        self.lbNumberHoliday = QLabel('0')
        self.lbNumberHoliday.setAlignment(Qt.AlignCenter)
        self.lbNumberHoliday.setFont(QFont('Arial',10,QFont.Bold))
        self.lbNumberExtra = QLabel('0')
        self.lbNumberExtra.setAlignment(Qt.AlignCenter)
        self.lbNumberExtra.setFont(QFont('Arial',10,QFont.Bold))
        formLayout.addRow('  Total Holidays: ',self.lbNumberHoliday)
        formLayout.addRow('  Total Extras:   ',self.lbNumberExtra)
        partition3.setLayout(formLayout)

        vlayout.addWidget(partition1)
        vlayout.addWidget(partition2)
        vlayout.addWidget(partition3)
        form2.setLayout(vlayout)
        hlayout.addWidget(form2)

        #3
        form3 = QFrame()
        form3.setFont(QFont('Arial',10,QFont.StyleItalic))
        self.setColor(form3,210,190,210)
        formLayout = QFormLayout()
        formLayout.setVerticalSpacing(25)
        self.lbNumberLate = QLabel('0')
        self.lbNumberLate.setFont(QFont('Arial',10,QFont.Bold))
        self.lbNumberLate.setAlignment(Qt.AlignCenter)
        btnEditLate = QPushButton('Edit')
        btnEditLate.setEnabled(False)
        formLayout.addRow('  Number of Lates:       ',self.lbNumberLate)
        formLayout.addRow('',btnEditLate)
        form3.setLayout(formLayout)
        hlayout.addWidget(form3)

        return hlayout


#####################################################3
    def setColor(self,widget,R,G,B):

         p = widget.palette()
         p.setColor(widget.backgroundRole(),QColor(R,G,B))
         widget.setAutoFillBackground(True)
         widget.setPalette(p)


####################################################

    def refreshInfo(self,empID):

        #display on info labels
        if 'Not Selected' not in empID:
            # opening connection
            con = sqlite3.connect(self.dbname)

            # select query
            # NA ,OD ,Leave ,WFH ,Holiday ,Extra ,Late ,Present ,TotalDays
            query = "select NA,OD,Leave,WFH,Holiday,Extra,Late,Present,TotalDays from Calculations where empID = '" + empID + "'"
            data = con.execute(query).fetchone()

            # display on labels
            self.lbNumberNA.setText(str(self.format(data[0])))         #NA
            self.lbNumberOD.setText(str(self.format(data[1])))         #OD
            self.lbNumberLeave.setText(str(self.format(data[2])))      #Leave
            self.lbNumberWFH.setText(str(self.format(data[3])))        #WFH
            self.lbNumberHoliday.setText(str(self.format(data[4])))    #Holiday
            self.lbNumberExtra.setText(str(self.format(data[5])))      #Extra
            self.lbNumberLate.setText(str(self.format(data[6])))       #Late
            self.lbTotalPresent.setText(str(self.format(data[7])))     #Present
            self.lbTotalDays.setText(str(self.format(data[8])))        #TotalDays

            # closing connection
            con.close()


###################################################

    def format(self,num):

        #if num is integer then convert into int
        if num-int(num) == 0:
            return int(num)
        else:
            return num


####################################################

    def calculateAll(self):

        # opening connection
        con = sqlite3.connect(self.dbname)

        empList = list()
        #get empID
        query = 'select empID from AttendanceIn'
        cursor = con.execute(query)
        for row in cursor:
            empList.append(row[0])

        #closing connection
        con.close()

        #calculate info for each empID
        for empID in empList:
            self.calculate(empID)


####################################################

    def calculate(self,empID):

         if 'Not Selected' not in empID:

            #no of days
            numDays = calendar.monthrange(self.year,self.month)[1]
            totalNA = 0
            totalOD = 0
            totalLeave = 0
            totalWFH = 0
            totalHoliday = 0
            totalExtra = 0
            late10 = 0
            late11 = 0

            # opening connection
            con = sqlite3.connect(self.dbname)

            ##getting holidays list
            query = 'select date from Holidays'
            holidayList = list()
            cursor = con.execute(query)
            for row in cursor:
                date = int(row[0].split('/')[1])
                holidayList.append(date)


            #getting AttendanceIn and out
            queryIn = "select * from AttendanceIn where empID = '" + empID + "'"
            cursorIn = con.execute(queryIn)
            queryOut = "select * from AttendanceOut where empID = '" + empID + "'"
            cursorOut = con.execute(queryOut)
            dataIn = cursorIn.fetchone()
            dataOut = cursorOut.fetchone()
            for i in range(1,numDays+1,1):
                strIn = dataIn[i+1]
                strOut = dataOut[i+1]
                if strIn == None:
                    strIn = 'NA'
                if strOut == None:
                    strOut = 'NA'


                ## Duty-on
                # calculate NA
                if 'NA' in strIn:
                    totalNA = totalNA + 0.5
                # calculate OD
                elif 'OD' in strIn:
                    totalOD = totalOD + 0.5
                # calculate Leave
                elif 'Leave' in strIn:
                    totalLeave = totalLeave + 0.5
                # calculate WFH
                elif 'WFH' in strIn:
                    totalWFH = totalWFH + 0.5
                # calculate Holiday
                elif 'Holiday' in strIn:
                    totalHoliday = totalHoliday + 0.5
                 # calculate Extra
                elif 'Holiday' not in strIn and i in holidayList:
                    totalExtra = totalExtra + 0.5
                # calculate late
                # replace with policy
                else:
                    tmp = float(strIn.split(':')[0] + '.' + strIn.split(':')[1])
                    if(tmp > 10 and tmp <= 11):
                        late10 = late10 + 1
                    elif(tmp > 11):
                        late11 = late11 + 1


                ## Duty-off
                # calculate NA
                if 'NA' in strOut:
                    totalNA = totalNA + 0.5
                # calculate OD
                elif 'OD' in strOut:
                    totalOD = totalOD + 0.5
                # calculate Leave
                elif 'Leave' in strOut:
                    totalLeave = totalLeave + 0.5
                # calculate WFH
                elif 'WFH' in strOut:
                    totalWFH = totalWFH + 0.5
                # calculate Holiday
                elif 'Holiday' in strOut:
                    totalHoliday = totalHoliday + 0.5
                # calculate Extra
                elif 'Holiday' not in strOut and i in holidayList:
                    totalExtra = totalExtra + 0.5



            # push calc into database
            #NA int,OD int,Leave int,WFH int,Holiday int,Extra int,Late int,Present int,TotalDays int
            totalLate = math.floor(late10/3)/2 + late11/2
            totalPresent = numDays - (totalLate + totalNA + totalLeave + totalHoliday)
            query = "update Calculations set NA = "+str(totalNA)+",OD ="+str(totalOD)+",Leave ="+str(totalLeave)+",WFH ="+str(totalWFH)+",Holiday ="+str(totalHoliday)+",Extra ="+str(totalExtra)+",Late ="+str(totalLate)+",Present ="+str(totalPresent)+",TotalDays =" +str(numDays) +" where empID = '" + empID + "'"
            con.execute(query)
            con.commit()

            # closing connection
            con.close()
