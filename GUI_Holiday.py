__author__ = 'Administrator'

from PySide2.QtGui import *
from PySide2.QtWidgets import *
from PySide2.QtCore import Slot,QDate
import calendar
import sqlite3

class GUI_Holiday:

    def __init__(self,dbname,year,month):
        self.dbname = dbname
        self.year = year
        self.month = month

    def openDialog(self,frame):

        self.returnVal = 0
        self.dialog = QDialog(frame)
        self.dialog.setWindowTitle('set holidays')
        self.dialog.setMinimumSize(600,420)
        # dialog.windowTitle('Set Holidays')
        vlayout = QVBoxLayout()

        #adding add/remove button
        hlayout = QHBoxLayout()
        #1
        btnAdd = QPushButton('Add')
        btnAdd.clicked.connect(self.addHoliday)
        btnAdd.setMaximumSize(50,40)
        hlayout.addWidget(btnAdd)
        #2
        self.btnRemove = QPushButton('Remove')
        self.btnRemove.setEnabled(False)
        self.btnRemove.clicked.connect(self.removeHoliday)
        self.btnRemove.setMaximumSize(50,40)
        hlayout.addWidget(self.btnRemove)
        vlayout.addLayout(hlayout)

        ##adding table of holidays
        self.tableWidget = QTableWidget(0, 2)
        self.tableWidget.itemSelectionChanged.connect(self.tableFocusGained)
        colHeader1 = QTableWidgetItem('Date')
        colHeader2 = QTableWidgetItem('Holiday Title')
        self.tableWidget.setColumnWidth(0,200)
        self.tableWidget.setColumnWidth(1,350)
        self.tableWidget.setHorizontalHeaderItem(0,colHeader1)
        self.tableWidget.setHorizontalHeaderItem(1,colHeader2)
        vlayout.addWidget(self.tableWidget)

        #adding Save/Cancel button
        hlayout = QHBoxLayout()
        #1
        btnSave = QPushButton('Save')
        btnSave.clicked.connect(self.saveHolidays)
        btnSave.setMaximumSize(50,40)
        hlayout.addWidget(btnSave)
        #2
        btnCancel = QPushButton('Cancel')
        btnCancel.clicked.connect(self.cancelButton)
        btnCancel.setMaximumSize(50,40)
        hlayout.addWidget(btnCancel)
        vlayout.addLayout(hlayout)


        #add holiday data from local.db
        self.getHolidaysList()

        self.dialog.setLayout(vlayout)

        self.dialog.exec_()
        return self.returnVal


    def getHolidaysList(self):

         # opening connection
        con = sqlite3.connect(self.dbname)
        #select query
        query = 'select date,title from Holidays'
        cursor = con.execute(query)

        r = 0
        for row in cursor:

            date = int(row[0].split('/')[1])
            title = row[1]

            self.tableWidget.insertRow(r)
            dateEdit = QDateEdit()
            dateEdit.setDisplayFormat('d-MMM')
            dateEdit.setDate(QDate(self.year,self.month,date))
            dateEdit.setDateRange(QDate(self.year, self.month, 1), QDate(self.year, self.month, calendar.monthrange(self.year,self.month)[1]))
            cal = QCalendarWidget()
            cal.setNavigationBarVisible(False)
            cal.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
            dateEdit.setCalendarWidget(cal)
            dateEdit.setCalendarPopup(True)
            self.tableWidget.setCellWidget(r,0,dateEdit)
            self.tableWidget.setItem(r,1,QTableWidgetItem(title))

            r = r +1

        # close connection
        con.close()


    @Slot()
    def addHoliday(self):

         # add row
        lastRow = self.tableWidget.rowCount()
        self.tableWidget.insertRow(lastRow)

        dateEdit = QDateEdit()
        dateEdit.setDisplayFormat('d-MMM')
        dateEdit.setDateRange(QDate(self.year, self.month, 1), QDate(self.year, self.month, calendar.monthrange(self.year,self.month)[1]))
        cal = QCalendarWidget()
        cal.setNavigationBarVisible(False)
        cal.setVerticalHeaderFormat(QCalendarWidget.NoVerticalHeader)
        dateEdit.setCalendarWidget(cal)
        dateEdit.setCalendarPopup(True)
        self.tableWidget.setCellWidget(lastRow,0,dateEdit)


    @Slot()
    def saveHolidays(self):

        # opening connection
        con = sqlite3.connect(self.dbname)

        #flag
        flagSuccess = False
        flagType = 0
        errMsg = ''

        #delete query
        #delete all rows from Holidays
        query = 'delete from Holidays'
        cursor = con.execute(query)
        #con.commit()

        #now save new records
        rows = self.tableWidget.rowCount()
        for i in range(0,rows,1):

            date = str(self.month) + '/' + str(int(str(self.tableWidget.cellWidget(i,0).text().split('-')[0]))) + '/' + str(self.year)
            if self.tableWidget.item(i,1) is not None:
                title = self.tableWidget.item(i,1).text().strip()
                #check not empty string
                if title:

                    try :
                        query = "insert into Holidays values('"+ date +"','"+title+"')"
                        con.execute(query)
                        flagSuccess = True
                        #con.commit()
                    except:
                        errMsg = 'Error: More then one occurences for same date !!'
                        flagSuccess = False
                        flagType = 2
                        break

                else:
                    # print '1 holiday title is blank for date ' + str(date)
                    flagType = 1
                    flagSuccess = False
                    break
            else:
                # print '2 holiday title is blank for date ' + str(date)
                flagSuccess = False
                flagType = 1
                break


        if flagSuccess:
            con.commit()
            self.dialog.close()
        else:
            if flagType == 2:
                errMsgBox = QMessageBox()
                errMsgBox.setWindowTitle('Error')
                errMsgBox.setText('Error: More than one occurences of same date !!')
                errMsgBox.exec_()
            elif flagType == 1:
                errMsgBox = QMessageBox()
                errMsgBox.setWindowTitle('Error')
                errMsgBox.setText('Error: Blank value in Holiday title !!')
                errMsgBox.exec_()
        con.close()
        self.returnVal = 1


    @Slot()
    def cancelButton(self):

        #close dialog
        self.dialog.close()
        self.returnVal = 0


    @Slot()
    def removeHoliday(self):

        #remove selected row
        currentRow = self.tableWidget.currentRow()
        if currentRow >= 0:
            self.tableWidget.removeRow(currentRow)
        else:
            #set disable
            self.btnRemove.setEnabled(False)


    @Slot()
    def tableFocusGained(self):
        #set enable
        self.btnRemove.setEnabled(True)


