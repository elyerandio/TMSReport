#***********************************
# Program Name : TMSReport.py
# Date Written : September 17, 2013
# Description  : a program to process the TMS process report and save to tms_report table
# Author       : Eleazer L. Erandio
#************************************

import sys
import os
from PySide.QtCore import *
from PySide.QtGui import *
import pyodbc
from datetime import *
import re
import ConfigParser


class TMSReport(QMainWindow):
    def __init__(self, parent=None):
        super(TMSReport, self).__init__(parent)
        form = TMSReportForm()
        self.setCentralWidget(form)
        self.setWindowTitle('TMSWin Report Process')
        #form.show()

class TMSReportForm(QDialog):

    def __init__(self, parent=None):
        global dateFromPrev, dateToPrev

        # get the to-date of last process, then add 1 day to get the next process date
        (pyear, pmonth, pday) = dateToPrev.split('-')
        nextDate = date(int(pyear), int(pmonth), int(pday))
        nextDate = nextDate + timedelta(1)
        super(TMSReportForm, self).__init__(parent)

        fromLabel = QLabel('From')
        fromLabel.setAlignment(Qt.AlignHCenter)
        toLabel = QLabel('To')
        toLabel.setAlignment(Qt.AlignHCenter)
        dateLabel = QLabel('Date')
        self.dateEditFrom = QDateTimeEdit(date(nextDate.year, nextDate.month, nextDate.day))
        self.dateEditFrom.setCalendarPopup(True)
        self.dateEditFrom.setMaximumDate(QDate.currentDate())
        self.dateEditTo = QDateTimeEdit(date(nextDate.year, nextDate.month, nextDate.day))
        self.dateEditTo.setCalendarPopup(True)
        self.labelStatus = QLabel('')
        self.labelSaveAs = QLabel('Save As')

        self.processButton = QPushButton('Process')
        self.cancelButton = QPushButton('Cancel')
        buttonLayout = QHBoxLayout()
        buttonLayout.addSpacing(2)
        buttonLayout.addWidget(self.processButton)
        buttonLayout.addWidget(self.cancelButton)

        layout = QGridLayout()
        layout.addWidget(fromLabel, 0, 1)
        layout.addWidget(toLabel, 0, 2)
        layout.addWidget(dateLabel, 3, 0)
        layout.addWidget(self.dateEditFrom, 3, 1)
        layout.addWidget(self.dateEditTo, 3, 2)
        layout.addLayout(buttonLayout, 6, 1, 1, 2)
        layout.addWidget(self.labelStatus, 7, 0, 1, 3)

        self.dateEditFrom.dateChanged.connect(self.setDateTo)
        self.dateEditTo.dateChanged.connect(self.setDateFrom)
        self.connect(self.processButton, SIGNAL('clicked()'), self.process)
        self.connect(self.cancelButton, SIGNAL('clicked()'), self.canceled)

        self.setLayout(layout)
        self.setWindowTitle('TMSWin Report Process')


    def setDateFrom(self):
        # don't allow dateEditTo.date to be less than dateEditFrom.date
        # set dateEditFrom.date to dateEditTo.date
        if self.dateEditTo.date() < self.dateEditFrom.date():
            self.dateEditFrom.setDate(self.dateEditTo.date())

        if self.dateEditTo.date() > QDate.currentDate():
            self.dateEditTo.setDate(QDate.currentDate())

    def setDateTo(self):
        # set dateEditTo.date to dateEditFrom.date
        self.dateEditTo.setDate(self.dateEditFrom.date())

        if self.dateEditFrom.date() > QDate.currentDate():
            self.dateEditFrom.setDate(QDate.currentDate())


    def process(self):

        # show the confirmation message
        flags = QMessageBox.StandardButton.Yes
        flags |= QMessageBox.StandardButton.No
        question = "Do you really want to process right now?"
        response = QMessageBox.question(self, "Confirm Process", question, flags, QMessageBox.No)
        if response == QMessageBox.No:
            return

        self.getActiveEmployees()
        self.countNumberDaysWorked()
        self.getTransactions()
        self.getLeaveInfo()
        self.saveReport()
        self.labelStatus.setText('Process finished!')
        self.saveIni()
        self.processButton.setEnabled(False)
        self.cancelButton.setText('Exit')


    def createEmployee(self):

        emp = {}
        emp['employee_id'] = ''
        emp['lastname'] = ''
        emp['firstname'] = ''
        emp['shift_schedule'] = ''
        emp['restday_schedule'] = ''
        emp['days_worked'] = 0
        emp['vl'] = 0
        emp['sl'] = 0
        emp['lwop'] = 0
        emp['bl'] = 0
        emp['oil'] = 0
        emp['ml'] = 0
        emp['marriage_days'] = 0
        emp['cl'] = 0
        emp['spl'] = 0
        emp['pl'] = 0
        emp['absent'] = 0
        emp['suspension'] = 0
        emp['late_undertime'] = 0
        emp['nd1'] = 0
        emp['ot1'] = 0
        emp['nd2'] = 0
        emp['ot2'] = 0
        emp['nd3'] = 0
        emp['ot3'] = 0
        emp['nd4'] = 0
        emp['ot4'] = 0
        emp['nd5'] = 0
        emp['ot5'] = 0
        emp['nd6'] = 0
        emp['ot6'] = 0
        emp['nd7'] = 0
        emp['ot7'] = 0
        emp['nd8'] = 0
        emp['ot8'] = 0
        emp['nd9'] = 0
        emp['ot9'] = 0
        emp['nd10'] = 0
        emp['ot10'] = 0
        emp['nd11'] = 0
        emp['ot11'] = 0
        emp['nd12'] = 0
        emp['ot12'] = 0
        return emp

    def getActiveEmployees(self):
		cur = connOriTMS.cursor()
		cur.execute("Select eb.employee_id, eb.employee_no, eb.employee_name from employee_biodata eb, employee_employment ee" \
			" where eb.employee_id = ee.employee_id and ee.employee_status = 'A' ")

		for rec in cur:
			employee_id = rec[0]
			employee_no = rec[1]
			fullname = rec[2]

			if fullname.count(',') > 1:
				(lastname, firstname, whatever) = fullname.split(',', 2)
			elif fullname.count(',') == 1:
				(lastname, firstname) = fullname.split(',')
			else:
				lastname = fullname
				firstname = ''

			employees[employee_no] = self.createEmployee()
			employees[employee_no]['employee_id'] = employee_id
			employees[employee_no]['lastname'] = lastname
			employees[employee_no]['firstname'] = firstname

    def countNumberDaysWorked(self):
        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()
        cur = connOriTMS.cursor()
        cur.execute("Select employee_no, actual_date, schedule_type, time_in1, time_out1, total_work_hour" \
                    " from employee_attendance where actual_date between '%s' and '%s'"\
                    " order by employee_no, actual_date " % (dateFrom.toPython(), dateTo.toPython()))

        sv_employee_no = ''
        days_count = 0
        for rec in cur:
            employee_no = rec[0]
            actual_date = rec[1].date().isoformat()
            schedule_type = rec[2]
            time_in = rec[3]
            time_out = rec[4]
            work_hour = rec[5]

            if employee_no != sv_employee_no:
                if sv_employee_no != '':
                    # save previous employee's data to employees hash
                    if sv_employee_no not in employees:
                        if days_count > 0:
                            employees[sv_employee_no] = self.createEmployee()
                            employees[sv_employee_no]['shift_schedule'] = 'DaysWorked'
                            employees[sv_employee_no]['days_worked'] = days_count
                    else:
                        employees[sv_employee_no]['days_worked'] = days_count

                # reset count for current employee
                sv_employee_no = employee_no
                days_count = 0

            if work_hour.hour >= 5:
                days_count += 1

    def getTransactions(self):
        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()
        cur = connOriTMS.cursor()
        cur.execute("Select employee_no, trx_date, trx_type, trx_code, rate, qty, amount, posted_status" \
                    " from employee_trxldg where trx_date between '%s' and '%s'"\
                    " order by employee_no, trx_date " % (dateFrom.toPython(), dateTo.toPython()))

        for rec in cur:
            employee_no = rec[0]
            trx_date = rec[1].date().isoformat()
            trx_type = rec[2]
            trx_code = rec[3]
            rate = rec[4]
            qty = rec[5]
            amount = rec[6]
            posted_status = rec[7]

            if employee_no not in employees:
                employees[employee_no] = self.createEmployee()
                employees[sv_employee_no]['shift_schedule'] = 'getTran'

            if trx_code == 'LATE' or trx_code == 'OVBR':
                if 'late_undertime' not in employees[employee_no]:
                    employees[employee_no]['late_undertime'] = 0
                employees[employee_no]['late_undertime'] += qty
            elif trx_code == 'LWOP':
                if 'lwop' not in employees[employee_no]:
                    employees[employee_no]['lwop'] = 0
                employees[employee_no]['lwop'] += qty
            elif trx_code == 'ND1':
                if 'nd1' not in employees[employee_no]:
                    employees[employee_no]['nd1'] = 0
                employees[employee_no]['nd1'] += qty
            elif trx_code == 'ND2':
                if 'nd2' not in employees[employee_no]:
                    employees[employee_no]['nd2'] = 0
                employees[employee_no]['nd2'] += qty
            elif trx_code == 'ND3':
                if 'nd3' not in employees[employee_no]:
                    employees[employee_no]['nd3'] = 0
                employees[employee_no]['nd3'] += qty
            elif trx_code == 'ND4':
                if 'nd4' not in employees[employee_no]:
                    employees[employee_no]['nd4'] = 0
                employees[employee_no]['nd4'] += qty
            elif trx_code == 'ND5':
                if 'nd5' not in employees[employee_no]:
                    employees[employee_no]['nd5'] = 0
                employees[employee_no]['nd5'] += qty
            elif trx_code == 'ND6':
                if 'nd6' not in employees[employee_no]:
                    employees[employee_no]['nd6'] = 0
                employees[employee_no]['nd6'] += qty
            elif trx_code == 'ND7':
                if 'nd7' not in employees[employee_no]:
                    employees[employee_no]['nd7'] = 0
                employees[employee_no]['nd7'] += qty
            elif trx_code == 'ND8':
                if 'nd8' not in employees[employee_no]:
                    employees[employee_no]['nd8'] = 0
                employees[employee_no]['nd8'] += qty
            elif trx_code == 'ND9':
                if 'nd9' not in employees[employee_no]:
                    employees[employee_no]['nd9'] = 0
                employees[employee_no]['nd9'] += qty
            elif trx_code == 'ND10':
                if 'nd10' not in employees[employee_no]:
                    employees[employee_no]['nd10'] = 0
                employees[employee_no]['nd10'] += qty
            elif trx_code == 'ND11':
                if 'nd11' not in employees[employee_no]:
                    employees[employee_no]['nd11'] = 0
                employees[employee_no]['nd11'] += qty
            elif trx_code == 'ND12':
                if 'nd12' not in employees[employee_no]:
                    employees[employee_no]['nd12'] = 0
                employees[employee_no]['nd12'] += qty
            elif trx_code == 'OT1':
                if 'ot1' not in employees[employee_no]:
                    employees[employee_no]['ot1'] = 0
                employees[employee_no]['ot1'] += qty
            elif trx_code == 'OT2':
                if 'ot2' not in employees[employee_no]:
                    employees[employee_no]['ot2'] = 0
                employees[employee_no]['ot2'] += qty
            elif trx_code == 'OT3':
                if 'ot3' not in employees[employee_no]:
                    employees[employee_no]['ot3'] = 0
                employees[employee_no]['ot3'] += qty
            elif trx_code == 'OT4':
                if 'ot4' not in employees[employee_no]:
                    employees[employee_no]['ot4'] = 0
                employees[employee_no]['ot4'] += qty
            elif trx_code == 'OT5':
                if 'ot5' not in employees[employee_no]:
                    employees[employee_no]['ot5'] = 0
                employees[employee_no]['ot5'] += qty
            elif trx_code == 'OT6':
                if 'ot6' not in employees[employee_no]:
                    employees[employee_no]['ot6'] = 0
                employees[employee_no]['ot6'] += qty
            elif trx_code == 'OT7':
                if 'ot7' not in employees[employee_no]:
                    employees[employee_no]['ot7'] = 0
                employees[employee_no]['ot7'] += qty
            elif trx_code == 'OT8':
                if 'ot8' not in employees[employee_no]:
                    employees[employee_no]['ot8'] = 0
                employees[employee_no]['ot8'] += qty
            elif trx_code == 'OT9':
                if 'ot9' not in employees[employee_no]:
                    employees[employee_no]['ot9'] = 0
                employees[employee_no]['ot9'] += qty
            elif trx_code == 'OT10':
                if 'ot10' not in employees[employee_no]:
                    employees[employee_no]['ot10'] = 0
                employees[employee_no]['ot10'] += qty
            elif trx_code == 'OT11':
                if 'ot11' not in employees[employee_no]:
                    employees[employee_no]['ot11'] = 0
                employees[employee_no]['ot11'] += qty
            elif trx_code == 'OT12':
                if 'ot12' not in employees[employee_no]:
                    employees[employee_no]['ot12'] = 0
                employees[employee_no]['ot12'] += qty

    def getLeaveInfo(self):
        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()
        cur = connOriTMS.cursor()
        cur.execute("Select li.employee_no, ld.leave_date, ld.status, ld.approve_date, li.leave_code from employee_leave_day ld, employee_leave_info li" \
            " where ld.employee_id = li.employee_id and ld.reference_id = li.reference_no and ld.status in ('P') and " \
            "ld.leave_date between '%s' and '%s'" % (dateFrom.toPython(), dateTo.toPython()))

        for rec in cur:
            employee_no = rec[0]
            leave_date = rec[1].date().isoformat()
            status = rec[2]
            approve_date = rec[3].date().isoformat()
            leave_code = rec[4]

            if employee_no not in employees:
                employees[employee_no] = self.createEmployee()
                employees[sv_employee_no]['shift_schedule'] = 'getLeave'

            if leave_code == 'VL':
                if 'vl' not in employees[employee_no]:
                    employees[employee_no]['vl'] = 0
                employees[employee_no]['vl'] += 1
            elif leave_code == 'SL':
                if 'sl' not in employees[employee_no]:
                    employees[employee_no]['sl'] = 0
                employees[employee_no]['sl'] += 1
            elif leave_code == 'LWOP':
                if 'lwop' not in employees[employee_no]:
                    employees[employee_no]['lwop'] = 0
                employees[employee_no]['lwop'] += 1
            elif leave_code == 'BL':
                if 'bl' not in employees[employee_no]:
                    employees[employee_no]['bl'] = 0
                employees[employee_no]['bl'] += 1
            elif leave_code == 'OIL':
                if 'oil' not in employees[employee_no]:
                    employees[employee_no]['oil'] = 0
                employees[employee_no]['oil'] += 1
            elif leave_code == 'ML':
                if 'ml' not in employees[employee_no]:
                    employees[employee_no]['ml'] = 0
                employees[employee_no]['ml'] += 1
            elif leave_code == 'CL':
                if 'cl' not in employees[employee_no]:
                    employees[employee_no]['cl'] = 0
                employees[employee_no]['cl'] += 1
            elif leave_code == 'SPL':
                if 'spl' not in employees[employee_no]:
                    employees[employee_no]['spl'] = 0
                employees[employee_no]['spl'] += 1
            elif leave_code == 'PL':
                if 'pl' not in employees[employee_no]:
                    employees[employee_no]['pl'] = 0
                employees[employee_no]['pl'] += 1
            elif leave_code == 'SUSP':
                if 'suspension' not in employees[employee_no]:
                    employees[employee_no]['suspension'] = 0
                employees[employee_no]['suspension'] += 1


    def saveReport(self):
        dateFrom = self.dateEditFrom.date()
        dateTo = self.dateEditTo.date()
        now = datetime.today()
        cur = connOriTMS.cursor()

        # truncate the table
        cur.execute("Truncate table user_tms_report")

        for emp_no in sorted(employees):
            emp = employees[emp_no]
            cur.execute("INSERT into user_tms_report ( EMPLOYEE_NO, SURNAME, FIRSTNAME, SHIFT_SCHEDULE, RESTDAY_SCHEDULE, DATE_FROM, DATE_TO, "\
                "DAYS_WORKED, VL, SL, LWOP, BL, OIL, ML, MARRIAGE_DAYS, CL, SPL, PL, ABSENT, SUSPENSION, LATE_UNDERTIME, "\
                "ND1, OT1, ND2, OT2, ND3, OT3, ND4, OT4, ND5, OT5, OT6, ND6, OT7, ND7, OT8, ND8, OT9, ND9, OT10, ND10, OT11, ND11, OT12, ND12, "\
                "CREATED_BY, CREATED_DATE) "\
                " VALUES('%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', "\
                "'%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s', '%s')" % \
                (emp_no, emp['lastname'], emp['firstname'], emp['shift_schedule'], emp['restday_schedule'], dateFrom.toPython(), dateTo.toPython(), emp['days_worked'],
                emp['vl'], emp['sl'],  emp['lwop'], emp['bl'], emp['oil'], emp['ml'], emp['marriage_days'], emp['cl'], emp['spl'], emp['pl'],
                emp['absent'], emp['suspension'], emp['late_undertime'],
                emp['nd1'], emp['ot1'], emp['nd2'], emp['ot2'],  emp['nd3'], emp['ot3'],  emp['nd4'], emp['ot4'],  emp['nd5'], emp['ot5'],
                emp['nd6'], emp['ot6'],  emp['nd7'], emp['ot7'],  emp['nd8'], emp['ot8'],  emp['nd9'], emp['ot9'],  emp['nd10'], emp['ot10'],
                emp['nd11'], emp['ot11'],  emp['nd12'], emp['ot12'], "TMSReport", now.strftime('%Y-%m-%d %H:%M:%S')))

        connOriTMS.commit()

    def canceled(self):
        # show the confirmation message
        if self.cancelButton.text() == 'Cancel':
            flags = QMessageBox.StandardButton.Yes
            flags |= QMessageBox.StandardButton.No
            question = "Do you really want to cancel?"
            response = QMessageBox.question(self, "Confirm Cancel", question, flags, QMessageBox.No)
            if response == QMessageBox.No:
                return

        self.abort()

    def abort(self):
        connOriTMS.close()
        self.reject()
        app.exit(1)

    def saveIni(self):

        # name of configuration file
        iniFile = 'TMSReport.ini'
        config = ConfigParser.ConfigParser()
        config.read(iniFile)

        # save dateFrom/dateTo
        config.set('History','DateFrom',self.dateEditFrom.date().toPython())
        config.set('History', 'DateTo', self.dateEditTo.date().toPython())

        ini = open(iniFile, 'w')
        config.write(ini)
        ini.close()



def readIni():
    global orisoftDsn, orisoftUser, orisoftPwd
    global dateFromPrev, dateToPrev

    # name of configuration file
    iniFile = 'TMSReport.ini'
    # test if config file exists
    if not os.path.exists(iniFile):
        QMessageBox.critical(None,'Config File Missing', "The configuration file '%s' in '%s' not found!" % \
            (iniFile, os.getcwd()))
        app.exit(1)

    try:
        config = ConfigParser.ConfigParser()
        config.read(iniFile)

        # read Orisoft TMS settings
        orisoftDsn = config.get('OrisoftTMSDSN', 'dsn')
        orisoftUser = config.get('OrisoftTMSDSN', 'uid')
        orisoftPwd = config.get('OrisoftTMSDSN', 'pwd')

        # read last date processed
        dateFromPrev = config.get('History', 'datefrom')
        dateToPrev = config.get('History', 'dateto')

    except ConfigParser.NoSectionError, e:
        QMessageBox.critical(None,'Config File Error', str(e))
        app.exit(1)

    except ConfigParser.NoOptionError, e:
        QMessageBox.critical(None, 'Config File Error', str(e))
        app.exit(1)


app = QApplication(sys.argv)
try:
    readIni()
    # connection for Orisoft TMS Database
    connOriTMS = pyodbc.connect('DSN=%s; UID=%s; PWD=%s' % (orisoftDsn, orisoftUser, orisoftPwd))

except pyodbc.Error, e:
    QMessageBox.critical(None,'TMSWin Report Connection Error', str(e))
    app.exit(1)

employees = {}                          # employees daily attendance
leaveInfo = {}                          # employees leave information

form = TMSReport()
form.show()
sys.exit(app.exec_())
