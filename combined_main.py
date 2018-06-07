import sys
import combineddialog
import os
import pyodbc
import PyQt5
import datetime
import re
from PyQt5.QtWidgets import QMainWindow, QApplication, QWidget, QPushButton, QAction, QLineEdit, QMessageBox, QLCDNumber, QGridLayout, QDesktopWidget, QTabWidget
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtCore import pyqtSlot, QTimer, QTime

class PiApp(QMainWindow, combineddialog.Ui_MainWindow):

	def __init__(self, parent=None):

		super(PiApp, self).__init__(parent)
		self.setupUi(self)

		# Code for sizing the program and centering it on screen.
		qtRectangle = self.frameGeometry()
		print(qtRectangle)
		centerPoint = QDesktopWidget().availableGeometry().center()
		qtRectangle.moveCenter(centerPoint)
		self.move(qtRectangle.topLeft())

		self.tabWidget.currentChanged.connect(self.tab_change)

		# Set focus to first textbox on load, so the WOID can be scanned.
		self.txtEID.setFocus()

		# Call method to disable buttons.
		self.disable_start()

		# Call method after scan on WOID.
		self.txtWOID.returnPressed.connect(self.scan_work_order)

		# Call method after scan on Employee Clock #.
		self.txtEID.returnPressed.connect(self.scan_employee)

		# Call method to handle Start btn.
		self.btnStart.clicked.connect(self.start_labor)

		# Call method to handle Stop btn.
		self.btnStop.clicked.connect(self.stop_labor)

		# Call method to clear form.
		self.btnClear.clicked.connect(self.clear_form)

		# Define message to be passed to message box class.
		PiApp.message = ''
		# What we do after WOID is scanned.
		self.txtWOBOMID.returnPressed.connect(self.wobomid_after_update)
		# What we do after Part is scanned.

		# What we do after clock # is scanned.
		self.txtClockID.returnPressed.connect(self.clockid_after_update)

		# Event for enter btn click.
		self.btnEnter.clicked.connect(self.on_click)

		# Clear values from field. Reset focus to WOID.
		self.btnClearParts.clicked.connect(self.clearForm)

		# On load, we're setting focus to the clockID
		self.txtClockID.setFocus()

		# On load, label qty set to null.
		self.lblQty.setText('')

		# Call clock method for showing time on both tabs.
		self.clock()

		# Disable parts issuing keypad at start.
		self.disable_keypad()
		self.btnEnter.setEnabled(False)

		# Buttons to represent a numpad.
		self.btn1.clicked.connect(self.btn1Click)
		self.btn2.clicked.connect(self.btn2Click)
		self.btn3.clicked.connect(self.btn3Click)
		self.btn4.clicked.connect(self.btn4Click)
		self.btn5.clicked.connect(self.btn5Click)
		self.btn6.clicked.connect(self.btn6Click)
		self.btn7.clicked.connect(self.btn7Click)
		self.btn8.clicked.connect(self.btn8Click)
		self.btn9.clicked.connect(self.btn9Click)
		self.btn0.clicked.connect(self.btn0Click)
		self.btnClearQty.clicked.connect(self.clearqtyclick)

##################################################################################################
# These are common functions to be used on both tabs.
##################################################################################################

	# Method for connecting through ODBC to SQL Server.
	def connect(self):

		global cnxn, cursor

		cnxn = pyodbc.connect("DSN=sqlserver;UID=FSUser;PWD=1@freeman")
		cursor = cnxn.cursor()

	# Method for disconnecting ODBC to SQL server.
	def disconnect(self):
		cursor = cnxn.cursor()
		cursor.close()
		del cursor
		cnxn.close()

	def call_msg_timer(self):
		msgBox = TimerMessageBox(10, self)
		msgBox.exec_()

	def clock(self):
		timer = QTimer(self)
		timer.timeout.connect(self.showTime)
		timer.start(1000)
		self.showTime()

	def showTime(self):
		time = QTime.currentTime()
		text = time.toString('hh:mm A')

		for char in text:
			if char in "AMP":
				text = text.replace(char, "")

		if (time.second() % 2) == 0:
			text = text[:2] + ' ' + text[3:]
			
		self.lcdTime.display(text)
		self.lcdTimeParts.display(text)

		print(text)

	def tab_change(self):

		x = self.tabWidget.currentIndex()
		if x == 0:
			self.txtEID.setFocus()
		elif x == 1:
			self.txtClockID.setFocus()


##################################################################################################
# This is the end of common functions.
##################################################################################################

##################################################################################################
# This section is for the parts issuing tab.
##################################################################################################

	# Parts Issuing tab.
	def clockid_after_update(self):

		global ClockID

		ClockID = self.txtClockID.text()

		# Here we take the clock ID and return the Employee ID.
		# This is because the employee is familiar with their clock ID,
		# But for storing data we want to use their Employee ID.
		if self.get_employee_info_parts() is False:
			self.clearForm()
			PiApp.message = "You are not on the list of current users."
			self.call_msg_timer()
			return

		self.txtWOBOMID.setFocus()

	def get_employee_info_parts(self):

		global EmpID, FirstName

		self.connect()

		ClockNumeric = ClockID.isnumeric()

		if ClockNumeric == False:
			return False

		# This is the select stmt for returning EmpID from ClockID.
		cursor.execute("SELECT ID, [First Name] AS FirstName, ClockID, [Status ID] " 
			"FROM Employees "
			"WHERE [Status ID] = 1 AND ClockID = ?", (ClockID))
		row = cursor.fetchone()

		if row:
			FirstName = row.FirstName
			EmpID = row.ID
		else:
			return False

		self.disconnect()

	# Function for moving to Part textbox after WOID has been scanned.
	# Return WOID information so the user can confirm.
	def wobomid_after_update(self):

		global WOID, ProductID, ProductCode, WOBOMID

		WOBOMID = self.txtWOBOMID.text()

		WOBOMIDNumeric = WOBOMID.isnumeric()

		if WOBOMIDNumeric == False:
			PiApp.message = "WOBOMID must be numeric."
			self.clearForm()
			self.clear_form()
			self.call_msg_timer()
			return

		self.connect()

		cursor.execute("SELECT [Work Order BOM].[ID] AS WOBOMID, [Work Order].[WOID], [Work Order].[WO Status], "
			"[Products].[Product Code] AS WOName, [Products_1].[Product Code] AS MaterialName, [Products_1].[ID] AS ProductID "
			"FROM [Work Order] INNER JOIN [Products] ON [Work Order].[ProductID] = [Products].[ID] INNER JOIN "
			"[Work Order BOM] ON [Work Order].[WOID] = [Work Order BOM].[WOID] INNER JOIN Products AS Products_1 "
			"ON [Work Order BOM].[ProductID] = [Products_1].[ID] "
			"WHERE ([Work Order].[WO Status] = 2 OR [Work Order].[WO Status] = 3) AND ([Work Order BOM].[ID] = ?)", (WOBOMID))

		row = cursor.fetchone()

		if row:
			self.lblWOIDReturn.setText(row.WOName)
			self.lblPartReturn.setText(row.MaterialName)
			ProductID = int(row.ProductID)
			WOID = int(row.WOID)
			ProductCode = row.MaterialName
		else:
			self.txtWOBOMID.clear()
			self.txtWOBOMID.setFocus()
			self.lblWOIDReturn.setText('')
			return

		self.lblQty.setFocus()
		self.enable_keypad()
		self.btnEnter.setEnabled(True)

		self.disconnect()

	# Click event for entering data to server.
	# First we validate data. This will become more extensive based on parameters.
	# Then we create a connection string. This will need to be more secure.
	# Following this will load variables from controls, then create INSERT statement.
	# Lastly, we reset controls to form load values.

	@pyqtSlot()

	def on_click(self):

		global rv

		if self.validate() == False:
			return

		#Quantity = self.slQty.value()
		#Quantity = int(self.lblQty.text())

		self.connect()

		# SQL string for stored procedure with a return value.
		sql = """\
		DECLARE @outRV int;
		EXEC @outRV = [dbo].[sprocFIFOIssueInventory] @pProductID = ?, @pQtyToIssue = ?, @pWOBOMID = ?, @pEmpID = ?;
		SELECT @outRV AS RV;
		"""
		# Parameters to feed into the stored procedure.
		params = (ProductID, Quantity, WOBOMID, EmpID, )
		try:
			cursor.execute(sql, params)
		except Exception as e:
			print(e)

		# Fetch the return value (either 1, 2, or 3)
		return_value = cursor.fetchone()

		# Turn it into a string.
		return_value = str(return_value)

		# Get rid of extra stuff and put it into a list.
		rv = [int(s) for s in re.findall(r'\d+', return_value)]

		# Pull it out of the list.
		rv = rv[0]

		# Commit the SPROC. Not doing this locks up the tables.
		cnxn.commit()

		# Clear form.
		self.clearForm()

		# Evaluate RV and determine message to be returned.
		self.check_return_value()

		# Close cursor and disconnect.
		self.disconnect()

	# Evaluate for the return value. These numbers correspond to the RETURNs on the sproc.
	# Create appropriate messageboxes. If it failed it didn't go through already.
	# We don't have to worry about RETURN here, or exiting early.
	def check_return_value(self):

		if rv == 1:
			PiApp.message = "Thanks " + FirstName + "! You issued " + str(Quantity) + \
				" units of " + ProductCode + " to Work Order: " + str(WOID)
			self.call_msg_timer()
		elif rv == 2:
			PiApp.message = "Unable to issue. Inventory shows none remaining. Please contact the tool-room."
			self.call_msg_timer()
		elif rv == 3:
			PiApp.message = "This item is not part of the Work Order's BOM. Please contact the tool-room."
			self.call_msg_timer()

	# Will call enable keypad after wobom id has been successfully scanned.
	def enable_keypad(self):

		self.btn1.setEnabled(True)
		self.btn2.setEnabled(True)
		self.btn3.setEnabled(True)
		self.btn4.setEnabled(True)
		self.btn5.setEnabled(True)
		self.btn6.setEnabled(True)
		self.btn7.setEnabled(True)
		self.btn8.setEnabled(True)
		self.btn9.setEnabled(True)
		self.btn0.setEnabled(True)
		self.btnClearQty.setEnabled(True)

	# Disable parts issue keypad after entry of material and on load.
	def disable_keypad(self):

		self.btn1.setEnabled(False)
		self.btn2.setEnabled(False)
		self.btn3.setEnabled(False)
		self.btn4.setEnabled(False)
		self.btn5.setEnabled(False)
		self.btn6.setEnabled(False)
		self.btn7.setEnabled(False)
		self.btn8.setEnabled(False)
		self.btn9.setEnabled(False)
		self.btn0.setEnabled(False)
		self.btnClearQty.setEnabled(False)

	def btn1Click(self):

		QtyClicked = '1'
		self.show_lbl_qty(QtyClicked)

	def btn2Click(self):

		QtyClicked = '2'
		self.show_lbl_qty(QtyClicked)

	def btn3Click(self):

		QtyClicked = '3'
		self.show_lbl_qty(QtyClicked)

	def btn4Click(self):

		QtyClicked = '4'
		self.show_lbl_qty(QtyClicked)

	def btn5Click(self):

		QtyClicked = '5'
		self.show_lbl_qty(QtyClicked)

	def btn6Click(self):

		QtyClicked = '6'
		self.show_lbl_qty(QtyClicked)

	def btn7Click(self):

		QtyClicked = '7'
		self.show_lbl_qty(QtyClicked)

	def btn8Click(self):

		QtyClicked = '8'
		self.show_lbl_qty(QtyClicked)

	def btn9Click(self):

		QtyClicked = '9'
		self.show_lbl_qty(QtyClicked)

	def btn0Click(self):

		QtyClicked = '0'
		self.show_lbl_qty(str(QtyClicked))

	# Function for showing label value on change.
	def show_lbl_qty(self, QtyClicked):
		global Quantity
		newNumber = self.lblQty.text() + str(QtyClicked)
		self.lblQty.setText(newNumber)
		Quantity = int(self.lblQty.text())

	def clearqtyclick(self):
		self.lblQty.setText('')

	# Function for validating data.
	def validate(self):

		if self.txtClockID.text() == "":
			PiApp.message = "Error: Please re-scan employee clock number."
			self.call_msg_timer()
			self.clearForm()
			self.txtClockID.setFocus()
			return False
		elif self.txtWOBOMID.text() == "":
			PiApp.message = "Error: Please re-scan WOBOMID."
			self.call_msg_timer()
			self.txtWOBOMID.clear()
			self.txtWOBOMID.setFocus()
			return False
		elif self.lblQty.text() == "":
			PiApp.message = "Error: Please select a quantity, 30 or fewer."
			self.call_msg_timer()
			self.lblQty.clear()
			self.lblQty.setFocus()
			return False
		elif Quantity > 30 or Quantity is None:
			PiApp.message = "Error: Please select a quantity, 30 or fewer."
			self.call_msg_timer()
			self.lblQty.clear()
			self.lblQty.setFocus()
			return False

	# Function for clearing form and resetting focus to on load parameters.
	def clearForm(self):
		self.txtWOBOMID.clear()
		self.txtClockID.clear()
		self.txtClockID.setFocus()
		self.lblWOIDReturn.setText('')
		self.lblPartReturn.setText('')
		self.lblQty.setText('')
		self.disable_keypad()
		self.btnEnter.setEnabled(False)

##################################################################################################
# End of the parts issuing tab.
##################################################################################################


##################################################################################################
# This section is for the labor entry tab.
##################################################################################################

	# Method for scanning WOID.
	# Labor tracking tab.
	def scan_work_order(self):

		global WOID, WOName

		WOID = self.txtWOID.text()

		WOIDNumeric = WOID.isnumeric()

		if WOIDNumeric == False:
			PiApp.message = "WOID must be numeric."
			self.call_msg_timer()
			self.txtWOID.clear()
			self.txtWOID.setFocus()
			return

		self.connect()

		cursor.execute("SELECT [Work Order].[WOID], [Work Order].[WO Status], "
					   "[Products].[Product Code] AS Name "
					   "FROM [Work Order] INNER JOIN [Products] ON "
					   "[Work Order].[ProductID] = [Products].[ID] "
					   "WHERE ([Work Order].[WO Status] = 2 OR [Work Order].[WO Status] = 3) "
					   "AND ([Work Order].[WOID] = ?)", (WOID))

		row = cursor.fetchone()
		# If we have data. Any Work Order that is RELASED or WIP.
		if row:

			self.btnStart.setFocus()
			WOName = row.Name

		else:
			# If the above criteria is not met.
			self.txtWOID.clear()
			self.txtWOID.setFocus()
			return

	# Method for scanning employee Clock #.
	# Labor tracking tab.
	def scan_employee(self):

		global ClockID, TimeNow, RecordID, TimeIn, LastLaborCode, CurrentWOID

		TimeNow = datetime.datetime.now()
		ClockID = self.txtEID.text()

		if self.get_employee_info_labor() is False:
			self.clear_form()
			PiApp.message = "You must be a current employee on the labor list to record labor."
			self.call_msg_timer()
			return

		self.connect()
		cursor.execute("SELECT ID, EmpClockID, WOID, TimeIn, TimeOut, LaborCode "
					   "FROM LaborTest WHERE EmpClockID = ? "
					   "ORDER BY ID DESC", (EmpID))

		row = cursor.fetchone()

		if row:
			# If the employee has a record for labor. All will have this except new hire on first day.
			RecordID = row.ID
			# Checking for a completed record, i.e. both Time In and Time Out fields populated.
			if (row.TimeIn is not None and row.TimeOut is not None) \
			or (row.TimeIn is None and row.TimeOut is None):
				# If so, they're starting new work. So we call function that enables starting.
				LastLaborCode = row.LaborCode
				self.auto_select_labor_code()
				self.enable_start()
				self.txtWOID.setFocus()
			else:
				# This means that they're currently working, so they'll only be able to stop work.
				TimeIn = row.TimeIn
				self.get_work_time()
				CurrentWOID = row.WOID
				self.get_woid_name()
				self.btnStop.setEnabled(True)
				self.btnStop.setFocus()

		else:
			# This is to handle new hires on their first day.
			# We could remove this if we automatically add an inital
			# entry with all fields completed.
			self.enable_start()
			self.txtWOID.setFocus()

	def get_employee_info_labor(self):

		global FirstName, EmpID

		ClockNumeric = ClockID.isnumeric()

		if ClockNumeric == False:
			return False

		self.connect()
		cursor.execute("SELECT [ID], [First Name] AS FirstName, ClockID, [Labor List], [Status ID] " 
					   "FROM Employees "
					   "WHERE [Labor List] = 1 AND [Status ID] = 1 AND ClockID = ?", (ClockID))

		row = cursor.fetchone()

		if row:
			FirstName = row.FirstName
			EmpID = row.ID
		else:
			return False

	# Method for returning labor code from radio button selection.
	def get_labor_code(self):

		global LaborCode

		if self.rbFab.isChecked():
			LaborCode = 1
		elif self.rbWeld.isChecked():
			LaborCode = 2
		elif self.rbAssembly.isChecked():
			LaborCode = 3
		elif self.rbPaint.isChecked():
			LaborCode = 4
		elif self.rbFinal.isChecked():
			LaborCode = 5
		elif self.rbIndirect.isChecked():
			LaborCode = 6
		elif self.rbElectric.isChecked():
			LaborCode = 10
		elif self.rbShipReceive.isChecked():
			LaborCode = 11
		elif self.rbMaterialHandling.isChecked():
			LaborCode = 13
		elif self.rbLab.isChecked():
			LaborCode = 20
		else:
			return False

	# Method for starting labor.
	def start_labor(self):

		if self.validate_form() is False:
			return
			
		self.connect()
		cursor.execute("INSERT INTO LaborTest (EmpClockID, WOID, TimeIn, LaborCode) "
							   "VALUES (?, ?, ?, ?)", (EmpID, WOID, TimeNow, LaborCode))
		cnxn.commit()

		self.get_employee_info_labor()

		PiApp.message = "Thanks " + FirstName + "! You've begun working on " + WOName + " - " + WOID + "."

		self.clear_form()
		self.call_msg_timer()

	# Method for stopping labor, fewer controls / options.
	def stop_labor(self):

		self.connect()
		cursor.execute("UPDATE LaborTest SET TimeOut = ? WHERE ID = ?", (TimeNow, RecordID))
		cnxn.commit()

		self.get_employee_info_labor()

		PiApp.message = ("Thanks " + FirstName +
								   "! You've stopped working on " + CurrentName + 
								   " - " + str(CurrentWOID) +
								   ". Your work segment was: " + tDelta + ".")

		self.clear_form()
		self.call_msg_timer()

	# Form validation to confirm there is data, meant to handle potential errors.
	def validate_form(self):

		if self.txtWOID.text() == "":
			self.txtWOID.setFocus()
			return False

		if self.txtEID.text() == "":
			self.txtEID.setFocus()
			return False

		if self.get_labor_code() == False:
			return False

	def get_woid_name(self):

		global CurrentName
		self.connect()
		cursor.execute("SELECT [Work Order].WOID, Products.[Product Code] AS Name FROM "
					   "[Work Order] INNER JOIN Products ON "
					   "[Work Order].ProductID = Products.ID "
					   "WHERE WOID = ?", (CurrentWOID))
		row = cursor.fetchone()

		if row:
			CurrentName = row.Name
		else:
			return

	# Method for enabling radio buttons and start button.
	def enable_start(self):

		self.rbFab.setEnabled(True)
		self.rbWeld.setEnabled(True)
		self.rbAssembly.setEnabled(True)
		self.rbPaint.setEnabled(True)
		self.rbFinal.setEnabled(True)
		self.rbIndirect.setEnabled(True)
		self.rbElectric.setEnabled(True)
		self.rbShipReceive.setEnabled(True)
		self.rbMaterialHandling.setEnabled(True)
		self.rbLab.setEnabled(True)
		self.btnStart.setEnabled(True)

	# On load method and after update on start/stop so that no
	# buttons can be pressed.
	def disable_start(self):

		self.rbFab.setEnabled(False)
		self.rbWeld.setEnabled(False)
		self.rbAssembly.setEnabled(False)
		self.rbPaint.setEnabled(False)
		self.rbFinal.setEnabled(False)
		self.rbIndirect.setEnabled(False)
		self.rbElectric.setEnabled(False)
		self.rbShipReceive.setEnabled(False)
		self.rbMaterialHandling.setEnabled(False)
		self.rbLab.setEnabled(False)
		self.btnStart.setEnabled(False)
		self.btnStop.setEnabled(False)

	# Clear form method after updating or inserting a record.
	def clear_form(self):

		self.txtWOID.clear()
		self.txtEID.clear()
		self.txtEID.setFocus()
		self.btnStart.setEnabled(False)
		self.btnStop.setEnabled(False)
		self.disable_start()

	# Return time spent on work segment to employee.
	def get_work_time(self):

		global tDelta

		FMT = '%H:%M:%S'
		strTimeIn = TimeIn.strftime(FMT)
		strTimeNow = TimeNow.strftime(FMT)

		tDelta = str(datetime.datetime.strptime(strTimeNow, FMT) - \
		datetime.datetime.strptime(strTimeIn, FMT))

	def auto_select_labor_code(self):

		if LastLaborCode == 1:
			self.rbFab.setChecked(True)
		elif LastLaborCode == 2:
			self.rbWeld.setChecked(True)
		elif LastLaborCode == 3:
			self.rbAssembly.setChecked(True)
		elif LastLaborCode == 4:
			self.rbPaint.setChecked(True)
		elif LastLaborCode == 5:
			self.rbFinal.setChecked(True)
		elif LastLaborCode == 6:
			self.rbIndirect.setChecked(True)
		elif LastLaborCode == 10:
			self.rbElectric.setChecked(True)
		elif LastLaborCode == 11:
			self.rbShipReceive.setChecked(True)
		elif LastLaborCode == 13:
			self.rbMaterialHandling.setChecked(True)
		elif LastLaborCode == 20:
			self.rbLab.setChecked(True)
		else:
			return

##################################################################################################
# End of the labor entry tab.
##################################################################################################

# Class for automatic close on messagebox after start and stop work.
# This closes after 10 seconds but can be closed manually by scanning.
class TimerMessageBox(QMessageBox):

	def __init__(self, timeout=10, parent=None):
		super(TimerMessageBox, self).__init__(parent)
		self.setWindowTitle('Clock Successful')
		self.time_to_wait = timeout
		font = QFont()
		font.setPointSize(24)
		self.setFont(font)
		self.setText(PiApp.message)
		self.setStandardButtons(QMessageBox.Ok)
		self.timer = QTimer(self)
		self.timer.setInterval(1000)
		self.timer.timeout.connect(self.change_timer)
		self.change_timer()
		self.timer.start()

	def change_timer(self):
		self.time_to_wait -= 1
		if self.time_to_wait <= 0:
			self.close()

	def closeEvent(self, event):
		self.timer.stop()
		event.accept()

def main():

	app = QApplication(sys.argv)
	form = PiApp()
	form.show()
	app.exec_()

if __name__=='__main__':

	main()