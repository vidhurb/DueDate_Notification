#! /usr/bin/python3 
import openpyxl
from openpyxl import Workbook
import datetime
from datetime import datetime, timedelta, date

wb_fileName = 'Forum_DB.xlsx'

#In the First
def main():
	print("Hello World")
	wb = Workbook()
	sheet = wb.active
	wb = openpyxl.load_workbook(wb_fileName)
	get_date(wb)

#Send Notifications
def send_notification(due_date, noOfDays):
	if noOfDays >= 0:
		msg = "Kindly note that your subcription expires on " + str(due_date.date()) +", Kindly renew."
	else:
		msg = "Kindly note that your subcription has already expired on " + str(due_date.date()) +", Kindly renew."
	print(msg)

def get_date(wb):
	startRow, endRow, DueDateCol = 2, 8, 3
	current = wb.active

	for i in range(startRow, endRow, 1):
		due_date = datetime.strptime(str(current.cell(row = i, column = DueDateCol).value), '%Y-%m-%d %H:%M:%S')
		today = datetime.today()
		ndays = (due_date - today).days
		print('Due Date: ', due_date)
		print('days: ', ndays)
		if ndays <= 3:
			send_notification(due_date, ndays)


if __name__=="__main__":
	main()



