#! /usr/bin/python3 
import openpyxl
 #Create a file called config.py and add acesss token in thew same dir, it will be ignored by git for security
import config
import datetime
from openpyxl import Workbook
from datetime import datetime, timedelta, date

wb_fileName = 'Forum_DB.xlsx'

#In the First
def main():
	wb = Workbook()
	sheet = wb.active
	wb = openpyxl.load_workbook(wb_fileName)
	get_date(wb)
	print('From Config: ', config.access_token)

#Send Notifications
def send_notification(due_date, noOfDays, phoneNum):
	if noOfDays >= 0:
		msg = "Kindly note that your subcription expires on " + str(due_date.date()) +", Kindly renew." + "\nPhone: " + phoneNum
	else:
		msg = "Kindly note that your subcription has already expired on " + str(due_date.date()) +", Kindly renew." + "\nPhone: " + phoneNum
	print(msg)

def get_date(wb):
	startRow, endRow, DueDateCol, phoneNumCol = 2, 8, 3, 4
	current = wb.active

	for i in range(startRow, endRow, 1):
		due_date = datetime.strptime(str(current.cell(row = i, column = DueDateCol).value), '%Y-%m-%d %H:%M:%S')
		phoneNum = str(current.cell(row = i, column = phoneNumCol).value)
		today = datetime.today()
		ndays = (due_date - today).days
		#print('Due Date: ', due_date)
		#print('days: ', ndays)
		if ndays <= 3:
			send_notification(due_date, ndays, phoneNum)

if __name__=="__main__":
	main()



