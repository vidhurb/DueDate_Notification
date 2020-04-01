#! /usr/bin/python3 
import openpyxl
from datetime import date
#In the First
wb = openpyxl.load_workbook('Forum_DB.xlsx')
print(type(wb))
current = wb.get_active_sheet()
print(current.cell(row = 2, column = 3).value)

#getting the due date
def get_date():
	for i in range(2, 8, 1):
		temp_date = (current.cell(row = i, column = 3).value)
		print(temp_date)
		due_date = datetime.datetime.strftime(temp_date, '%d/%m/%y') - datetime.timedelta(days=1)
		print(due_date,i)
		check_date(due_date,i)
	return due_date

#checking the due date
def check_date(due_date,row_no):
	print(date.today())
	if(due_date == date.today()):
		send_notification(self,send_flag = 1,row_no)
	else:
		send_notification(self,send_flag = 0,row_no)

#Send Notifications
def send_notification(self, send_flag,row_no):
	print(current.cell(row= row_no, column = 4))
	print("sms sent")
	print(current.cell(row= row_no, column = 5))
	print(" Email sent")

