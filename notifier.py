#! /usr/bin/python3 
import openpyxl
from datetime import datetime

#In the First
def main():
	print("Hello World")
	wb = openpyxl.load_workbook('Forum_DB.xlsx')
	print(type(wb))
	current = wb.get_active_sheet()
	#print(current.cell(row = 2, column = 3).value)
	get_date()
	#getting the due date
def get_date():
	wb = openpyxl.load_workbook('Forum_DB.xlsx')
	print(type(wb))
	current = wb.get_active_sheet()
	for i in range(2, 8, 1):
		temp_date = (current.cell(row = i, column = 3).value)
		print('1')
		print(temp_date)
		print(type(temp_date))
		due_date = datetime.strftime(temp_date, '%d/%m/%y') - datetime.timedelta(days=1)
		print(due_date,i)
		check_date(due_date,i)
	return due_date
if __name__=="__main__":
	main()

#checking the due date
def check_date(due_date,row_no):
	print(date.today())
	if(due_date == date.today()):
		send_notification(self, 1,row_no)
	else:
		send_notification(self, 0,row_no)

#Send Notifications
def send_notification(self, send_flag,row_no):
	print(current.cell(row= row_no, column = 4))
	print("sms sent")
	print(current.cell(row= row_no, column = 5))
	print(" Email sent")
