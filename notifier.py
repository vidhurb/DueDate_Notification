#! /usr/bin/python3 
import openpyxl
from openpyxl import Workbook
from datetime import datetime

wb_fileName = 'Forum_DB.xlsx'
#In the First
def main():
	print("Hello World")

	#print(current.cell(row = 2, column = 3).value)
	get_date()
	#getting the due date

def get_date():
	wb = openpyxl.load_workbook(wb_fileName)
	print('WB type:',type(wb))
	current = wb.active
	
	for i in range(2, 8, 1):
		temp_date = (current.cell(row = i, column = 3).value)
		print('temp date:',temp_date,'temp date',type(temp_date))
		due_date = datetime.strptime(str(temp_date), '%y/%m/%d %H:%M:%S'%y/%m/%d %H:%M:%S'') - datetime.timedelta(days=1)
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
