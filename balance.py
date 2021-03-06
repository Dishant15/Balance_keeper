#!/usr/local/bin/python

from xlwt import *
from xlrd import *
from xlutils.copy import copy
from datetime import *

filename = "balance.xls"

def init_new_file():
	global filename
	print "Enter New Balance : "
	Balance = input(">")
	#print "Enter Pocket Balance : "
	pock_bal = 0 #input(">")
	write_book = Workbook()
	total_bal_write = write_book.add_sheet("total_bal")
	pocket_bal_write = write_book.add_sheet("pocket_bal")
	total_bal_write.write(0,0,"Gain",easyxf('pattern: pattern solid, fore_colour yellow;'))
	total_bal_write.write(0,1,"Lost",easyxf('pattern: pattern solid, fore_colour yellow;'))
	total_bal_write.write(0,2,"Comment",easyxf('pattern: pattern solid, fore_colour yellow;'))
	total_bal_write.write(0,3,"Date",easyxf('pattern: pattern solid, fore_colour yellow;'))
	total_bal_write.write(0,4,"Curr Balance",easyxf('pattern: pattern solid, fore_colour yellow;'))
	pocket_bal_write.write(0,0,"Gain",easyxf('pattern: pattern solid, fore_colour yellow;'))
	pocket_bal_write.write(0,1,"Lost",easyxf('pattern: pattern solid, fore_colour yellow;'))
	pocket_bal_write.write(0,2,"Comment",easyxf('pattern: pattern solid, fore_colour yellow;'))
	pocket_bal_write.write(0,3,"Date",easyxf('pattern: pattern solid, fore_colour yellow;'))
	pocket_bal_write.write(0,4,"Curr Balance",easyxf('pattern: pattern solid, fore_colour yellow;'))
	total_bal_write.col(2).width = 20000
	#pocket_bal_write.col(2).width = 65536

	total_bal_write.write(1,0,Balance)
	total_bal_write.write(1,2,"Starting balance")
	total_bal_write.write(1,3,datetime.today(),easyxf(num_format_str='DD-MM-YYYY'))
	total_bal_write.write(1,4,Balance)
	pocket_bal_write.write(1,0,pock_bal)
	pocket_bal_write.write(1,2,"Starting balance")
	pocket_bal_write.write(1,3,datetime.today(),easyxf(num_format_str='DD-MM-YYYY'))
	pocket_bal_write.write(1,4,pock_bal)

	write_book.save(filename)

def update_balance(x, amount, comment):
	#open a readable xls book object
	read_book = open_workbook(filename,formatting_info=True)
	total_bal_read = read_book.sheet_by_index(0)
	pocket_bal_read = read_book.sheet_by_index(1)
	#copy the old object into a new writable object so we dont loose old data
	write_book = copy(read_book)
	total_bal_write = write_book.get_sheet(0)
	pocket_bal_write = write_book.get_sheet(1)
	#code to update Balance data
	entry_len = len(total_bal_read.col_values(3))
	curr_bal = total_bal_read.col_values(4)[-1]
	#write gain or lose in file
	total_bal_write.write(entry_len,x,amount)
	total_bal_write.write(entry_len,2,comment)  #Write comment
	total_bal_write.write(entry_len,3,datetime.today(),easyxf(num_format_str='DD-MM-YYYY'))  #Write date
	#calculate balance and store it
	if x == 1: #if lose entry
		curr_bal = curr_bal - amount
	else:      #if Gain entry
		curr_bal = curr_bal + amount
	total_bal_write.write(entry_len,4,curr_bal)
	#Save new updated data
	write_book.save(filename)

#User input options
print "1 : Expenditure Entry"
print "2 : Income Entry"
print "3 : Creat new xls file"
print "Enything else to exit"
x = 0
entry = raw_input(">")
if entry == '1':
	x = 1
elif entry == '2':
	x = 2
elif entry == '3':
	init_new_file()
	exit()
else:
	print "Wrong Entry"
	exit()

while True:
	print "Type in amount :"
	amount = input(">")
	print "Comment :"
	comment = raw_input(">")
	update_balance(x, amount , comment)
	print "More Expenditure?"
	n = raw_input("(y/n) >")
	if n == 'n':
		break

