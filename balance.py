from xlwt import *
from xlrd import *
from xlutils.copy import copy
from datetime import *

filename = "balance.xls"

def init_new_file():
	global filename
	print "Enter New Balance : "
	Balance = input(">")

	write_book = Workbook()
	total_bal_write = write_book.add_sheet("total_bal")
	pocket_bal_write = write_book.add_sheet("pocket_bal")
	total_bal_write.write(0,0,"Lost")
	total_bal_write.write(0,1,"Gain")
	total_bal_write.write(0,2,"Comment")
	total_bal_write.write(0,3,"Date")
	total_bal_write.write(0,4,"Balance")

	total_bal_write.write(1,1,Balance)
	total_bal_write.write(1,2,"Starting balance")
	total_bal_write.write(1,3,datetime.today(),easyxf(num_format_str='DD-MM-YYYY'))
	total_bal_write.write(1,4,Balance)

	write_book.save(filename)

def update_balance():
	read_book = open_workbook(filename,formatting_info=True)
	total_bal_read = read_book.sheet_by_index(0)
	pocket_bal_read = read_book.sheet_by_index(1)
	write_book = copy(read_book)
	total_bal_write = write_book.get_sheet(0)
	pocket_bal_write = write_book.get_sheet(1)

	write_book.save(filename)

init_new_file()



