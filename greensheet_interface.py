import psycopg2
import openpyxl
from openpyxl.utils import get_column_letter
import sqlquery

def row_adjust(worksheet, item_list, row_start, row_finish):
	
	cur_rowrange = find_row_indexes(worksheet, row_start, row_finish)
	row_count = cur_rowrange[1] - cur_rowrange[0] - 1
	req_rows = len(item_list)
	rows_to_add = req_rows - row_count
	rows_to_delete = row_count - req_rows

	if req_rows > row_count:
		worksheet.insert_rows(cur_rowrange[1] - 2, rows_to_add)
		copy_row_to_rows(worksheet, cur_rowrange[0] + 3, cur_rowrange[0] + 3 + rows_to_add, 11)
	elif req_rows < row_count:
		worksheet.delete_rows(cur_rowrange[0] + 1, rows_to_delete)
	else:
		return

def find_row_indexes(worksheet, start_ref, finish_ref):

	for row in worksheet['A']:
		if row.value == start_ref:
			row_start = row.row
		elif row.value == finish_ref:
			row_finish = row.row
	return [row_start, row_finish]

def copy_row_to_rows(worksheet, start_row, finish_row, src_row):

	for row in range(start_row, finish_row + 1):
		copy_row(worksheet, row, src_row)

def copy_row(worksheet, tgt_row, src_row):

	tgt_range = '{}:{}'.format(tgt_row, tgt_row)

	for col in range(1, num_styled_cols(worksheet, tgt_range)):
		worksheet[formatter(col, row=tgt_row)]._style = worksheet[formatter(col, row=src_row)]._style
		worksheet[formatter(col, row=tgt_row)].value = worksheet[formatter(col=col, row=src_row)].value

def num_styled_cols(worksheet, tgt_range):

	return len([cell._style for cell in worksheet[tgt_range]])

def insert(worksheet, item_list, start_ref, finish_ref):

	r = find_row_indexes(worksheet, start_ref, finish_ref)[0] + 1
	for row in item_list:
		c = 1
		for column in row:
			worksheet.cell(row = r, column = c).value = column
			c +=1
		r +=1
	   
def reformat_sheet(worksheet):

	for row in range(1, worksheet.max_row + 1):
		worksheet.row_dimensions[row].height = 21.0

def formatter(col="", div="", row=""):

	return get_column_letter(col) + div + str(row)
	
def main():

	# Open and activate workbook to write to. 

	wb = openpyxl.load_workbook('testproject.xlsx')
	worksheet = wb.active

	# Retrieve job_id from workbook to query DB with. 

	job_id = worksheet['A1'].value

	# Retrieve lists from Postgres database. 
	cost_list = sqlquery.fetchR2([1, 59999], job_id)
	labour_list = sqlquery.fetchR2([60000, 69999], job_id)

	# Adjust number of rows in spreadworksheet to match len() of lists. 
	row_adjust(worksheet, cost_list, 'CS', 'CF')
	row_adjust(worksheet, labour_list, 'LS', 'LF')
	
	# Insert retrieved values from database into spreadworksheet. 
	insert(worksheet, cost_list, 'CS', 'CF')
	insert(worksheet, labour_list, 'LS', 'LF')

	# Reformat cells, colours etc. 
	reformat_sheet(worksheet)

	wb.save('testproject1.xlsx') 

main()
