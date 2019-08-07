import openpyxl
from openpyxl.utils import get_column_letter

def reformat_sheet(worksheet):

	for row in range(1, worksheet.max_row + 1):
		worksheet.row_dimensions[row].height = 21.0

def formatter(col="", div="", row=""):

	return get_column_letter(col) + div + str(row)