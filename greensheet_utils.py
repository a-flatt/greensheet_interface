import openpyxl
from openpyxl.utils import get_column_letter

def reformat_sheet(worksheet):

	for row in range(1, worksheet.max_row + 1):
		worksheet.row_dimensions[row].height = 21.0

def formatter(col="", row=""):

	return get_column_letter(col) + str(row)

def num_styled_cols(worksheet, tgt_range):

	return len([cell._style for cell in worksheet[tgt_range]])
