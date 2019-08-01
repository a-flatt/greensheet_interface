import psycopg2
import openpyxl
from openpyxl.utils import get_column_letter

def fetchlist(list_len):

    conn = psycopg2.connect("dbname=testdumpdb user=adamp port=5433 password=Hundredandone1")
    cur = conn.cursor()
    cur.execute("SELECT job_id FROM jobs;")
    list = cur.fetchmany(size=list_len)
    cur.close()
    conn.close()
    return [row[0] for row in list]

def adjust_cost_rows(current_sheet, cost_list):

    row_start = 'CS'
    row_finish = 'CF'
    row_adjust(current_sheet, cost_list, row_start, row_finish)
    
def adjust_labour_rows(current_sheet, labour_list):

    row_start = 'LS'
    row_finish = 'LF'
    row_adjust(current_sheet, labour_list, row_start, row_finish)

def row_adjust(current_sheet, item_list, row_start, row_finish):

    current_row_range = find_row_indexes(current_sheet, row_start, row_finish)
    row_count = current_row_range[1] - current_row_range[0] - 1
    required_rows = len(item_list)
    rows_to_add = required_rows - row_count
    rows_to_delete = row_count - required_rows

    if required_rows > row_count:
        current_sheet.insert_rows(current_row_range[1] - 2, rows_to_add)
        copy_row_to_rows(current_sheet, current_row_range[0] + 3, current_row_range[0] + 3 + rows_to_add, 11)
    elif required_rows < row_count:
        current_sheet.delete_rows(current_row_range[0], rows_to_delete)
    else:
        return

def find_row_indexes(current_sheet, start_val, finish_val):

    for row in current_sheet['A']:
        if row.value == start_val:
            row_start = row.row
        elif row.value == finish_val:
            row_finish = row.row
    print(start_val, finish_val, row_start, row_finish)
    return [row_start, row_finish]

def copy_row_to_rows(sheet, start_row, finish_row, src_row):

    for row in range(start_row, finish_row + 1):
        copy_row(sheet, row, src_row)

def copy_row(sheet, tgt_row, src_row):

    tgt_range = '{}:{}'.format(tgt_row, tgt_row)

    for col in range(1, num_styled_cols(sheet, tgt_range)):
        col_letter = get_column_letter(col)
        sheet['{}{}'.format(col_letter, tgt_row)]._style = sheet['{}{}'.format(col_letter, src_row)]._style

def num_styled_cols(sheet, tgt_range):

    return len([cell._style for cell in sheet[tgt_range]])

def insert(test_sheet, contents, start_val, finish_val):

    r = find_row_indexes(test_sheet, start_val, finish_val)[0] + 1
    for item in contents:
        test_sheet.cell(row = r, column = 1).value = item
        r +=1
       
def reformat(sheet):

    for row in range(1, sheet.max_row + 1):
        sheet.row_dimensions[row].height = 21.0

def main():

    wb = openpyxl.load_workbook('testproject.xlsx')
    sheet = wb.active

    # Retrieve lists from Postgres database. 
    cost_list = fetchlist(10)
    labour_list = fetchlist(4)

    # Adjust number of rows in spreadsheet to match len() of lists. 
    adjust_cost_rows(sheet, cost_list)
    adjust_labour_rows(sheet, labour_list)

    # Insert retrieved values from database into spreadsheet. 
    insert(sheet, cost_list, 'CS', 'CF')
    insert(sheet, labour_list, 'LS', 'LF')

    # Reformat cells, colours etc. 
    reformat(sheet)

    wb.save('testproject2.xlsx') 

main()
