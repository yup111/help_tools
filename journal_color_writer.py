from xlrd import open_workbook
from xlwt import Workbook
import xlwt
import os

def get_row_color(wb, sheet):
    result = {}
    rows, cols = sheet.nrows, sheet.ncols
    
    for row in range(rows):
        thiscell = sheet.cell(row, cols-2)
        
        if thiscell.value == 'ERROR JES':
            result[sheet.cell(row, 1).value] = thiscell.value

        elif thiscell.value == 'C&G JES':
            if sheet.cell(row, 1).value in result:
                if result[sheet.cell(row, 1).value] == 'ERROR JES':
                    continue
            
            result[sheet.cell(row, 1).value] = thiscell.value

        elif thiscell.value == 'Student division JES':
            if sheet.cell(row, 1).value in result:
                if (result[sheet.cell(row, 1).value] == 'ERROR JES') or (result[sheet.cell(row, 1).value] == 'Student division JES'):
                    continue

            result[sheet.cell(row, 1).value] = thiscell.value

        elif thiscell.value == 'GA JES':
            if sheet.cell(row, 1).value in result:
                continue
            result[sheet.cell(row, 1).value] = thiscell.value

    return result

def out_put_result(filename, sheet, result):
    book = Workbook()
    sheet1 = book.add_sheet('Sheet 1')
    
    for row in range(rows):
        joural_id = sheet.cell(row, 1).value
        for col in range(cols):
            thiscell = sheet.cell(row, col)   
    
            if col == cols-2:
                if joural_id in result:
                    color = result[joural_id]
                else:
                    color = thiscell.value
        
                sheet1.write(row, col, color)
        
            else:
                sheet1.write(row, col, thiscell.value)
    
    book.save(filename)

if __name__ == '__main__':
    for root, dirs, files in os.walk('./input'):
        for f in files:
            path = os.path.join(root, f)

            wb = open_workbook(path)
            sheet = wb.sheet_by_name("Sheet1")
            rows, cols = sheet.nrows, sheet.ncols
            result = get_row_color(wb, sheet)
            out_put_result('./output/' + f.split('.')[0] + '.xls', sheet, result)
