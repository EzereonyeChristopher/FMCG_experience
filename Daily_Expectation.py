# We have three departments so we are expecting production and warehouse 
# production is an independent variable warehouse and logistics are dependent variable
# w + l = p 
# new_p + old_w - log = new_w 
# days_production + previous_stockount - days_logistics = new_warehouse
# you want to have a spreadshhet workbook for each day for each daily trasaction from production to warehouse

import openpyxl as xl
import pandas as pd 
import numpy as np

def spreadsheet_comp(filepath):
    df = pd.read_excel(filepath, sheet_name=None) 
    df1 = df["logistics"] 
    df2 = df["warehouse"] 
    
type(df1)    

def get_row_values(worksheet):
    """
    return data structure:
    [
        [A1, B1, C1, ...],
        [A2, B2, C2, ...],
        ...
    ]
    """
    result = []
    for i in worksheet.rows:
        row_data = []
        for j in i:
            row_data.append(j.value)
        result.append(row_data)
    return result


if __name__ == '__main__':
    # load excel file
    wb = xl.load_workbook('test1.xlsx')
    ws1 = wb.worksheets[0]
    ws2 = wb.worksheets[1]

    # get data from the first 2 worksheets
    ws1_rows = get_row_values(ws1)
    ws2_rows = get_row_values(ws2)

    # calculate and make a new sheet
    ws_new = wb.create_sheet('Done')
    # insert header
    ws_new.append(ws1_rows[0])
    for row in range(1, len(ws1_rows)):
        # do the substract cell by cell
        row_data = []
        for column, value in enumerate(ws1_rows[row]):
            if column == 0:
                # insert first column
                row_data.append(value)
            else:
                if ws1_rows[row][0] == ws2_rows[row][0]:
                    # process only when first column match
                    row_data.append(value - ws2_rows[row][column])
        ws_new.append(row_data)
    wb.save('test2.xlsx')