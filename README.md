# OPENPYXL
Python Code

#this code does not copy/paste cells from one workbook to another for me. The script does run and produces a list in the Spyder Console of <Cell u'0 - 45'.P7> etc but nothing happening inside the workbooks. Python 2.7
import openpyxl

wb1 = openpyxl.load_workbook('Source.xlsx')
wb2 = openpyxl.load_workbook('Output.xlsx')
ws1 = wb1.active
ws2 = wb2.active
for col in ws1.iter_cols(min_row=7, max_row=15, min_col=16, max_col=16):
    for cell in col:
        print(cell)
for idx, cell in enumerate(col, start = 1):
    ws2.cell(row = idx, column = 2).value = cell.value
