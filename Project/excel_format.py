from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import Alignment

# Column letters and widths
columns_att = {'A':15.00, 'B':12.33, 'C':9.7, 'D':7.8, 'E':6.3, 'F':7.7, 'G':11.11, 'H':12.70, 'I':5.30, 'J':11.30, 'K':10.30, 'L':7.60}
columns_total = {'A':11.00, 'B':16.44, 'C':16.00, 'D':21.78}

col_diff = 0.78

def cell_center(sheet):
    i = 0
    for x in range(sheet.max_row):
        sheet.cell(row=2 + i, column=2).alignment = Alignment(horizontal='center')
        sheet.cell(row=2 + i, column=3).alignment = Alignment(horizontal='center')
        sheet.cell(row=2 + i, column=4).alignment = Alignment(horizontal='center')
        i += 1

def excel_format():
    wb = load_workbook("Wage Times.xlsx")

    ws = wb['Att Week One']
    wso = wb['Att Week Two']
    wst = wb['Att Total']

    for col, size in columns_att.items():
        ws.column_dimensions[col].width = size + col_diff
        wso.column_dimensions[col].width = size + col_diff

    for col, size in columns_total.items():
        wst.column_dimensions[col].width = size + col_diff

    cell_center(wst)


    wb.save("Wage Times.xlsx")
    wb.close()