from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.worksheet.dimensions import ColumnDimension

columns_att = {'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L'}
columns_total = ['A', 'B', 'C', 'D']
# def excel_format():

wb = load_workbook("Wage Times.xlsx")
# Format Att Week One Sheet
ws = wb['Att Week One']
wso = wb['Att Week Two']
wst = wb['Att Total']

for x in columns_att:
    ws.column_dimensions[x].auto_size = True
     

# for x in columns_att:
#     wso.column_dimensions[x] = ColumnDimension(wso, auto_size=True)

# for x in columns_total:
#     wst.column_dimensions[x] = ColumnDimension(wst, auto_size=True)

wb.save("Wage Times.xlsx")
wb.close()