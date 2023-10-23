from datetime import datetime
import pandas as pd
from openpyxl import Workbook
import re
import sqlite3

# ==============================================================================
# FUNCTIONS
# ==============================================================================

# Get Time in / Time Out
# def first(weekday):
# 		first = float(re.findall('[0-9]+(?=.*\-)', weekday)[0])
# 		return first

def first(weekday):
    if weekday == 'AF' or weekday == ' ' or weekday == '0':
        return 0.0
    else:
        first = float(re.findall('[0-9]+(?=.*\-)', weekday)[0])
        return first

def second(weekday):
    if weekday == 'AF' or weekday == ' ' or weekday == '' or weekday == '0':
        return 0.0
    else:
        second = float(re.findall('\-(.*)', weekday)[0])
        return second

# ==============================================================================
# IMPORT ROSTER TIMES AND DATES (ATTENDENTS)
# ==============================================================================
file = '../Attendant_Carwash_Roster.xls'

# Get Columns
columns = ['idx','ATTENDANTS', 'THURS', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']

# Get Times 
data = pd.read_excel(file, index_col=0, header=1, usecols=columns)
data = data.fillna(0)

# Get Dates
data_date = pd.read_excel(file, header=None, nrows=2, usecols='C:I')
data_date_ex = data_date.loc[0]

week1_dates = {}
for x in data_date_ex:
	week1_dates[x.strftime("%A")] = x.date().strftime("%d/%m/%y")

# Get week one from excel sheet
week_one_data = data.loc[0:14]
week_one = []
for x in week_one_data.to_numpy().tolist():
    if str(x[0]) != 'nan':
        if x[0] != 0:
            week_one.append(x)

# ==============================================================================
# CREATE DATABASE SQLITE
# ==============================================================================

con = sqlite3.connect("test.db")
c = con.cursor()
# c.execute("""CREATE TABLE IF NOT EXISTS roster (
#          name TEXT,
#          thur TEXT,
#         fri TEXT,
#         sat TEXT,
#          sun TEXT,
#         mon TEXT,
#          tue TEXT,
#          wed TEXT
#          )""")



# for week in week_one:
#     x = (week[0], week[1],week[2],week[3],week[4],week[5],week[6],week[7])
#     week1 = """INSERT INTO roster (
#          name,
#          thur,
#          fri,
#          sat,
#          sun,
#          mon,
#          tue,
#          wed
#          )
#          VALUES (?, ?, ?, ?, ? ,? ,?, ?)"""

#     c.execute(week1, (x))
#     con.commit()




c.execute('SELECT name FROM roster')
name_records = c.fetchall()

con.close()

# print(records[0][0])
# print(records[0][1])
# print(records[0][2])

# ==============================================================================
# TEST CODE START
# ==============================================================================
print(week1_dates)

# ==============================================================================
# TEST CODE END
# ==============================================================================

# ==============================================================================
# CREATE TIME EXCEL FILE
# ==============================================================================

wb = Workbook()
ws = wb.active
ws.title = 'Wages Calculator'

# Create Cloumn Headings
ws['A1'] = 'Name'
ws['B1'] = 'Badge Number'
ws['C1'] = 'Date'
ws['D1'] = 'Week Day'
ws['E1'] = 'Time In'
ws['F1'] = 'Time Out'
ws['G1'] = 'Hours'

# Start from the first cell. Rows and columns are zero indexed.
row = 2
col = 3

# Iterate over the data and write it out row by row
for date, dayname in week1_dates.items():
    ws.cell(row, col, value=date)
    ws.cell(row, col + 1, value=dayname)
    row += 1

# Get Time in / Time Out
t_row = 2
t_col = 1

i_row_n = 0

i_row = 0 

for record in name_records:
    con = sqlite3.connect("test.db")
    c = con.cursor()

    c.execute('SELECT * FROM roster where name=?', (record[0],))
    r = c.fetchall()
    # print(type(r[0][1]))
    

    name = r[0][0]
    thur = r[0][1]
    # fri = week[2]
    # sat = week[3]
    # sun = week[4]
    # mon = week[5]
    # tue = week[6]
    # wed = week[7]
    # sune = 0.0
    # mone = 0.0
    
    if thur == 'AF':
        thur = 0
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=thur)

    elif first(thur) == 18:
        thur_s = first(thur)
        thur_e = second(thur)
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=thur_s)
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=thur_e)
        i_row_n += 1
        i_row += 1
    else:
        thur_s = first(thur)
        thur_e = second(thur)
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=thur_s)
        ws.cell(t_row + i_row, t_col + 5, value=thur_e)
    
    i_row_n += 1
    i_row += 1
    con.close()

# Close workbook
wb.save('Wage Times.xlsx')