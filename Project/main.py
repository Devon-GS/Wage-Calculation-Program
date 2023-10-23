from datetime import datetime
import pandas as pd
from openpyxl import Workbook
from openpyxl import load_workbook
import re
import sqlite3

# ==============================================================================
# FUNCTIONS
# ==============================================================================

# Get Time in / Time Out
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

weekone_dates = {}
for x in data_date_ex:
	weekone_dates[x.strftime("%A")] = x.date().strftime("%d/%m/%y")

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

    # c.execute(week1, (x))
    # con.commit()

c.execute('SELECT name FROM roster')
name_records = c.fetchall()

con.close()

# ==============================================================================
# TEST CODE START
# ==============================================================================



# ==============================================================================
# TEST CODE END
# ==============================================================================

# ==============================================================================
# CREATE TIME EXCEL FILE
# ==============================================================================

wb = Workbook()
ws = wb.active
ws.title = 'Wages Calculator'

# Create Column Headings
ws['A1'] = 'Name'
ws['B1'] = 'Badge Number'
ws['C1'] = 'Date'
ws['D1'] = 'Week Day'
ws['E1'] = 'Time In'
ws['F1'] = 'Time Out'
ws['G1'] = 'Hours'

# Get Time in / Time Out
# Start row and col
t_row = 2
t_col = 1
# Date and day start row and col
d_row = 2
d_col = 3
# Date and day continue
i_row_d = 0
# Name contine row
i_row_n = 0
# Row continue
i_row = 0 

# Loop through database and save to excel sheet
for record in name_records:
    con = sqlite3.connect("test.db")
    c = con.cursor()

    # Grab data from database using name of person
    c.execute('SELECT * FROM roster where name=?', (record[0],))
    r = c.fetchall()
    
    name = r[0][0]
    thur = r[0][1]
    fri = r[0][2]
    sat = r[0][3]
    sun = r[0][4]
    mon = r[0][5]
    tue = r[0][6]
    wed = r[0][7]

    thursday = weekone_dates['Thursday']
    friday = weekone_dates['Friday']
    saturday = weekone_dates['Saturday']
    sunday = weekone_dates['Sunday']
    monday = weekone_dates['Monday']
    tuesday = weekone_dates['Tuesday']
    wednesday = weekone_dates['Wednesday']
    
    # Thursday
    if thur == 'AF':
        thur = 0
        # Add name to column A
        ws.cell(t_row + i_row_n, t_col, value=name)
        # Add roster time to column E and F
        ws.cell(t_row + i_row, t_col + 4, value=thur)
        ws.cell(t_row + i_row, t_col + 5, value=thur)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Thursday')
        ws.cell(d_row + i_row_d, d_col + 1, value=thursday)
    
    elif first(thur) == 18:
        thur_s = first(thur)
        thur_e = second(thur)
        # Add name to column A and start time to column E
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=thur_s)
        # Add name to column A and shift row down by one and
        # Add end time to column F
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=thur_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Thursday')
        ws.cell(d_row + i_row_d, d_col + 1, value=thursday)
        # Add dates to columns C and D and shift row down by 1
        ws.cell(d_row + i_row_d + 1, d_col, value='Thursday')
        ws.cell(d_row + i_row_d + 1, d_col + 1, value=thursday)
        # Add one to shift rows down for next days info
        i_row_n += 1
        i_row += 1
        i_row_d += 1
    else:
        thur_s = first(thur)
        thur_e = second(thur)
        # Add name to column A and roster times to column E and F
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=thur_s)
        ws.cell(t_row + i_row, t_col + 5, value=thur_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Thursday')
        ws.cell(d_row + i_row_d, d_col + 1, value=thursday)

    # Add one to shift rows down for next dayd info
    i_row_n += 1
    i_row += 1
    i_row_d += 1 
    
    # Friday
    if fri == 'AF':
        fri = 0
        # Add name to column A
        ws.cell(t_row + i_row_n, t_col, value=name)
        # Add roster time to column E and F
        ws.cell(t_row + i_row, t_col + 4, value=fri)
        ws.cell(t_row + i_row, t_col + 5, value=fri)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Friday')
        ws.cell(d_row + i_row_d, d_col + 1, value=friday)
    
    elif first(fri) == 18:
        fri_s = first(fri)
        fri_e = second(fri)
        # Add name to column A and start time to column E
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=fri_s)
        # Add name to column A and shift row down by one and
        # Add end time to column F
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=fri_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Friday')
        ws.cell(d_row + i_row_d, d_col + 1, value=friday)
        # Add dates to columns C and D and shift row down by 1
        ws.cell(d_row + i_row_d + 1, d_col, value='Friday')
        ws.cell(d_row + i_row_d + 1, d_col + 1, value=friday)
        # Add one to shift rows down for next days info
        i_row_n += 1
        i_row += 1
        i_row_d += 1
    else:
        fri_s = first(fri)
        fri_e = second(fri)
        # Add name to column A and roster times to column E and F
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=fri_s)
        ws.cell(t_row + i_row, t_col + 5, value=fri_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Friday')
        ws.cell(d_row + i_row_d, d_col + 1, value=friday)

    # Add one to shift rows down for next dayd info
    i_row_n += 1
    i_row += 1
    i_row_d += 1 

    # Saturday
    if sat == 'AF':
        sat = 0
        # Add name to column A
        ws.cell(t_row + i_row_n, t_col, value=name)
        # Add roster time to column E and F
        ws.cell(t_row + i_row, t_col + 4, value=sat)
        ws.cell(t_row + i_row, t_col + 5, value=sat)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Saturday')
        ws.cell(d_row + i_row_d, d_col + 1, value=saturday)
    
    elif first(sat) == 18:
        sat_s = first(sat)
        sat_e = second(sat)
        # Add name to column A and start time to column E
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=sat_s)
        # Add name to column A and shift row down by one and
        # Add end time to column F
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=sat_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Saturday')
        ws.cell(d_row + i_row_d, d_col + 1, value=saturday)
        # Add dates to columns C and D and shift row down by 1
        ws.cell(d_row + i_row_d + 1, d_col, value='Saturday')
        ws.cell(d_row + i_row_d + 1, d_col + 1, value=saturday)
        # Add one to shift rows down for next days info
        i_row_n += 1
        i_row += 1
        i_row_d += 1
    else:
        sat_s = first(sat)
        sat_e = second(sat)
        # Add name to column A and roster times to column E and F
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=sat_s)
        ws.cell(t_row + i_row, t_col + 5, value=sat_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Saturday')
        ws.cell(d_row + i_row_d, d_col + 1, value=saturday)
    
    # Add one to shift rows down for next dayd info
    i_row_n += 1
    i_row += 1
    i_row_d += 1 

    # Sunday
    if sun == 'AF':
        sun = 0
        # Add name to column A
        ws.cell(t_row + i_row_n, t_col, value=name)
        # Add roster time to column E and F
        ws.cell(t_row + i_row, t_col + 4, value=sun)
        ws.cell(t_row + i_row, t_col + 5, value=sun)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Sunday')
        ws.cell(d_row + i_row_d, d_col + 1, value=sunday)
    
    elif first(sun) == 18:
        sun_s = first(sun)
        sun_e = second(sun)
        # Add name to column A and start time to column E
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=sun_s)
        # Add name to column A and shift row down by one and
        # Add end time to column F
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=sun_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Sunday')
        ws.cell(d_row + i_row_d, d_col + 1, value=sunday)
        # Add dates to columns C and D and shift row down by 1
        ws.cell(d_row + i_row_d + 1, d_col, value='Sunday')
        ws.cell(d_row + i_row_d + 1, d_col + 1, value=sunday)
        # Add one to shift rows down for next days info
        i_row_n += 1
        i_row += 1
        i_row_d += 1
    else:
        sun_s = first(sun)
        sun_e = second(sun)
        # Add name to column A and roster times to column E and F
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=sun_s)
        ws.cell(t_row + i_row, t_col + 5, value=sun_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Sunday')
        ws.cell(d_row + i_row_d, d_col + 1, value=sunday)
    
    # Add one to shift rows down for next dayd info
    i_row_n += 1
    i_row += 1
    i_row_d += 1

    # Monday
    if mon == 'AF':
        mon = 0
        # Add name to column A
        ws.cell(t_row + i_row_n, t_col, value=name)
        # Add roster time to column E and F
        ws.cell(t_row + i_row, t_col + 4, value=mon)
        ws.cell(t_row + i_row, t_col + 5, value=mon)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Monday')
        ws.cell(d_row + i_row_d, d_col + 1, value=monday)
    
    elif first(mon) == 18:
        mon_s = first(mon)
        mon_e = second(mon)
        # Add name to column A and start time to column E
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=mon_s)
        # Add name to column A and shift row down by one and
        # Add end time to column F
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=mon_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Monday')
        ws.cell(d_row + i_row_d, d_col + 1, value=monday)
        # Add dates to columns C and D and shift row down by 1
        ws.cell(d_row + i_row_d + 1, d_col, value='Monday')
        ws.cell(d_row + i_row_d + 1, d_col + 1, value=monday)
        # Add one to shift rows down for next days info
        i_row_n += 1
        i_row += 1
        i_row_d += 1
    else:
        mon_s = first(mon)
        mon_e = second(mon)
        # Add name to column A and roster times to column E and F
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=mon_s)
        ws.cell(t_row + i_row, t_col + 5, value=mon_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Monday')
        ws.cell(d_row + i_row_d, d_col + 1, value=monday)
    
    # Add one to shift rows down for next dayd info
    i_row_n += 1
    i_row += 1
    i_row_d += 1

    # Tueday
    if tue == 'AF':
        tue = 0
        # Add name to column A
        ws.cell(t_row + i_row_n, t_col, value=name)
        # Add roster time to column E and F
        ws.cell(t_row + i_row, t_col + 4, value=tue)
        ws.cell(t_row + i_row, t_col + 5, value=tue)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Tuesday')
        ws.cell(d_row + i_row_d, d_col + 1, value=tuesday)
    
    elif first(tue) == 18:
        tue_s = first(tue)
        tue_e = second(tue)
        # Add name to column A and start time to column E
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=tue_s)
        # Add name to column A and shift row down by one and
        # Add end time to column F
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=tue_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Tuesday')
        ws.cell(d_row + i_row_d, d_col + 1, value=tuesday)
        # Add dates to columns C and D and shift row down by 1
        ws.cell(d_row + i_row_d + 1, d_col, value='Tuesday')
        ws.cell(d_row + i_row_d + 1, d_col + 1, value=tuesday)
        # Add one to shift rows down for next days info
        i_row_n += 1
        i_row += 1
        i_row_d += 1
    else:
        tue_s = first(tue)
        tue_e = second(tue)
        # Add name to column A and roster times to column E and F
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=tue_s)
        ws.cell(t_row + i_row, t_col + 5, value=tue_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Tuesday')
        ws.cell(d_row + i_row_d, d_col + 1, value=tuesday)
    
    # Add one to shift rows down for next dayd info
    i_row_n += 1
    i_row += 1
    i_row_d += 1

    # Wednesday
    if wed == 'AF':
        wed = 0
        # Add name to column A
        ws.cell(t_row + i_row_n, t_col, value=name)
        # Add roster time to column E and F
        ws.cell(t_row + i_row, t_col + 4, value=wed)
        ws.cell(t_row + i_row, t_col + 5, value=wed)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Wednesday')
        ws.cell(d_row + i_row_d, d_col + 1, value=wednesday)
    
    elif first(wed) == 18:
        wed_s = first(wed)
        wed_e = second(wed)
        # Add name to column A and start time to column E
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=wed_s)
        # Add name to column A and shift row down by one and
        # Add end time to column F
        ws.cell(t_row + i_row_n + 1, t_col, value=name)
        ws.cell(t_row + i_row + 1, t_col + 5, value=wed_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Wednesday')
        ws.cell(d_row + i_row_d, d_col + 1, value=wednesday)
        # Add dates to columns C and D and shift row down by 1
        ws.cell(d_row + i_row_d + 1, d_col, value='Wednesday')
        ws.cell(d_row + i_row_d + 1, d_col + 1, value=wednesday)
        # Add one to shift rows down for next days info
        i_row_n += 1
        i_row += 1
        i_row_d += 1
    else:
        wed_s = first(wed)
        wed_e = second(wed)
        # Add name to column A and roster times to column E and F
        ws.cell(t_row + i_row_n, t_col, value=name)
        ws.cell(t_row + i_row, t_col + 4, value=wed_s)
        ws.cell(t_row + i_row, t_col + 5, value=wed_e)
        # Add dates to columns C and D
        ws.cell(d_row + i_row_d, d_col, value='Wednesday')
        ws.cell(d_row + i_row_d, d_col + 1, value=wednesday)
    
    # Add one to shift rows down for next dayd info
    i_row_n += 1
    i_row += 1
    i_row_d += 1 
    
    
    # Add one to shift rows down for next dayd info
    i_row_n += 2
    i_row += 2
    i_row_d += 2
    con.close()

# Close workbook
wb.save('Wage Times.xlsx')
wb.close()