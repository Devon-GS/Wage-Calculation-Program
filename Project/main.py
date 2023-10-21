from datetime import datetime
import pandas as pd
import xlsxwriter
import re

# ==============================================================================
# IMPORT ROSTER TIMES AND DATES (ATTENDENTS)
# ==============================================================================
file = '../Attendant_Carwash_Roster.xls'

# Get Columns
columns = ['idx','ATTENDANTS', 'THURS', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']
# Get Times 
data = pd.read_excel(file, index_col=0, header=1, usecols=columns)
# Get Dates
data_date = pd.read_excel(file, header=None, nrows=2, usecols='C:I')
data_date_ex = data_date.loc[0]

week1_dates = {}
for x in data_date_ex:
    week1_dates[x.date().strftime("%d/%m/%y")] = x.strftime("%A")

# Get week one from excel sheet
week_one_data = data.loc[0:15]
week_one = []
for x in week_one_data.to_numpy():
    if str(x[0]) != 'nan':
        week_one.append(x)


# Get Time in / Time Out
def first(weekday):
        first = float(re.findall('[0-9]+(?=.*\-)', weekday)[0])
        return first

# def second(weekday):
#        second = float(re.findall('\-(.*)', weekday)[0])
#        return second

for week in week_one:
        name = week[0]
        thur = first(week[1])
        # fri = week[2]
        # sat = week[3]
        # sun = week[4]
        # mon = week[5]
        # tue = week[6]
        # wed = week[7]
        # sune = 0.0
        # mone = 0.0

        print(thur)



# ==============================================================================
# CREATE TIME EXCEL FILE
# ==============================================================================

# workbook = xlsxwriter.Workbook('Wage Times.xlsx')
 
# # The workbook object add worksheet
# worksheet = workbook.add_worksheet()
 
# # Create Cloumn Headings
# worksheet.write('A1', 'Name')
# worksheet.write('B1', 'Badge Number')
# worksheet.write('C1', 'Date')
# worksheet.write('D1', 'Week Day')
# worksheet.write('E1', 'Time In')
# worksheet.write('F1', 'Time Out')
# worksheet.write('G1', 'Hours')

# # Start from the first cell. Rows and columns are zero indexed.
# row = 1
# col = 2

# # Iterate over the data and write it out row by row
# for date, dayname in week1_dates.items():
#     worksheet.write(row, col, date)
#     worksheet.write(row, col + 1, dayname)
#     row += 1

# # Get Time in / Time Out

# # Close workbook
# workbook.close()