import os
from datetime import datetime, time
from openpyxl import Workbook
from openpyxl import load_workbook
import pandas as pd
import sqlite3

# Loop through clock files and collect last 30 files
clock_list = []
for filename in os.listdir('../Uniclox/'):
    file = filename.replace(" ","")
    if 'TL' in file and file[-7:-4] != '000':
        clock_list.append(filename)
        
recent = clock_list[-100:]

# Loop though each clock file by line and append badge and times to dat_times list
date_times = []
for file in recent:    
    f = open('../Uniclox/' + file, 'r')
    for line in f:
        line = line.strip()
        line = line.split(',')
        badge = int(line[0])
        dated = line[1]
        timestamp = datetime.strptime(dated, '%Y-%m-%d %H:%M:%S').strftime("%d/%m/%y %H:%M:%S")
        split_timestamp = timestamp.split()
        x = [badge, split_timestamp[0], split_timestamp[1]]
        date_times.append(x)

# # Push all info in date_times list to database
# for dt in date_times:
#     con = sqlite3.connect("test.db")
#     c = con.cursor()
#     c.execute("""CREATE TABLE IF NOT EXISTS clockTime (badge TEXT, date TEXT, time TEXT)""")

#     query = """INSERT INTO clockTime (badge, date, time) VALUES (?, ?, ?)"""

#     c.execute(query, (dt[0], dt[1], dt[2]))

#     con.commit()

# con.close()

# ###############################################################################################
# WORKING
# ###############################################################################################

con = sqlite3.connect("test.db")
c = con.cursor()

wb = load_workbook('Wage Times.xlsx')
ws = wb.active

i = 0

for x in range(ws.max_row + 1):
    times_in_before = []
    times_in_late = []

    times_out_before = []
    times_out_late = []
    
    
    badge = ws.cell(row=2 + i, column=2).value
    date = ws.cell(row=2 + i, column=4).value
    
    if badge != None and date != None:
        c.execute('SELECT time FROM clockTime WHERE badge = ? AND date = ?', (badge, date))
        clock_times = c.fetchall()

        times = []
        if len(clock_times) >= 3:
            for x in clock_times:
                times.append(x[0])
        else:
            for x in clock_times:
                times.append(x[0])
            
        for x in times:
            format = "%H:%M"
            t = ws.cell(row=2 + i, column=5).value
            to = ws.cell(row=2 + i, column=6).value
        
            time_in = time(t, 0, 0).strftime(format)
            time_out = time(to, 0, 0).strftime(format)

            if x <= time_in:
                times_in_before.append(x)
            elif x > time_in:
                times_in_late.append(x)

            if x <= time_out:
                times_out_before.append(x)
            elif x > time_in:
                times_out_late.append(x)

        if times_in_before:
            time_r = max(times_in_before)
            ws.cell(row=2+ i, column=7, value=time_r)
        elif times_in_late:
            time_r = min(times_in_late)
            ws.cell(row=2 + i, column=7, value=time_r)

    i += 1
    
wb.save('Wage Times.xlsx')
wb.close()


# loop through excel and search database for time that match
# then past to clock time to excel
# handle double clocks