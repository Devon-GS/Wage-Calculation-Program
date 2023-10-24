import os
from datetime import datetime
import pandas as pd
import sqlite3

clock_list = []
for filename in os.listdir('../Uniclox/'):
    file = filename.replace(" ","")
    if 'TL' in file and file[-7:-4] != '000':
        clock_list.append(filename)
        
# recent = clock_list[-30:]
recent = clock_list[-30:]

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


# for dt in date_times:
#     con = sqlite3.connect("test.db")
#     c = con.cursor()
#     c.execute("""CREATE TABLE IF NOT EXISTS clockTime (badge TEXT, date TEXT, time TEXT)""")

#     query = """INSERT INTO clockTime (badge, date, time) VALUES (?, ?, ?)"""

#     c.execute(query, (dt[0], dt[1], dt[2]))

#     con.commit()

# con.close()



con = sqlite3.connect("test.db")
c = con.cursor()

c.execute('SELECT time FROM clockTime WHERE badge = ? AND date = ?', ('20', '19/10/23'))
gg = c.fetchall()

con.close()
print(gg)

if gg[0][0] < '04:05:04':
    print('yes')
else:
    print('no')


# loop through excel and search database for time that match
# then past to clock time to excel
# handle double clocks