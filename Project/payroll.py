from openpyxl import Workbook
from openpyxl import load_workbook
import sqlite3

# GET ALL HOURS FROM DATA BASE
con = sqlite3.connect("wageTimes.db")
c = con.cursor()

c.execute("SELECT name, badge, normal, sunday, public FROM attTotal")
records_att = c.fetchall()

c.execute("SELECT name, badge, normal, sunday, public, cashier FROM cashierTotal")
records_cash = c.fetchall()

# CARWASH TIMES
# c.execute("SELECT name, badge, normal, sunday, public, cashier FROM cashierTotal")
# records_cash = c.fetchall()
# total_records.append(records_cash)

con.commit()
con.close()

# JOIN ALL INFO FROM DATABASE INTO ONE TABLE
total_records = []

for rec in records_att:
    total_records.append(rec)

for rec in records_cash:
    total_records.append(rec)


# COPY HOURS FROM TOTAL INFO TO PAYROLL.XLSX
wb = load_workbook("Payroll/Payroll.xlsx")
ws = wb['Wages']

# Iterate over payrol columns
columns = ws.iter_cols(min_row=1, min_col=3)

# If badge number in payrol match info copy hours
for col in columns:
    for r in total_records:
        if 'b' in str(col[1].value) and str(col[1].value)[:-1] == str(r[1]):
            col[2].value = float(r[5])
            col[11].value = 0.00
            col[14].value = 0.00
        elif col[1].value == int(r[1]):
                col[2].value = float(r[2])
                col[11].value = float(r[3])
                col[14].value = float(r[4])

wb.save("Payroll/Payroll.xlsx")
wb.close()