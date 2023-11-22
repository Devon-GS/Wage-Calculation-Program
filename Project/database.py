import sqlite3

def clean_db():
    con = sqlite3.connect("wageTimes.db")
    c = con.cursor()

    
    c.execute('DELETE FROM rosterAttWeekOne')
    c.execute('DELETE FROM rosterAttWeektwo')
    c.execute('DELETE FROM ClockTimeAttendent')
    c.execute('DELETE FROM attTotal')

    c.execute('DELETE FROM rosterCashierWeekOne')
    c.execute('DELETE FROM rosterCashierWeekTwo')
    c.execute('DELETE FROM ClockTimeCashier')
    c.execute('DELETE FROM cashierTotal')


    con.commit()
    con.close()


# con = sqlite3.connect("wageTimes.db")
# c = con.cursor()

# c.execute('DROP TABLE clockTimeAWO')


# con.commit()
# con.close()