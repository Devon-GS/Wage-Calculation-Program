import sqlite3

def clean_db():
    con = sqlite3.connect("wageTimes.db")
    c = con.cursor()

    c.execute('DELETE FROM clockTimeAWO')
    c.execute('DELETE FROM rosterAttWeekOne')
    c.execute('DELETE FROM rosterAttWeektwo')
    c.execute('DELETE FROM attTotal')


    con.commit()
    con.close()


# con = sqlite3.connect("wageTimes.db")
# c = con.cursor()

# c.execute('DROP TABLE attTotal')

# con.commit()
# con.close()