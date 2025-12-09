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

    c.execute('DELETE FROM carwashTotal')

    con.commit()
    con.close()

def clean_carwash():
    con = sqlite3.connect("wageTimes.db")
    c = con.cursor()
    
    c.execute('DELETE FROM carwashTotal')

    con.commit()
    con.close()
    
def clean_db_recal():
    con = sqlite3.connect("wageTimes.db")
    c = con.cursor()
    
    c.execute('DELETE FROM attTotal')
    c.execute('DELETE FROM cashierTotal')

    con.commit()
    con.close()

def db_init():
    con = sqlite3.connect("wageTimes.db")
    c = con.cursor()
    
    # CREATE ATTENDENT TABLES FOR FIRST TIME
    c.execute("""CREATE TABLE IF NOT EXISTS rosterAttWeekOne (
            name TEXT,
            badge TEXT,
            thur TEXT,
            fri TEXT,
            sat TEXT,
            sun TEXT,
            mon TEXT,
            tue TEXT,
            wed TEXT
            )"""
    )

    c.execute("""CREATE TABLE IF NOT EXISTS rosterAttWeekTwo (
            name TEXT,
            badge TEXT,
            thur TEXT,
            fri TEXT,
            sat TEXT,
            sun TEXT,
            mon TEXT,
            tue TEXT,
            wed TEXT
            )"""
    )

    c.execute("""CREATE TABLE IF NOT EXISTS ClockTimeAttendent (badge TEXT, date TEXT, time TEXT)""")

    c.execute("""CREATE TABLE IF NOT EXISTS attTotal (
                name TEXT,
                badge TEXT,
                normal TEXT,
                sunday TEXT,
                public TEXT,
                noClock TEXT
                )""")
    
    # CREATE CASHIER TABLES FOR FIRST TIME
    c.execute("""CREATE TABLE IF NOT EXISTS rosterCashierWeekOne (
            name TEXT,
            badge TEXT,
            thur TEXT,
            fri TEXT,
            sat TEXT,
            sun TEXT,
            mon TEXT,
            tue TEXT,
            wed TEXT
            )"""
    )

    c.execute("""CREATE TABLE IF NOT EXISTS rosterCashierWeekTwo (
            name TEXT,
            badge TEXT,
            thur TEXT,
            fri TEXT,
            sat TEXT,
            sun TEXT,
            mon TEXT,
            tue TEXT,
            wed TEXT
            )"""
    )

    c.execute("""CREATE TABLE IF NOT EXISTS ClockTimeCashier (badge TEXT, date TEXT, time TEXT)""")

    c.execute("""CREATE TABLE IF NOT EXISTS cashierTotal (
                name TEXT,
                badge TEXT,
                normal TEXT,
                sunday TEXT,
                public TEXT,
                noClock TEXT,
                cashier TEST
                )""")
    
    # CREATE CARWASH TABLES FOR FIRST TIME
    c.execute("""CREATE TABLE IF NOT EXISTS carwashTotal (
        name TEXT,
        badge TEXT,
        normal TEXT,
        sunday TEXT,
        public TEXT,
        extra TEXT
        )"""
)
    
    # CREATE CARWASH TABLES FOR FIRST TIME
    c.execute("""CREATE TABLE IF NOT EXISTS employeeNames (
        englishName TEXT,
        fullName TEXT,
        Surname TEXT,
        idPass TEXT
        )"""
)

    con.commit()
    con.close()

# ############# TESTING ##############
# con = sqlite3.connect("wageTimes.db")
# c = con.cursor()

# c.execute('DROP TABLE cashierTotal')

# con.commit()
# con.close()
# ############# TESTING ##############