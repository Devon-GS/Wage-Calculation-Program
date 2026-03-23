# Handles all SQLite interactions


import sqlite3
from config import DB_PATH

class DatabaseManager:
    def get_connection(self):
        return sqlite3.connect(DB_PATH)

    def initialize_tables(self):
        with self.get_connection() as con:
            c = con.cursor()
            # Roster & Total Tables
            for role in ["Att", "Cashier"]:
                for week in ["One", "Two"]:
                    c.execute(f"CREATE TABLE IF NOT EXISTS roster{role}Week{week} (name TEXT, badge TEXT, thur TEXT, fri TEXT, sat TEXT, sun TEXT, mon TEXT, tue TEXT, wed TEXT)")
            
            c.execute("CREATE TABLE IF NOT EXISTS ClockTimeAttendent (badge TEXT, date TEXT, time TEXT)")
            c.execute("CREATE TABLE IF NOT EXISTS ClockTimeCashier (badge TEXT, date TEXT, time TEXT)")
            c.execute("CREATE TABLE IF NOT EXISTS attTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, noClock TEXT)")
            c.execute("CREATE TABLE IF NOT EXISTS cashierTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, noClock TEXT, cashier TEXT)")
            c.execute("CREATE TABLE IF NOT EXISTS carwashTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, extra TEXT)")
            c.execute("CREATE TABLE IF NOT EXISTS employeeNames (englishName TEXT, fullName TEXT, Surname TEXT, idPass TEXT)")
            con.commit()

    def clear_session_data(self):
        with self.get_connection() as con:
            c = con.cursor()
            tables = ['ClockTimeAttendent', 'ClockTimeCashier', 'attTotal', 'cashierTotal', 'carwashTotal']
            for table in tables: c.execute(f"DELETE FROM {table}")
            con.commit()