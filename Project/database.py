# Handles all SQLite interactions


import sqlite3
from CTkMessagebox import CTkMessagebox
from contextlib import closing
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

	# EMPLOYEE MANAGEMENT 
	def add_employees(self, ename, fname, sname, id):
		with closing(self.get_connection()) as con:
			c = con.cursor()
			try:
				if ename == '' or sname == '' or id == '':
					raise ValueError('First Name, Surname and ID Cannot Be Blank!')
				else:
					# Check to see if non english name
					if fname == '':
						fname = '0'

					query = """INSERT INTO employeeNames (englishName, fullName, Surname, idPass)
							VALUES (?, ?, ?, ?)"""
					
					c.execute(query, (ename, fname, sname, id))
					con.commit()
					CTkMessagebox(title="Success", message="Employee Added Successfully", icon="info")

			except Exception as error:
				CTkMessagebox(title="Error", message=error, icon="cancel")

	def search_employees(self):
		with closing(self.get_connection()) as con:
			c = con.cursor()
			try:
				con = sqlite3.connect("wageTimes.db")
				c = con.cursor()

				c.execute("SELECT englishName, idPass FROM employeeNames")
				
				records = c.fetchall()
				con.commit()

				# Dic
				results = {}

				# Made records in to dic
				for x in records:
					results[x[0]] = x[1]
				
			except Exception as error:
				CTkMessagebox(title="Error", message=error, icon="cancel")
		
		
		return results

	def employee_selected_option(self, id):
		with closing(self.get_connection()) as con:
			c = con.cursor()
			try:
				con = sqlite3.connect("wageTimes.db")
				c = con.cursor()

				c.execute(f"""SELECT englishName,
									fullName,
									Surname,
									idPass
								FROM
									employeeNames
								WHERE
									idpass = {id}
			  			""")
								
				record = c.fetchall()
				con.commit()
				
			except Exception as error:
				CTkMessagebox(title="Error", message=error, icon="cancel")
		
		
		return record		
	
# ------------- Working ---------------------