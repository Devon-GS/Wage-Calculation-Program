import sqlite3
import pandas as pd
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
			c.execute("CREATE TABLE IF NOT EXISTS employeeNames (englishName TEXT, fullName TEXT, Surname TEXT, idPass TEXT UNIQUE)")
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
	
	def update_employees(self, ename, fname, sname, id):
		with closing(self.get_connection()) as con:
			c = con.cursor()
			try:
				# Check to see if non english name
				if fname == '':
					fname = '0'

				c.execute(f'''UPDATE employeeNames SET
								englishName = :ename,
								fullName = :fname,
								surname = :sname

								WHERE idPass = :id''',
								{
									'ename' : ename,
									'fname' : fname,
									'sname' : sname,
									'id' : id
								})

				con.commit()
				CTkMessagebox(title="Update Employee", message="Employee Update Successfuly", icon="info")
			except Exception as error:
				CTkMessagebox(title="Error", message=error, icon="cancel")

	def delete_employees(self, id):
		with closing(self.get_connection()) as con:
			c = con.cursor()
			try:
				c.execute(f'''DELETE FROM employeeNames WHERE idPass = :id''',
								{
									'id' : id
								})

				con.commit()
				CTkMessagebox(title="Delete Employee", message="Employee Deleted Successfuly", icon="info")
			except Exception as error:
				CTkMessagebox(title="Error", message=error, icon="cancel")
				
	def bulk_add_employees(self):
		with closing(self.get_connection()) as con:
			c = con.cursor()
			try:
				# Get employee info from bulk file
				employee_names_file = 'Templates/Bulk_Employee_Add.csv'
				employee_info = pd.read_csv(employee_names_file)
				employee_list = employee_info.values.tolist()

				# Loop through and add to database
				for x in employee_list:
					ename = str(x[0]).strip()
					fname = str(x[1]).strip()
					sname = str(x[2]).strip()
					id = str(x[3]).strip()

					query = """INSERT INTO employeeNames (englishName, fullName, Surname, idPass)
								VALUES (?, ?, ?, ?)"""
						
					c.execute(query, (ename, fname, sname, id))

					con.commit()
					CTkMessagebox(title="Bulk Add", message="Bulk Add Complete Successfuly", icon="info")
			except Exception as error:
				CTkMessagebox(title="Error", message=error, icon="cancel")