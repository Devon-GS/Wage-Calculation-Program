import sqlite3
import pandas as pd
from CTkMessagebox import CTkMessagebox
from contextlib import closing
from config import DB_PATH


def get_connection():
	return sqlite3.connect(DB_PATH)

def initialize_tables():
	with get_connection() as con:
		c = con.cursor()
	
		c.execute("CREATE TABLE IF NOT EXISTS rosterAttendant (name TEXT, badge TEXT, day TEXT, date TEXT, shift TEXT, week TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS rosterCashier (name TEXT, badge TEXT, day TEXT, date TEXT, shift TEXT, week TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS uniclox (badge TEXT, date TEXT, time TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS attTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, noClock TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS cashierTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, noClock TEXT, cashier TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS carwashTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, extra TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS employeeNames (englishName TEXT, fullName TEXT, Surname TEXT, idPass TEXT UNIQUE)")
		c.execute("CREATE TABLE IF NOT EXISTS publicHolidays (date TEXT)")

		con.commit()
		CTkMessagebox(title="Success", message="Successfully Initialized The Database", icon="info")

def clear_session_data(table=None):
	with closing(get_connection()) as con:
		c = con.cursor()

		if table == None:
			tables_name = ['rosterAttendant', 'rosterCashier',
				    'uniclox', 'attTotal', 'cashierTotal', 'carwashTotal']
		else:
			tables_name = [table]
		
		for table in tables_name: c.execute(f"DELETE FROM {table}")
		con.commit()

# --- GET PUBBLIC HOLIDAYS --- 
def public_holidays_db(holidays):
	with closing(get_connection()) as con:
		c = con.cursor()
		try:
			clear_session_data('publicHolidays')

			c.executemany(f"INSERT INTO publicHolidays (date) VALUES (?)", holidays)
			con.commit()

		except Exception as error:
			CTkMessagebox(title="Error", message=error, icon="cancel")

def get_public_holidays():
	# public_holidays = []
	with closing(get_connection()) as con:
		c = con.cursor()
		try:
			c.execute("SELECT * FROM publicHolidays")
			
			public_holidays = [records[0] for records in c.fetchall()]
			con.commit()

		except Exception as error:
			CTkMessagebox(title="Error", message=error, icon="cancel")
	
	return public_holidays

# --- EMPLOYEE MANAGEMENT --- 
def add_employees(ename, fname, sname, id):
	with closing(get_connection()) as con:
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

def search_employees():
	with closing(get_connection()) as con:
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

def employee_selected_option(id):
	with closing(get_connection()) as con:
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

def update_employees(ename, fname, sname, id):
	with closing(get_connection()) as con:
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

def delete_employees(id):
	with closing(get_connection()) as con:
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
			
def bulk_add_employees():
	with closing(get_connection()) as con:
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


# --- ADD SHIFTS ---
def add_shifts(shifts, role, week):
	with closing(get_connection()) as con:
		c = con.cursor()
		if role == "Attendant":
			table = "rosterAttendant"
		else:
			table = "rosterCashier"

		try:
			query = f"""INSERT INTO {table} (name, badge, day, date, shift, week)
							VALUES (?, ?, ?, ?, ?, ?)"""
			
			# Loop through shift info and add to database
			for x in shifts:
				c.execute(query, (x[0], x[1], x[2], x[3], x[4], x[5]))
				con.commit()

			# CTkMessagebox(title="Add Shifts", message=f"Added {role} shifts for {week[:4] + " " + week[4:]} Successfuly", icon="info")
		except Exception as error:
			CTkMessagebox(title="Error", message=error, icon="cancel")

#  -- ADD CLOCK TIMES --
def add_clock_times(clock_times):
	with closing(get_connection()) as con:
		c = con.cursor()
		try:
			c.executemany(f"INSERT INTO uniclox (badge, date, time) VALUES (?, ?, ?)", clock_times)
			con.commit()
		except Exception as error:
			CTkMessagebox(title="Error", message=error, icon="cancel")

# --- GET SHIFT TIMES FOR EXCEL ---
def get_shift_times_db(roster, week):
	with closing(get_connection()) as con:
		c = con.cursor()
		try:
			if roster == "Attendant":
				table = 'rosterAttendant'
			else:
				table = 'rosterCashier'

			c.execute(f"SELECT * FROM {table} WHERE week=?", (week,))
							
			records = c.fetchall()
			con.commit()

		except Exception as error:
			CTkMessagebox(title="Error", message=error, icon="cancel")

		return records
	
# --- GET CLOCK TIMES ---
def get_clock_times():
	with closing(get_connection()) as con:
		c = con.cursor()
		try:
			c.execute("SELECT * FROM uniclox")
			
			clocks = c.fetchall()
			con.commit()

		except Exception as error:
			CTkMessagebox(title="Error", message=error, icon="cancel")
		
		return clocks
	
# --- CARWASH TIMES ---
def carwash_db(data):
	with closing(get_connection()) as con:
		c = con.cursor()
		try:
			query = """INSERT INTO carwashTotal (name, badge, normal, sunday, public, extra)
				  			VALUES (?, ?, ?, ?, ?, ?)"""
			
			for k, v in data.items():
				c.execute(query,(v[0], k, v[1], v[2], 0, v[3]))			
				con.commit()

		except Exception as error:
			CTkMessagebox(title="Error", message=error, icon="cancel")	