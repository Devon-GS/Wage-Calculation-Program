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
		c.execute("CREATE TABLE IF NOT EXISTS cashierTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, noClock TEXT, cashier TEXT, cashierSun TEXT, cashierPub TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS carwashTotal (name TEXT, badge TEXT, normal TEXT, sunday TEXT, public TEXT, extra TEXT)")
		c.execute("CREATE TABLE IF NOT EXISTS employeeNames (englishName TEXT, fullName TEXT, Surname TEXT, idPass TEXT UNIQUE)")
		c.execute("CREATE TABLE IF NOT EXISTS publicHolidays (date TEXT)")

		con.commit()
	
def clear_session_data(table=None):
	with closing(get_connection()) as con:
		c = con.cursor()

		if table == None:
			tables_name = ['rosterAttendant', 'rosterCashier',
					'uniclox', 'attTotal', 'cashierTotal', 'carwashTotal']
		else:
			tables_name = [table]
		
		for table in tables_name: 
			c.execute(f"DELETE FROM {table}")
		con.commit()

# --- GET PUBBLIC HOLIDAYS --- 
def public_holidays_db(holidays):
	with closing(get_connection()) as con:
		c = con.cursor()
		
		clear_session_data('publicHolidays')

		c.executemany(f"INSERT INTO publicHolidays (date) VALUES (?)", holidays)
		con.commit()

def get_public_holidays():
	# public_holidays = []
	with closing(get_connection()) as con:
		c = con.cursor()
		
		c.execute("SELECT * FROM publicHolidays")
			
		public_holidays = [records[0] for records in c.fetchall()]

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
			
# --- EMPLOYEE INFO FOR PAYSLIPS ---
def get_emp_info():
	with closing(get_connection()) as con:
		c = con.cursor()
	
		c.execute("SELECT * FROM employeeNames")
		records = c.fetchall()
		
		# Convert the list of tuples into a dictionary
		emp_dict = {}
		for row in records:
			# Get employee name as key
			excel_name_key = row[0].strip() 
			
			# Assign the whole row as the value for this key
			emp_dict[excel_name_key] = row

		return emp_dict

# --- ADD SHIFTS ---
def add_shifts(shifts, role, week):
	with closing(get_connection()) as con:
		c = con.cursor()
		if role == "Attendant":
			table = "rosterAttendant"
		else:
			table = "rosterCashier"

		query = f"""INSERT INTO {table} (name, badge, day, date, shift, week)
						VALUES (?, ?, ?, ?, ?, ?)"""
		
		# Loop through shift info and add to database
		for x in shifts:
			c.execute(query, (x[0], x[1], x[2], x[3], x[4], x[5]))
			con.commit()

#  -- ADD CLOCK TIMES --
def add_clock_times(clock_times):
	with closing(get_connection()) as con:
		c = con.cursor()

		c.executemany(f"INSERT INTO uniclox (badge, date, time) VALUES (?, ?, ?)", clock_times)
		con.commit()

# --- GET SHIFT TIMES FOR EXCEL ---
def get_shift_times_db(roster, week):
	with closing(get_connection()) as con:
		c = con.cursor()
		
		if roster == "Attendant":
			table = 'rosterAttendant'
		else:
			table = 'rosterCashier'

		c.execute(f"SELECT * FROM {table} WHERE week=?", (week,))
						
		records = c.fetchall()

	return records
	
# --- GET CLOCK TIMES ---
def get_clock_times():
	with closing(get_connection()) as con:
		c = con.cursor()

		c.execute("SELECT * FROM uniclox")
		
		clocks = c.fetchall()
	
	return clocks
	
# --- ADD CARWASH TIMES ---
def carwash_db(data):
	with closing(get_connection()) as con:
		c = con.cursor()
	
		query = """INSERT INTO carwashTotal (name, badge, normal, sunday, public, extra)
					VALUES (:name, :badge, :n_hours, :s_hours, 0, :amount)"""
	
		records = []
		
		# Add badge to data, so SQL has access
		for badge, emp_data in data.items():
			emp_data['badge'] = badge 
			records.append(emp_data)

		c.executemany(query, records)			
		
		con.commit()

# --- ADD TOTAL HOURS ---
def add_total_hours_db(totals, role):
	# Configuration based on role
	if role == "Attendant":
		table = 'attTotal'
		cols = "(name, badge, normal, sunday, public, noClock)"
	   
		data_list = [
			(k, v['badge'], v['std'], v['sun'], v['pub'], v['nc']) 
			for k, v in totals.items()
		]
	else:
		table = 'cashierTotal'
		cols = "(name, badge, normal, sunday, public, noClock, cashier, cashierSun, cashierPub)"
		data_list = [
			(k, v['badge'], v['std'], v['sun'], v['pub'], v['nc'], v['cstd'], v['csun'], v['cpub']) 
			for k, v in totals.items()
		]

	# Create placeholders (?, ?, ?) based on the number of columns
	placeholders = ", ".join(["?"] * len(data_list[0]))
	query = f"INSERT INTO {table} {cols} VALUES ({placeholders})"

	# Database Operation
	with closing(get_connection()) as con:
		c = con.cursor()
		
		c.executemany(query, data_list)
		con.commit()

# --- Get TOTAL HOURS ---
def get_total_hours():
	all_hours = []
	
	querys = [
		'SELECT * FROM carwashTotal',
		'SELECT * FROM attTotal',
		'SELECT * FROM cashierTotal'           
	]

	# Ensure you have imported closing and get_connection is defined
	with closing(get_connection()) as con:
		c = con.cursor()
		for query in querys:
			c.execute(query)
			rec = c.fetchall()
			# Append the results to our list
			all_hours.extend(rec) 
		
		# Return the data so you can use it outside the function
		return all_hours