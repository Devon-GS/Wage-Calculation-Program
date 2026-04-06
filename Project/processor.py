import os
import re
import pandas as pd
import database as db
from datetime import datetime, timedelta, time
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from config import (CREATE_EXCEL, WAGE_TIMES_FILE, PUBLIC_HOILIDAY_FILE, UNICLOX_FOLDER, ATT_ROSTER_FILE, CAS_ROSTER_FILE, 
					BADGE_NUMBER_FILE, BAKER_CASHIER_FILE, CARWASH_FILE, COLUMN_WIDTHS_ATT, COLUMN_WIDTHS_TOTALS ,COL_DIFF)


# --- Helper Functions ---
def clear_excel():
	if os.path.isfile(WAGE_TIMES_FILE):
		os.remove(WAGE_TIMES_FILE)

def load_excel():
	"""Opens the workbook and returns the object."""
	if not os.path.isfile(WAGE_TIMES_FILE):
		CREATE_EXCEL()

	wb = load_workbook(WAGE_TIMES_FILE) 
	return wb  

def save_workbook(wb):
	"""Saves the workbook to the disk."""
	wb.save(WAGE_TIMES_FILE)
	wb.close()

def get_badge_mapping():
	"""Creates a dictionary {Name: BadgeID} from the badges.xlsx file."""
	mapping = {}
	if os.path.exists(BADGE_NUMBER_FILE):
		# Load without headers
		df = pd.read_excel(BADGE_NUMBER_FILE, header=None)
		for index, row in df.iterrows():
			# row[0] is Name, row[1] is Badge
			mapping[str(row[0]).strip()] = str(row[1]).strip()
	return mapping

def get_cashier_dates():
	"""Gets dates of cashier shifts for employee that works cashier and baker shifts"""
	if os.path.exists(BAKER_CASHIER_FILE):
		wb = load_workbook(BAKER_CASHIER_FILE, data_only=True)
		ws = wb.active

		bc_working = []
		for row in ws.iter_rows(min_row=2, max_col=2, max_row=20, values_only=True):
			x = row
			if x[0] != None:
				name = x[0]
				cashier_date = x[1].strftime('%d/%m/%Y')
				bc = [name, cashier_date]
				bc_working.append(bc)

		wb.close()
		return bc_working

def split_roster_time(val):
	"""Replaces the old first() and second() regex functions."""
	if val in ["AF", " ", "0", 0, None, ""]:
		return 0.0, 0.0
	try:
		# Matches "08-17" or "18-06"
		times = re.findall(r"(\d+)", str(val))
		return int(times[0]), int(times[1])
	except:
		return 0.0, 0.0

def collect_public_holidays():
	"""Get public holidays and save to database"""
	holidays = []
	if os.path.exists(PUBLIC_HOILIDAY_FILE):
		wb = load_workbook(PUBLIC_HOILIDAY_FILE, data_only=True)
		ws = wb.active
		for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
			if row[0]: 
				holidays.append((row[0].strftime('%d/%m/%Y'),))
		wb.close()
	db.public_holidays_db(holidays)

def adjust_time(clock_hours, roster_h, day, is_in):
	"""
	1. Rounding logic - changes the dicimal to 15, 30 or 45
	2. m = Minutes and h = Hours 
	"""
	if not clock_hours: return float(roster_h)
	h, m = map(int, clock_hours.split(':'))

	# Clock In Logic
	if is_in: 
		if h > roster_h or (h == roster_h and m > 0):
			# Special logic for Sunday: No 4-minute grace period
			if day == "Sunday":
				if m <= 15: 
					return h + 0.25
				elif m <= 30: 
					return h + 0.50
				elif m <= 45: 
					return h + 0.75
				else: 
					return float(h + 1)
			 # Standard logic for all other days
			else:
				if m <= 4: 
					return float(h) # Gives employee 4 min to clock in
				elif m <= 15: 
					return h + 0.25
				elif m <= 30: 
					return h + 0.50
				elif m <= 45: 
					return h + 0.75
				else: 
					return float(h + 1)
		return float(roster_h)
	# Clock Out Logic
	else: 
		if h < roster_h:
			# if m <= 4: 
			# 	return float(h)
			if m <= 15: 
				return float(h)
			elif m <= 30: 
				return h + 0.25
			elif m <= 45: 
				return h + 0.50
			else: 
				return h + 0.75
		return float(roster_h)


# --- Helper Functions End ---


# --- Step 1: Roster to Database ---
def roster_shift_to_db(role="Attendant", week="WeekOne"):
	"""
	1. Reads the Roster (Attendant or Cashier).
	2. Gets the Badge Mapping.
	3. Get roster shifts week one and two for attendants and cashiers.
	4. Writes names, badges, days, dates and shift to database.
	"""
	# Get Path and Badge Mapping
	file_path = ATT_ROSTER_FILE if role == "Attendant" else CAS_ROSTER_FILE
	badges = get_badge_mapping()
	
	# Load the Roster via Pandas
	try:
		if role == "Attendant" and week == "WeekOne":
			# Columns used
			cols = "B:I"

			# Dates slice
			drow = 0
			d_col_start = 1
			d_col_end = 8

			# Week times slice
			wrow = 2
			wrow_end = 17

		# Att week two
		elif role == "Attendant" and week == "WeekTwo":
			# Columns used
			cols = "B:I"
			
			# Dates slice
			drow = 28
			d_col_start = 1
			d_col_end = 8

			# Week times slice
			wrow = 30
			wrow_end = 45

		# Cashier week one 
		elif role == "Cashier" and week == "WeekOne":
			# Columns used
			cols = "B:I"
			
			# Dates slice
			drow = 3
			d_col_start = 1
			d_col_end = 8

			# Week times slice
			wrow = 5
			wrow_end = 11

			# Week times slice (bakers)
			wbrow = 31
			wbrow_end = 33
		
		# Cashier week two 
		elif role == "Cashier" and week == "WeekTwo":
			# Columns used
			cols = "B:I"
			
			# Dates slice
			drow = 34
			d_col_start = 1
			d_col_end = 8

			# Week times slice
			wrow = 14
			wrow_end = 20

			# Week times slice (bakers)
			wbrow = 36	
			wbrow_end = 38
		
		# Get times from excel
		df = pd.read_excel(file_path, header=None, usecols=cols, nrows=46)
		data = df.fillna(0)

		# Extract the dates  
		week_dates = data.iloc[drow, d_col_start : d_col_end]

		# Extract the employee schedule block
		if role == "Cashier":
			cashier_times = data.iloc[wrow : wrow_end].copy()
			
			# Add bakers 
			baker_times = data.iloc[wbrow : wbrow_end]

			# Combine them
			week_times = pd.concat([cashier_times, baker_times])
		else:
			week_times = data.iloc[wrow : wrow_end]		

		# Create an empty list to store the final tuples
		schedule_list = []

		# Iterate through every row in the week_times dataframe
		for index, row in week_times.iterrows():
			name = row[1]  # Column 0 (Excel column B) contains the employee names
			
			# Check if we have a valid name (skip empty rows filled with 0)
			if str(name) != 'nan' and name != 0:
				
				# Iterate over the column indices where we know the dates are (1 through 7)
				for col_idx in week_dates.index:
					shift = row[col_idx]
			
					# If the employee actually has a shift that day (not 0)
					if shift != 0 and str(shift) != 'nan':
						date_obj = week_dates[col_idx]

						# Convert the string/object to a reliable pandas datetime object
						dt_obj = pd.to_datetime(date_obj, dayfirst=True, errors='coerce')
						
						# Format the date and day (assuming date_obj is a datetime object)
						if pd.notna(dt_obj):
							day_name = dt_obj.strftime("%A").capitalize()  # e.g., "Monday"
							date_str = dt_obj.strftime("%d/%m/%Y")    # e.g., "03/03/2026"

							# Get badge
							badge_id = badges.get(name, "NOT FOUND")

							schedule_list.append((name, badge_id, day_name, date_str, shift, week))
	
		# Add shifts to database
		db.add_shifts(schedule_list, role, week)

	except Exception as e:
		print(f"Error initializing roster: {e}")

# --- Step 2: Clock Collection (Logic from att_clock_times.py) ---
def collect_clock_times():
	"""Reads last 5 files from Uniclox folder and saves to DB."""
	clock_times = []
	clock_files = [f for f in os.listdir(UNICLOX_FOLDER) if 'TL' in f and f[-7:-4] != '000']
	recent_files = clock_files[-5:]

	for filename in recent_files:
		with open(os.path.join(UNICLOX_FOLDER, filename), 'r') as f:
			for line in f:
				parts = line.strip().split(',')
				if len(parts) < 2: continue
				badge = parts[0]
				dt_obj = datetime.strptime(parts[1], '%Y-%m-%d %H:%M:%S')
				clock_times.append((badge, dt_obj.strftime("%d/%m/%Y"), dt_obj.strftime("%H:%M:%S")))

	db.add_clock_times(clock_times)

# --- Step 3: Write Roster Shifts to Excel (Logic from cas_clock_times.py) ---
def sync_shifts_to_excel(wb, sheet_name):
	"""
	1. Write roster data for name, badges, dates, shift times.
	2. Adds total headings under weekly shifts	
	"""
	# wb = load_workbook(WAGE_TIMES_FILE)
	ws = wb[sheet_name]

	# Get selected data from database
	if sheet_name == "Att Week One":
		data = db.get_shift_times_db('Attendant', 'WeekOne')
	elif sheet_name == "Att Week Two":
		data = db.get_shift_times_db('Attendant', 'WeekTwo')
	elif sheet_name == "Cashier Week One":
		data = db.get_shift_times_db('Cashier', 'WeekOne')
	else:
		data = db.get_shift_times_db('Cashier', 'WeekTwo')

	# Write shifts, badges, days, dates to Excel 
	current_row = 2
	prev_name = None
	prev_shift = []

	for row in data:
		name = row[0]
		badge = row[1]
		day = row[2]
		date = row[3]
		shift = row[4]

		# Skip two lines between employees
		if prev_name == None:
			pass
		elif prev_name != name:
			ws.cell(row=current_row, column=1, value=f'{prev_name.upper()} Total')
			current_row += 2
		# else:
		# 	ws.cell(row=current_row, column=1, value=f'{prev_name.upper()} Total')
		
		# Split shift times into start and end
		shift_times = split_roster_time(shift)
		shift_start = shift_times[0]
		shift_end = shift_times[1]	

		# Logic for night shift
		if shift_start == 0 and prev_shift >= 18:
			current_row -= 1
		
		elif shift_start >= 18:
			# Shift start
			ws.cell(row=current_row, column=1, value=name)
			ws.cell(row=current_row, column=2, value=badge)
			ws.cell(row=current_row, column=3, value=day)
			ws.cell(row=current_row, column=4, value=date)
			ws.cell(row=current_row, column=5, value=shift_start)
			ws.cell(row=current_row, column=6, value=0)

			# Get shift end and next day 
			# Convert the string into a datetime object (format: day/month/year)
			date_obj = datetime.strptime(date, "%d/%m/%Y")
		
			# Add one day
			next_date = date_obj + timedelta(days=1)
			new_date = next_date.strftime("%d/%m/%Y")

			# Get the day name (e.g., Thursday)
			day_name = next_date.strftime("%A")

			# Shift end
			ws.cell(row=current_row + 1, column=1, value=name)
			ws.cell(row=current_row + 1, column=2, value=badge)
			ws.cell(row=current_row + 1, column=3, value=day_name)
			ws.cell(row=current_row + 1, column=4, value=new_date)
			ws.cell(row=current_row + 1, column=5, value=0)
			ws.cell(row=current_row + 1, column=6, value=shift_end)

			current_row += 1

		else:
			ws.cell(row=current_row, column=1, value=name)
			ws.cell(row=current_row, column=2, value=badge)
			ws.cell(row=current_row, column=3, value=day)
			ws.cell(row=current_row, column=4, value=date)
			ws.cell(row=current_row, column=5, value=shift_start)
			ws.cell(row=current_row, column=6, value=shift_end)
		 
		current_row += 1
		prev_name = name
		prev_shift = shift_start

	ws.cell(row=current_row, column=1, value=f'{prev_name.upper()} Total')

# --- Step 3: Write Clocks to Excel ---
def sync_clocks_to_excel(wb, sheet_name):
	"""
	1. Matches clockings to the shift rows in the Excel sheet.
	2. ti = Time in | to = Time out [Actual clock times]
	3. Moves Sunday and Public holiday to right columns
	4. For cahsiers moves cashier/baker employee's cashiers times to right column
	"""

	clocks = db.get_clock_times()

	# wb = load_workbook(WAGE_TIMES_FILE)
	ws = wb[sheet_name]
	
	for i in range(2, ws.max_row + 1):
		badge = ws.cell(row=i, column=2).value
		date = ws.cell(row=i, column=4).value
		if not badge or not date: continue

		clocking_times = []

		for c_badge, c_date, c_time in clocks:
			if badge == c_badge and date == c_date:
				clocking_times.append(c_time)

		ti_roster = ws.cell(row=i, column=5).value
		to_roster = ws.cell(row=i, column=6).value

		# If shift is 'AF' ignore all clocking
		if ti_roster == 0.0 and to_roster == 0.0:
			continue

		# Handle night shift
		elif ti_roster == 18:
			t = time.fromisoformat(max(clocking_times)).strftime('%H:%M')
			ws.cell(row=i, column=7, value=t)

		# Handle morning of night shift
		elif to_roster == 6 or to_roster == 7:
			t = time.fromisoformat(min(clocking_times)).strftime('%H:%M')
			ws.cell(row=i, column=8, value=t)

		# Handle when employee clocks in/out only once 
		elif ti_roster > 0 and to_roster > 0 and len(clocking_times) == 1:
			# Logic for picking min/max based on shift
			# Single clocking: Determine if it's an IN or an OUT
			clock_h = int(clocking_times[0].split(':')[0])
			if abs(clock_h - (ti_roster or 0)) < abs(clock_h - (to_roster or 0)):
				t = time.fromisoformat(clocking_times[0]).strftime('%H:%M')
				ws.cell(row=i, column=7, value=t)
			else:
				t = time.fromisoformat(clocking_times[0]).strftime('%H:%M')
				ws.cell(row=i, column=8, value=t)
		
		# Handle when employee clocks in/out multiple times but does not clock in/out
		elif ti_roster > 0 and to_roster > 0 and len(clocking_times) >= 2:
			# Get clock times
			clock_min = (min(clocking_times).split(':')[0])
			clock_max = (max(clocking_times).split(':')[0])
			
			# Get roster shifts 
			roster_min = ws.cell(row=i, column=5).value
			roster_max = ws.cell(row=i, column=6).value

			# Check if the two clock times match
			if clock_min == clock_max:
				# Find if  clocks where start or end shift
				low = int(roster_min) - int(clock_min)
				high = int(roster_max) - int(clock_min)

				# Start shift
				if abs(low) < abs(high):
					t = time.fromisoformat(clocking_times[0]).strftime('%H:%M')
					ws.cell(row=i, column=7, value=t)
				else:
					t = time.fromisoformat(clocking_times[0]).strftime('%H:%M')
					ws.cell(row=i, column=8, value=t)
			# Handle normal shift
			else:
				ti = time.fromisoformat(min(clocking_times)).strftime('%H:%M')
				ws.cell(row=i, column=7, value=ti)

				to = time.fromisoformat(max(clocking_times)).strftime('%H:%M')
				ws.cell(row=i, column=8, value=to)


# --- Step 4: Calculate Hours (Logic from att_cal_hours.py) ---
def calculate_hours(wb, sheet_name):
	"""
	1. Calculates shift vs clocking hours
	2. Calculate total normal, sunday, public hours
	"""
	# wb = load_workbook(WAGE_TIMES_FILE)
	ws = wb[sheet_name]

	# Get public holidays
	holidays = db.get_public_holidays()

	# -- Caculate Shift vs Clocking Times To Get Hours Worked ---
	for i in range(2, ws.max_row + 1):
		name = ws.cell(row=i, column=1).value
		if not name or 'Total' in name: 
			continue

		day = ws.cell(row=i, column=3).value
		date = ws.cell(row=i, column=4).value
		ti = ws.cell(row=i, column=5).value  # Roster In
		to = ws.cell(row=i, column=6).value  # Roster Out
		ci = ws.cell(row=i, column=7).value  # Clock In (str HH:MM)
		co = ws.cell(row=i, column=8).value  # Clock Out (str HH:MM)

		# Checks to see if an employee did not clock
		if (ti and ti > 0 and not ci) or (to and to > 0 and not co):
			ws.cell(row=i, column=12, value="No Clock")
			continue

		# Rounding Logic
		calc_ti = adjust_time(ci, ti, day, True) if ci else 0
		calc_to = adjust_time(co, to, day, False) if co else 0

		# Night Shift Logic
		if ti == 18:
			hours = 24.0 - calc_ti
		elif ti == 0 and to > 0:
			hours = calc_to
		else:
			hours = calc_to - calc_ti

		# Assign columns
		# Get cashier dates
		if date in holidays:
			ws.cell(row=i, column=9, value='')
			ws.cell(row=i, column=11, value=hours)
		elif day == "Sunday":
			ws.cell(row=i, column=9, value='') 
			ws.cell(row=i, column=10, value=hours)
		elif sheet_name in ['Cashier Week One', 'Cashier Week Two']:
			bc = get_cashier_dates()
			for dy, dt in bc:
				if dy.upper() == name.upper() and dt == date:
					ws.cell(row=i, column=9, value='')
					ws.cell(row=i, column=13, value=hours)
				else:
					ws.cell(row=i, column=9, value=hours)
		else: 
			ws.cell(row=i, column=9, value=hours)



# ****** WORKING ******

# --- Step 5: Total Hours Worked ---
def cal_total_hours(wb, role="Attendant"):

	# Check what role is being calculated
	if role == "Attendant":
		sheets = ['Att Week One', 'Att Week Two'] 
	else:	
		sheets = ['Cashier Week One', 'Cashier Week Two']

	# Initilize dic
	# totals = {}	

	# Loop through sheets an calulate totals
	for sheet in sheets:
		ws = wb[sheet]

		totals = {}	
    
		# Iterate through rows (start at row 2 to skip headers)
		# Using ws.max_row + 1 to ensure the last person's total is written
		for row in range(2, ws.max_row + 2):
			name = ws.cell(row=row, column=1).value
			day = ws.cell(row=row, column=3).value

			# Determine if this is a "Total" row or an empty break row
			is_total_row = name and "Total" in str(name)
			# is_empty_row = name is None

			# If it's a normal day row, accumulate hours
			if name and not is_total_row:
				# Create name key in dic
				totals.setdefault(name, {'std': 0, 'sun': 0, 'pub': 0, 'nc': 0})
				
				# Accumulate values
				nc = ws.cell(row=row, column=12).value
				if nc is not None:
					totals[name]['nc'] = 1
				elif day == 'Sunday':
					totals[name]['sun'] += (ws.cell(row=row, column=10).value or 0)
				elif ws.cell(row=row, column=11).value is not None:
					totals[name]['pub'] += ws.cell(row=row, column=11).value
				else:
					totals[name]['std'] += (ws.cell(row=row, column=9).value or 0)
			
			elif name and is_total_row:
				# Get name without 'Total'
				name_total = ws.cell(row=row - 1, column=1).value

				# Add to total coloumn in excel
				ws.cell(row=row, column=9, value=totals[name_total]['std'])
				ws.cell(row=row, column=10, value=totals[name_total]['sun'])
				ws.cell(row=row, column=11, value=totals[name_total]['pub'])
				ws.cell(row=row, column=12, value=totals[name_total]['nc'])

		# for data in totals.items():
		# 	print(data[0])

		print(totals)

				

			
	

   








# --- Step 4: Formating Excel (Logic from att_cal_hours.py) ---
def format_excel(wb):
	# Get column configs
	cols_att = COLUMN_WIDTHS_ATT
	cols_tot = COLUMN_WIDTHS_TOTALS
	col_diff = COL_DIFF

	# Sheet names
	weekly_sheets = ['Att Week One', 'Att Week Two', 'Cashier Week One', 'Cashier Week Two']
	total_sheets = ['Att Total', 'Cashier Total']

	# Apply formats to sheets
	for sheet_name in weekly_sheets + total_sheets:
		if sheet_name not in wb.sheetnames: 
			continue
		
		ws = wb[sheet_name]

		if sheet_name in weekly_sheets:
			# Apply Column Widths
			for col, size in cols_att.items():
				ws.column_dimensions[col].width = size + col_diff
			
			# Style 'Total' rows
			style_cols = [1, 2, 9, 10, 11, 12, 13] if 'Cashier' in sheet_name else [1, 2, 9, 10, 11, 12]

			for row in range(2, ws.max_row + 1):
				if ws.cell(row=row, column=1).value and 'Total' in str(ws.cell(row=row, column=1).value):
					for c in style_cols:
						ws.cell(row=row, column=c).style = "total_style"
		
		else: # Logic for Total sheets
			# Apply Column Widths
			for col, size in cols_tot.items():
				ws.column_dimensions[col].width = size + col_diff
			
			# # Center Align columns B through F
			# for row in range(2, ws.max_row + 1):
			# 	for col_idx in range(2, 7):
			# 		ws.cell(row=row, column=col_idx).alignment = Alignment(horizontal='center')




# # --- Running functions ---

# 		# - Clear Excel -
# clear_excel()

	# - Load Workbook -
wb = load_excel()


cal_total_hours(wb)

# 		# - Clear database -
# db.clear_session_data()

# 		# - Send roster shift to db -
# roster_shift_to_db("Attendant", "WeekOne")
# roster_shift_to_db("Attendant", "WeekTwo")
# roster_shift_to_db("Cashier", "WeekOne")
# roster_shift_to_db("Cashier", "WeekTwo")

# 		# - Collect Clocks -
# collect_clock_times()

#  		# - Shifts -
# sync_shifts_to_excel(wb, 'Att Week One')
# sync_shifts_to_excel(wb, 'Att Week Two')
# sync_shifts_to_excel(wb, 'Cashier Week One')
# sync_shifts_to_excel(wb, 'Cashier Week Two')


# 		# - Clock -
# sync_clocks_to_excel(wb, 'Att Week One')
# sync_clocks_to_excel(wb, 'Att Week Two')
# sync_clocks_to_excel(wb, 'Cashier Week One')
# sync_clocks_to_excel(wb, 'Cashier Week Two')

# 		# - Calculate -
# calculate_hours(wb, 'Att Week One')
# calculate_hours(wb, 'Att Week Two')
# calculate_hours(wb, 'Cashier Week One')
# calculate_hours(wb, 'Cashier Week Two')

# format_excel(wb)

save_workbook(wb)


#  ---- WORKING AND ERRORS ---


# no clock = 1

# Formats

# Reculculate wages function

# total wages normal, sunday, public

# Add error handleing to processor functions

# --------------------------



# # --- Step 5: Carwash (Logic from carwash_db.py) ---
# def process_carwash(self):
#     wb = load_workbook(CARWASH_FILE, data_only=True)
#     ws = wb['Times']
#     data = []
#     for row in ws.iter_rows(min_row=3, max_row=10, min_col=12, max_col=16, values_only=True):
#         if row[0]: data.append((row[0], row[1], row[2], row[3], '0', '0'))
	
#     with self.db.get_connection() as con:
#         c = con.cursor()
#         c.executemany("INSERT INTO carwashTotal VALUES (?,?,?,?,?,?)", data)
#         con.commit()