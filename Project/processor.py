import os
import re
import pandas as pd
import database as db
from datetime import datetime, timedelta, time
from openpyxl import load_workbook
from config import (WAGE_TIMES_FILE, PUBLIC_HOILIDAY_FILE, UNICLOX_FOLDER, 
					ATT_ROSTER_FILE, CAS_ROSTER_FILE, BADGE_NUMBER_FILE, CARWASH_FILE)


# --- Helper Functions ---
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

def get_public_holidays():
	holidays = []
	if os.path.exists(PUBLIC_HOILIDAY_FILE):
		wb = load_workbook(PUBLIC_HOILIDAY_FILE, data_only=True)
		ws = wb.active
		for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
			if row[0]: holidays.append(row[0].strftime('%d/%m/%y'))
		wb.close()
	return holidays

# --- Helper Functions End ---

# --- Step 1: Roster to Excel ---
def initialize_roster_to_excel(role="Attendant", week="WeekOne"):
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
	"""Reads last 20 files from Uniclox folder and saves to DB."""
	clock_times = []
	clock_files = [f for f in os.listdir(UNICLOX_FOLDER) if 'TL' in f and f[-7:-4] != '000']
	recent_files = clock_files[-20:]

	for filename in recent_files:
		with open(os.path.join(UNICLOX_FOLDER, filename), 'r') as f:
			for line in f:
				parts = line.strip().split(',')
				if len(parts) < 2: continue
				badge = parts[0]
				dt_obj = datetime.strptime(parts[1], '%Y-%m-%d %H:%M:%S')
				clock_times.append((badge, dt_obj.strftime("%d/%m/%y"), dt_obj.strftime("%H:%M:%S")))

	db.add_clock_times(clock_times)






# ****** WORKING ******

# --- Step 3: Match Clocks to Excel (Logic from cas_clock_times.py) ---
def sync_clocks_to_excel(sheet_name, role="Attendant"):
	"""
	1. Write roster data for name, badges, dates, shift times.
	2. Matches DB clockings to the rostered rows in the Excel sheet.
	"""

	# -- Write shifts to excel ---

	wb = load_workbook(WAGE_TIMES_FILE)
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

	for x in data:
		print(x)
	
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
			current_row += 2
		
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
		


		
	
















	# wb = load_workbook(WAGE_TIMES_FILE)
	# ws = wb[sheet_name]
	# table = "ClockTimeAttendent" if role == "Att" else "ClockTimeCashier"
	
	# with db.get_connection() as con:
	# 	c = con.cursor()
	# 	for i in range(2, ws.max_row + 1):
	# 		badge = ws.cell(row=i, column=2).value
	# 		date = ws.cell(row=i, column=4).value
	# 		if not badge or not date: continue

	# 		c.execute(f"SELECT time FROM {table} WHERE badge = ? AND date = ?", (str(badge), str(date)))
	# 		clocks = sorted([x[0] for x in c.fetchall()])
			
	# 		if not clocks: continue
			
	# 		ti_roster = ws.cell(row=i, column=5).value
	# 		to_roster = ws.cell(row=i, column=6).value

	# 		# Logic for picking min/max based on shift
	# 		if len(clocks) == 1:
	# 			# Single clocking: Determine if it's an IN or an OUT
	# 			clock_h = int(clocks[0].split(':')[0])
	# 			if abs(clock_h - (ti_roster or 0)) < abs(clock_h - (to_roster or 0)):
	# 				ws.cell(row=i, column=7, value=clocks[0][:5])
	# 			else:
	# 				ws.cell(row=i, column=8, value=clocks[0][:5])
	# 		else:
	# 			# Multiple clockings
	# 			ws.cell(row=i, column=7, value=clocks[0][:5]) # Earliest
	# 			ws.cell(row=i, column=8, value=clocks[-1][:5]) # Latest
	
	wb.save(WAGE_TIMES_FILE)




sync_clocks_to_excel('Att Week One')
sync_clocks_to_excel('Att Week Two')
sync_clocks_to_excel('Cashier Week One', 'Cashier')
sync_clocks_to_excel('Cashier Week Two', 'Cashier')























# # --- Step 4: Calculate Hours (Logic from att_cal_hours.py) ---
# def calculate_hours(self, sheet_name):
#     wb = load_workbook(WAGE_TIMES_FILE)
#     ws = wb[sheet_name]
#     holidays = self.get_public_holidays()

#     for i in range(2, ws.max_row + 1):
#         name = ws.cell(row=i, column=1).value
#         if not name or 'Total' in name: continue

#         ti = ws.cell(row=i, column=5).value  # Roster In
#         to = ws.cell(row=i, column=6).value  # Roster Out
#         ci = ws.cell(row=i, column=7).value  # Clock In (str HH:MM)
#         co = ws.cell(row=i, column=8).value  # Clock Out (str HH:MM)
#         day = ws.cell(row=i, column=3).value
#         date = ws.cell(row=i, column=4).value

#         if (ti and ti > 0 and not ci) or (to and to > 0 and not co):
#             ws.cell(row=i, column=12, value="No Clock")
#             continue

#         # Rounding Logic
#         calc_ti = self._adjust_time(ci, ti, True) if ci else 0
#         calc_to = self._adjust_time(co, to, False) if co else 0

#         # Special Night Shift Logic
#         if ti == 18:
#             hours = 24.0 - calc_ti
#         elif ti == 0 and to > 0:
#             hours = calc_to
#         else:
#             hours = calc_to - calc_ti

#         # Assign columns
#         if date in holidays: ws.cell(row=i, column=11, value=hours)
#         elif day == "Sunday": ws.cell(row=i, column=10, value=hours)
#         else: ws.cell(row=i, column=9, value=hours)

#     wb.save(WAGE_TIMES_FILE)

# def _adjust_time(self, clock_str, roster_h, is_in):
#     """The specific rounding logic from your scripts."""
#     if not clock_str: return float(roster_h)
#     h, m = map(int, clock_str.split(':'))
	
#     if is_in: # Clock In Logic
#         if h > roster_h or (h == roster_h and m > 0):
#             if m <= 15: return h + 0.25
#             elif m <= 30: return h + 0.50
#             elif m <= 45: return h + 0.75
#             else: return float(h + 1)
#         return float(roster_h)
#     else: # Clock Out Logic
#         if h < roster_h:
#             if m <= 4: return float(h)
#             elif m <= 15: return float(h)
#             elif m <= 30: return h + 0.25
#             elif m <= 45: return h + 0.50
#             else: return h + 0.75
#         return float(roster_h)

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
		










# ------- EXTRA ----------

# Assuming header=4 for Cashiers or header=1 for Attendants
	# hdr = 1 if role == "Att" else 4
	# cols = ['idx','ATTENDANTS', 'THURS', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED'] if role == "Att" \
	# 		else ['idx','CASHIERS', 'THU', 'FRI', 'SAT', 'SUN', 'MON', 'TUE', 'WED']

# --- to excel ---
# wb = load_workbook(WAGE_TIMES_FILE)
		# sheet_name = f"{role} Week One"
		# if sheet_name not in wb.sheetnames:
		# 	wb.create_sheet(sheet_name)
		# ws = wb[sheet_name]
		
		# 3. Write to Excel and map Badges
		# current_row = 2
		# for index, row in data.iterrows():
		# 	name = str(row[cols[1]]).strip()
		# 	badge_id = badges.get(name, "NOT FOUND") # <--- USING THE BADGE FILE HERE
		# 	print(name)
			
			# Write Name and Badge to columns A and B
			# ws.cell(row=current_row, column=1, value=name)
			# ws.cell(row=current_row, column=2, value=badge_id)
			
			# ... rest of your logic to fill dates and rostered times ...
			# current_row += 1
			
		# wb.save(WAGE_TIMES_FILE)

# ------- EXTRA END ----------




















# import os
# import re
# from datetime import datetime, time
# from openpyxl import load_workbook
# from config import WAGE_TIMES_FILE, PUBLIC_HOILIDAY_FILE



# class WageProcessor:
# 	def __init__(self, db_manager):
# 		self.db = db_manager

# 	def get_public_holidays(self):
# 		holidays = []
# 		try:
# 			wb = load_workbook(PUBLIC_HOILIDAY_FILE, data_only=True)
# 			ws = wb.active
# 			for row in ws.iter_rows(min_row=2, max_col=1, max_row=20, values_only=True):
# 				if row[0]:
# 					# Standardize format to d/m/y to match roster dates
# 					holidays.append(row[0].strftime('%d/%m/%y'))
# 			wb.close()
# 		except Exception as e:
# 			print(f"Error loading holidays: {e}")
# 		return holidays

# 	def _adjust_time(self, clock_val, roster_h, is_clock_in):
# 		"""
# 		Implements the 15-minute rounding logic from your original script.
# 		clock_val: string "HH:MM"
# 		roster_h: int (e.g., 8 or 18)
# 		is_clock_in: Boolean (True for clock in, False for clock out)
# 		"""
# 		if not clock_val:
# 			return float(roster_h)

# 		# Ensure clock_val is HH:MM string
# 		if not isinstance(clock_val, str):
# 			clock_val = clock_val.strftime("%H:%M")

# 		h = int(clock_val.split(":")[0])
# 		m = int(clock_val.split(":")[1])
# 		roster_time_str = f"{int(roster_h):02d}:00"

# 		if is_clock_in:
# 			# Logic: If clocked in LATER than rostered time, round UP
# 			if clock_val > roster_time_str:
# 				if m <= 15: return h + 0.25
# 				elif m <= 30: return h + 0.50
# 				elif m <= 45: return h + 0.75
# 				else: return float(h + 1)
# 			else:
# 				return float(roster_h)
# 		else:
# 			# Logic: If clocked out EARLIER than rostered time, round DOWN
# 			if clock_val < roster_time_str:
# 				if m <= 4: return float(h)
# 				elif m <= 15: return float(h)
# 				elif m <= 30: return h + 0.25
# 				elif m <= 45: return h + 0.50
# 				else: return h + 0.75
# 			else:
# 				return float(roster_h)

# 	def calculate_sheet_hours(self, sheet_name, role):
# 		print('yes')
# 		wb = load_workbook(WAGE_TIMES_FILE)
# 		if sheet_name not in wb.sheetnames:
# 			return
			
# 		ws = wb[sheet_name]
# 		holidays = self.get_public_holidays()

# 		for i in range(2, ws.max_row + 1):
# 			name = ws.cell(row=i, column=1).value
# 			# Skip empty rows or summary rows
# 			if not name or 'Total' in name:
# 				continue

# 			# Load values from Excel
# 			day = ws.cell(row=i, column=3).value
# 			date_str = ws.cell(row=i, column=4).value
# 			ti = ws.cell(row=i, column=5).value  # Roster In (int)
# 			to = ws.cell(row=i, column=6).value  # Roster Out (int)
# 			ci = ws.cell(row=i, column=7).value  # Clock In (str/time)
# 			co = ws.cell(row=i, column=8).value  # Clock Out (str/time)

# 			# 1. No Clock check
# 			if (ti and ti > 0 and not ci) or (to and to > 0 and not co):
# 				ws.cell(row=i, column=12, value="No Clock")
# 				continue

# 			# 2. Calculate adjusted hours
# 			hours_worked = 0.0

# 			if ti == 18:
# 				# Night Shift Logic: Only calculate hours from Clock In until Midnight (24:00)
# 				# The remaining hours (Midnight to 06:00) are usually handled by the next day's row
# 				tti = self._adjust_time(ci, ti, is_clock_in=True)
# 				hours_worked = 24.0 - tti
			
# 			elif ti == 0 and to > 0:
# 				# Part of Night Shift (Morning finish): 00:00 to clock out
# 				tto = self._adjust_time(co, to, is_clock_in=False)
# 				hours_worked = tto # Hours since midnight
				
# 			elif ti is not None and to is not None:
# 				# Normal Day Shift logic
# 				tti = self._adjust_time(ci, ti, is_clock_in=True)
# 				tto = self._adjust_time(co, to, is_clock_in=False)
# 				hours_worked = tto - tti

# 			# 3. Prevent negative hours (just in case)
# 			hours_worked = max(0, hours_worked)

# 			# 4. Assign to correct column (9: Normal, 10: Sunday, 11: Public Holiday)
# 			# Column 11: Public Holiday
# 			if date_str in holidays:
# 				ws.cell(row=i, column=11, value=hours_worked)
# 				ws.cell(row=i, column=9, value=None) # Clear normal column if it was filled
# 			# Column 10: Sunday
# 			elif day == "Sunday":
# 				ws.cell(row=i, column=10, value=hours_worked)
# 			# Column 9: Normal Weekday
# 			else:
# 				ws.cell(row=i, column=9, value=hours_worked)

# 		wb.save(WAGE_TIMES_FILE)
# 		wb.close()