import xlwings as xw
import pandas as pd
from CTkMessagebox import CTkMessagebox
from openpyxl import load_workbook
from config import PAYROLL_FILE
import database as db


def run_payroll():
	try:
		wb = load_workbook(PAYROLL_FILE)
		ws = wb['Wages']

		# Get all total hours
		total_records =  db.get_total_hours()

		# 1. Pre-process records into a dictionary
		records_dict = {str(r[1]): r for r in total_records}

		# Iterate over payroll columns starting from column C
		columns = ws.iter_cols(min_row=1, min_col=3)

		for col in columns:
			# col[1] is Row 2 in Excel (0-indexed tuple)
			badge_val = col[1].value 
			
			if badge_val is None:
				continue

			# Convert to string and clean it up
			badge_str = str(badge_val).strip().lower()

			# 2. Figure out the badge number and the match type
			if badge_str.endswith('c'):
				badge_key = badge_str[:-1]
				match_type = 'c'
			elif badge_str.endswith('b'):
				badge_key = badge_str[:-1]
				match_type = 'b'
			else:
				# Strips the .0 if it exists so it matches r[1]
				try:
					badge_key = str(int(float(badge_val)))
				except ValueError:
					badge_key = badge_str
				match_type = 'exact'

			# 3. Look up the record
			r = records_dict.get(badge_key)

			# Skip to the next column if no matching record is found
			if not r:
				continue  

			# Asign to variables
			norm = float(r[2])
			sun = float(r[3])
			pub = float(r[4])

			if int(badge_key) > 1000:
				extra_pay = int(r[5])

			# 4. Apply the payment logic
			# Baker's cashier hours
			if match_type == 'c':
				cnorm = float(r[6])
				csun = float(r[7])
				cpub = float(r[8])

				# Handle 'Null' to prevent ValueError crashes
				col[2].value  = cnorm
				col[11].value = csun
				col[14].value = cpub

			# Baker's baker hours
			elif match_type == 'b':
				col[2].value  = norm
				col[11].value = sun
				col[14].value = pub

			# Carwash hours 
			elif match_type == 'exact' and int(badge_key) > 1000:
				col[2].value  = norm
				col[11].value = sun
				col[14].value = pub
				col[20].value = extra_pay

			# Every one else
			else:
				col[2].value  = norm
				col[11].value = sun
				col[14].value = pub	
		
		wb.save(PAYROLL_FILE)
		wb.close()
	except Exception as error:
		CTkMessagebox(title="Error", message=str(error), icon="cancel")





def calculate_tax(self):
	
	app = xw.App(visible=False)
	try:
		book = app.books.open(PAYROLL_FILE)
		book.save()
		book.close()
		# Perform pandas tax bracket matching here
	finally:
		app.quit()