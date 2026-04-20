import os
from pathlib import Path
from openpyxl import Workbook 
from openpyxl.styles import Alignment, Font, Border, Side,  NamedStyle, PatternFill

# --- SETUP PATHS ---

# Get payroll excel file - don't have to change name of file to payroll
def PAYROLL_FILE_LOC():
	cwd = Path(__file__).parent 
	excel_file = cwd / "Payroll"

	folder = Path(excel_file)
	
	# Search for Excel files and filter out hidden/temp files
	excel_files = folder.glob("*.xlsx*")
	valid_files =[file for file in excel_files if not file.name.startswith("~$")]
	
	# Check only one file in folder 
	if len(valid_files) == 1:
		return str(valid_files[0].resolve())	
	elif len(valid_files) == 0:
		return None
	else: 
		return None

DB_PATH = "wageTimes.db"
WAGE_TIMES_FILE = "Wage Times.xlsx"

# --- GET ALL STATIC DIRECTORIES --
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

BAKER_CASHIER_FILE = os.path.join(BASE_DIR, "Baker Cashier", "Baker Cashier Work.xlsx")
BADGE_NUMBER_FILE = os.path.join(BASE_DIR, "Badge Numbers", "Badges.xlsx")
PUBLIC_HOILIDAY_FILE = os.path.join(BASE_DIR, "Public Holidays", "Public Holidays.xlsx")
ROSTER_FOLDER = os.path.join(BASE_DIR, "Rosters")
UNICLOX_FOLDER = os.path.join(BASE_DIR, "Uniclox")
PUBLIC_HIOLIDAY_FILE = os.path.join(BASE_DIR, "Public Holidays", "Public Holidays.xlsx")
ATT_ROSTER_FILE = os.path.join(ROSTER_FOLDER, "Attendant_Carwash_Roster.xlsx")
CAS_ROSTER_FILE = os.path.join(ROSTER_FOLDER, "CASHIERS_ROSTER.xlsx")
CARWASH_FILE = os.path.join(BASE_DIR, "Carwash Times", "Carwash Times.xlsx")
CARWASH_HOURS_FILE = os.path.join(BASE_DIR, "Carwash Times", "Carwash Hours", "Carwash Hours.xlsx")
TAX_RATES_FILE = os.path.join(BASE_DIR, "Tax", "Tax_rates", "PAYE_Fortnight.xlsx")
TAX_RESULTS = os.path.join(BASE_DIR, "Tax", "tax_results.xlsx")
PAYROLL_FOLDER = os.path.join(BASE_DIR, "Payroll")
PAYSLIP_TEMPLATE = os.path.join(BASE_DIR, "Templates", "Payslip_Template.xlsx")
PAYSLIP_FOLDER = os.path.join(BASE_DIR, "Payslips")
COPY_FOLDER = os.path.join(BASE_DIR, "Copy Folder")

# --- STYLES FOR EXCEL ---
THIN_SIDE = Side(style='thin', color="000000")
DOUBLE_SIDE = Side(style='double', color="000000")

TOTAL_STYLE = NamedStyle(
	name = "total_style",
	font = Font(bold=True),
	border = Border(top=THIN_SIDE, bottom=DOUBLE_SIDE)	
)

COLUMN_WIDTHS_ATT = {'A':15.00, 'B':13.57, 'C':10.71, 'D':10.0, 'E':6.86, 'F':8.43, 'G':12.00, 'H':13.71,
					'I':5.43, 'J':12.43, 'K':11.29, 'L':8.00, 'M':12.39, 'N':14.83, 'O':13.61}
COLUMN_WIDTHS_TOTALS = {'A':11.00, 'B':17.57, 'C':17.43, 'D':23.71, 'E':11.11, 'F':12.39, 'G':14.83, 'H':13.61}

COL_DIFF = 0.78

# --- INITILIZE EXCEL WAGE TIMES WORKBOOK ---
def CREATE_EXCEL():
	if not os.path.isfile(WAGE_TIMES_FILE):
		wb = Workbook()

		# Add named style
		wb.add_named_style(TOTAL_STYLE)

		# Create the bold font style
		header_font = Font(bold=True, underline='single')

		# Remove and add sheets
		wb.remove(wb['Sheet'])
		sheet_list = ['Att Week One', 'Att Week Two', 'Att Total', 
				'Cashier Week One', 'Cashier Week Two', 'Cashier Total']
		
		for sheet in sheet_list:
			wb.create_sheet(sheet)
		
		# Create heading
		for ws in wb.worksheets:
			if ws in [wb['Att Total'], wb['Cashier Total']]:
				if ws == wb['Att Total']:
					ws["A1"] = 'Name'
					ws["B1"] = 'Total Normal Hours'
					ws["C1"] = 'Total Sunday Hours'
					ws["D1"] = 'Total Public Holiday Hours'
					ws["E1"] = 'No Clock'
				else:
					ws["A1"] = 'Name'
					ws["B1"] = 'Total Normal Hours'
					ws["C1"] = 'Total Sunday Hours'
					ws["D1"] = 'Total Public Holiday Hours'
					ws["E1"] = 'No Clock'
					ws["F1"] = "Cashier Hours"
					ws["G1"] = "C. Sunday Hours"
					ws["H1"] = "C. Public Hours"
			else:
				ws["A1"] = "Name"
				ws["B1"] = "Badge Number"
				ws["C1"] = "Week Day"
				ws["D1"] = "Date"
				ws["E1"] = "Time In"
				ws["F1"] = "Time Out"
				ws["G1"] = "Clock Time In"
				ws["H1"] = "Clock Time Out"
				ws["I1"] = "Hours"
				ws["J1"] = "Sunday Hours"
				ws["K1"] = "Public Hours"
				ws["L1"] = "No Clock"
				if ws in [wb['Cashier Week One'], wb['Cashier Week Two']]:
					ws["M1"] = 'Cashier Hours'
					ws["N1"] = 'C. Sunday Hours'
					ws["O1"] = 'C. Public Hours'

			# Bold headings
			for cell in ws[1]:
				cell.font = header_font

		wb.save(WAGE_TIMES_FILE)
		wb.close()

# --- INITILIZE EXCEL CARWASH TIMES WORKBOOK ---
def CREATE_CARWASH_TIMES():
	if not os.path.isfile(CARWASH_FILE):
		wb = Workbook()

		# Remove and add sheets
		wb.remove(wb['Sheet'])
		wb.create_sheet('Times')
		ws = wb['Times']

	# -- STYLES --
	left_alignment = Alignment(horizontal='left')
	bold_font = Font(bold=True, underline='single')
	center_alignment = Alignment(horizontal='center') 

	# Define the yellow fill (using the hex code for standard yellow)
	yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

	# Define a thin border for all four sides
	thin_border = Border(
	left=Side(style='thin', color='000000'),
	right=Side(style='thin', color='000000'),
	top=Side(style='thin', color='000000'),
	bottom=Side(style='thin', color='000000')
	)

	# Border line styles
	thick_side = Side(style='thick', color='000000')
	thin_side = Side(style='thin', color='000000')

	# -- --

	# Loop through the first row (1) from column 1 to 16
	for col_num in range(1, 17):
		# Assign the cell to a variable to make it cleaner to apply multiple styles
		cell = ws.cell(row=1, column=col_num)
		
		# Apply the font and the alignment
		cell.font = bold_font
		cell.alignment = center_alignment

	# CHANGE WIDTH OF COLUMNS
	# 2. Define your desired column widths in a dictionary
	column_widths = {
		'A': 9.72,
		'B': 8.83,
		'C': 10.42,
		'D': 10.52,
		'E': 8.83,
		'G': 9.72,
		'H': 8.83,
		'I': 10.42,
		'J': 10.52,
		'K': 8.83,
		'M': 9.72,
		'N': 12.90,
		'O': 11.83,
		'P': 11.94
	}

	# Loop through the dictionary and apply the widths
	for col_letter, width_value in column_widths.items():
		ws.column_dimensions[col_letter].width = width_value

	# Loop through rows starting from row 2 up to the last row with data
	for row in range(2, ws.max_row + 1):
		ws[f'B{row}'].alignment = left_alignment
		ws[f'H{row}'].alignment = left_alignment

	# Merge the cells foe extra time heading
	ws.merge_cells('M11:P11')
	ws['M11'].alignment = center_alignment

	# Align cells for extra time
	for row in ws['M12:P12']:
		for cell in row:
			cell.alignment = center_alignment

	# Range M2:P9 yellow highlight
	for row in ws['M2:P9']:
		for cell in row:
			cell.fill = yellow_fill


	# Range M2:P9 border with thick outside border
	for row in ws['M2:P9']:
		for cell in row:
			# Start by assuming the cell just needs regular thin inner borders
			top_border = thin_side
			bottom_border = thin_side
			left_border = thin_side
			right_border = thin_side
			
			# Check if the cell is on the TOP edge of our box
			if cell.row == 2:
				top_border = thick_side
				
			# Check if the cell is on the BOTTOM edge of our box
			if cell.row == 9:
				bottom_border = thick_side
				
			# Check if the cell is on the LEFT edge of our box
			if cell.column_letter == 'M':
				left_border = thick_side
				
			# Check if the cell is on the RIGHT edge of our box
			if cell.column_letter == 'P':
				right_border = thick_side
				
			# Apply the combined border to the cell
			cell.border = Border(top=top_border, bottom=bottom_border, left=left_border, right=right_border)

	# 1. Define your styles
	thick_side = Side(style='thick', color='000000')
	thin_side = Side(style='thin', color='000000')
	yellow_fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

	# ---------------------------------------------------------
	# TASK 1: Thick outside border around merged cells M11:P11
	# ---------------------------------------------------------
	for row in ws['M11:P11']:
		for cell in row:
			# Top and bottom are thick for the whole merged block
			top_border = thick_side
			bottom_border = thick_side
			
			# cell.column returns an integer: M is 13, P is 16
			left_border = thick_side if cell.column == 13 else None
			right_border = thick_side if cell.column == 16 else None
			
			cell.border = Border(top=top_border, bottom=bottom_border, left=left_border, right=right_border)


	# ---------------------------------------------------------
	# TASK 2 & 3: Borders for M12:P20 and Highlighting N & P
	# ---------------------------------------------------------
	for row in ws['M12:P20']:
		for cell in row:
			
			# --- BORDER LOGIC ---
			# Start by assuming normal (thin) borders inside the grid
			top_border = thin_side
			bottom_border = thin_side
			left_border = thin_side
			right_border = thin_side
			
			# Apply thick borders to the outside edges
			if cell.row == 12:
				top_border = thick_side
			if cell.row == 20:
				bottom_border = thick_side
			if cell.column == 13: # Column M
				left_border = thick_side
			if cell.column == 16: # Column P
				right_border = thick_side
				
			cell.border = Border(top=top_border, bottom=bottom_border, left=left_border, right=right_border)
			
			# --- HIGHLIGHT LOGIC ---
			# Highlight columns N (14) and P (16), but only for rows 13 through 20
			if cell.row >= 13 and cell.row <= 20:
				if cell.column == 14 or cell.column == 16: 
					cell.fill = yellow_fill
