import os
from openpyxl import Workbook 
from openpyxl.styles import Alignment, Font, Border, Side,  NamedStyle

# --- SETUP PATHS ---
DB_PATH = "wageTimes.db"
WAGE_TIMES_FILE = "Wage Times.xlsx"
PAYROLL_FILE = "Payroll/Payroll.xlsx"

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

# --- STYLES FOR EXCEL ---
THIN_SIDE = Side(style='thin', color="000000")
DOUBLE_SIDE = Side(style='double', color="000000")

TOTAL_STYLE = NamedStyle(
	name = "TOTAL_STYLE",
	font = Font(bold=True),
	border = Border(top=THIN_SIDE, bottom=DOUBLE_SIDE)	
)

COLUMN_WIDTHS_ATT = {'A':15.00, 'B':13.57, 'C':10.71, 'D':10.0, 'E':6.86, 'F':8.43, 'G':12.00, 'H':13.71, 'I':5.43, 'J':12.43, 'K':11.29, 'L':8.00, 'M':6.68}
COLUMN_WIDTHS_TOTALS = {'A':11.00, 'B':16.44, 'C':16.00, 'D':21.78, 'E':11.11, 'F':17.11}

COL_DIFF = 0.78

# --- INITILIZE EXCEL WORKBOOK ---
def CREATE_EXCEL():
	if not os.path.isfile(WAGE_TIMES_FILE):
		wb = Workbook()

		# Add named style
		wb.add_named_style(TOTAL_STYLE)

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
					ws["F1"] = 'Baker/Cashier Hours'
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
					ws["M1"] = 'Cashier'

		wb.save(WAGE_TIMES_FILE)
		wb.close()