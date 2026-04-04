import os
from openpyxl import Workbook 
from openpyxl.styles import Alignment, Font, Border, Side

# SETUP PATHS
DB_PATH = "wageTimes.db"
WAGE_TIMES_FILE = "Wage Times.xlsx"
PAYROLL_FILE = "Payroll/Payroll.xlsx"

# Get the directory where main.py is located
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

# Check if wage time and create
if not os.path.isfile(WAGE_TIMES_FILE):
	wb = Workbook()
	wb.remove(wb['Sheet'])
	sheet_list = ['Att Week One', 'Att Week Two', 'Att Total', 
			   'Cashier Week One', 'Cashier Week Two', 'Cashier Total']
	
	for sheet in sheet_list:
		wb.create_sheet(sheet)

	
	# Create heading
	for ws in wb.worksheets:
		if ws not in [wb['Att Total'], wb['Cashier Total']]:
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

	wb.save(WAGE_TIMES_FILE)
	wb.close()

# Shared Styles
THIN_SIDE = Side(style='thin', color="000000")
DOUBLE_SIDE = Side(style='double', color="000000")

TOTAL_STYLE = {
	"font": Font(bold=True),
	"border": Border(top=THIN_SIDE, bottom=DOUBLE_SIDE)
}

COLUMN_WIDTHS_ATT = {'A':15, 'B':12, 'C':10, 'D':10, 'E':8, 'F':8, 'G':12, 'H':12, 'I':8, 'J':12, 'K':12, 'L':10}