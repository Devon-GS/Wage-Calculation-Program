import os
from pathlib import Path
from openpyxl import Workbook 
from openpyxl.styles import Alignment, Font, Border, Side,  NamedStyle

# --- SETUP PATHS ---

# Get payroll excel file - don't have to change name of file to payroll
def get_payroll_path(relative_folder_path):
	folder = Path(relative_folder_path)
	
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

cwd = Path(__file__).parent 
excel_file = cwd / "Payroll" 
payroll_path = get_payroll_path(excel_file)

PAYROLL_FILE = payroll_path
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
TAX_RATES_FILE = os.path.join(BASE_DIR, "Tax", "Tax_rates", "PAYE_Fortnight.xlsx")
TAX_RESULTS = os.path.join(BASE_DIR, "Tax", "tax_results.xlsx")

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

# --- INITILIZE EXCEL WORKBOOK ---
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