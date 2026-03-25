import os
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


# Shared Styles
THIN_SIDE = Side(style='thin', color="000000")
DOUBLE_SIDE = Side(style='double', color="000000")

TOTAL_STYLE = {
	"font": Font(bold=True),
	"border": Border(top=THIN_SIDE, bottom=DOUBLE_SIDE)
}

COLUMN_WIDTHS_ATT = {'A':15, 'B':12, 'C':10, 'D':10, 'E':8, 'F':8, 'G':12, 'H':12, 'I':8, 'J':12, 'K':12, 'L':10}