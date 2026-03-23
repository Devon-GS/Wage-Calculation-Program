# This file stores your constants and styles


from openpyxl.styles import Alignment, Font, Border, Side

DB_PATH = "wageTimes.db"
WAGE_TIMES_FILE = "Wage Times.xlsx"
PAYROLL_FILE = "Payroll/Payroll.xlsx"

# Shared Styles
THIN_SIDE = Side(style='thin', color="000000")
DOUBLE_SIDE = Side(style='double', color="000000")

TOTAL_STYLE = {
    "font": Font(bold=True),
    "border": Border(top=THIN_SIDE, bottom=DOUBLE_SIDE)
}

COLUMN_WIDTHS_ATT = {'A':15, 'B':12, 'C':10, 'D':10, 'E':8, 'F':8, 'G':12, 'H':12, 'I':8, 'J':12, 'K':12, 'L':10}