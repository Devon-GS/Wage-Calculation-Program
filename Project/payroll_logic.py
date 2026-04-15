import xlwings as xw
import pandas as pd
from openpyxl import load_workbook
from config import PAYROLL_FILE


def update_payroll_from_db(self):
	# Logic to pull from attTotal/cashierTotal and write to Payroll.xlsx
	pass

def calculate_tax(self):
	app = xw.App(visible=False)
	try:
		book = app.books.open(PAYROLL_FILE)
		book.save()
		book.close()
		# Perform pandas tax bracket matching here
	finally:
		app.quit()