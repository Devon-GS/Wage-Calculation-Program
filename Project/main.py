from tkinter import *
from tkinter import messagebox
import os
import shutil
import gen_payslips as pay 
import functions as f

root = Tk()

# FUNCTIONS
def setup_options(button_id):
	if button_id == 1:
		os.system('start "EXCEL.EXE" "Badge Numbers/Badges.xlsx"')
	elif button_id == 2:
		os.system('start "EXCEL.EXE" "Baker Cashier/Baker Cashier Work.xlsx"')
	elif button_id == 3:
		os.system('start "EXCEL.EXE" "Public Holidays/Public Holidays.xlsx"')
	elif button_id == 4:
		os.startfile('Rosters')
	elif button_id == 5:
		os.system('start "EXCEL.EXE" "Wage Times.xlsx"')
	elif button_id == 6:
		os.startfile("Uniclox")
	elif button_id == 7:
		os.system('start "EXCEL.EXE" "Payroll/Payroll.xlsx"')
	elif button_id == 8:
		os.startfile('Payroll')
	elif button_id == 9:
		response = messagebox.askyesno('Employee Information', 'Do you want to update employee infomation?')
		if response == 1:
			top = Toplevel()
			top.attributes("-topmost", True)
			top.geometry("305x355")
			top.title("Add Employee Information")

			ename_label = Label(top, text="English Name:")
			ename_label.grid(row=0, column=0, padx=5, pady=10)
			ename_entry = Entry(top, width=30)
			ename_entry.grid(row=0, column=1, columnspan=2, sticky='EW', padx=5, pady=10) 

			fname_label = Label(top, text="Full Name:")
			fname_label.grid(row=1, column=0, padx=5, pady=5)
			fname_entry = Entry(top)
			fname_entry.grid(row=1, column=1, columnspan=2, sticky='EW', padx=5, pady=5)

			sname_label = Label(top, text="Surname Name:")
			sname_label.grid(row=2, column=0, padx=5, pady=5)
			sname_entry = Entry(top)
			sname_entry.grid(row=2, column=1, columnspan=2, sticky='EW', padx=5, pady=5)

			id_label = Label(top, text="ID/Passport:")
			id_label.grid(row=3, column=0, padx=5, pady=5)
			id_entry = Entry(top)
			id_entry.grid(row=3, column=1, columnspan=2, sticky='EW', padx=5, pady=5)

			def save():
				response = messagebox.askyesno('Add Employee', 'Are you sure you want to add an employee?')
				if response == 1:
					# Get entry information
					english_name = ename_entry.get().capitalize()
					full_name = fname_entry.get().capitalize()
					surname = sname_entry.get().capitalize()
					id_pass = id_entry.get()

					# Save to database
					pay.add_employees(english_name, full_name, surname, id_pass)

					# clear entry boxes
					ename_entry.delete(0, END)
					fname_entry.delete(0, END)
					sname_entry.delete(0, END)
					id_entry.delete(0, END)
				else:
					messagebox.showinfo('Add Employee', 'Nothing happened!')

			def search():
				# Get search name
				english_name = ename_entry.get()
				ename_entry.delete(0, END)

				# Get matching results
				search_results = pay.search_employees(english_name)

				# Loop through and select right result
				for x in search_results:
					response = messagebox.askyesno('Employee Information', f'{x} : Is this the right employee?')

					if response == 1:
						ename_entry.insert(0, x[0])
						fname_entry.insert(0, x[1])
						sname_entry.insert(0, x[2])
						id_entry.insert(0, x[3])
						break
				id_entry.config(state="readonly")
			
			def update():
				response = messagebox.askyesno('Update Employee', 'Are you sure you want to update the employee?')
				if response == 1:
					# Get entry information
					english_name = ename_entry.get().capitalize()
					full_name = fname_entry.get().capitalize()
					surname = sname_entry.get().capitalize()
					id_pass = id_entry.get()

					# Save to database
					pay.update_employees(english_name, full_name, surname, id_pass)

					# clear entry boxes
					ename_entry.delete(0, END)
					fname_entry.delete(0, END)
					sname_entry.delete(0, END)
					id_entry.config(state="normal")
					id_entry.delete(0, END)
					
				else:
					messagebox.showinfo('Update Employee', 'Nothing happened!')
			
			def delete():
				response = messagebox.askyesno('Delete Employee', 'Are you sure you want to delete the employee?')
				if response == 1:
					# Get entry information
					id_pass = id_entry.get()

					# Delete employee
					pay.delete_employees(id_pass)

					# clear entry boxes
					ename_entry.delete(0, END)
					fname_entry.delete(0, END)
					sname_entry.delete(0, END)
					id_entry.config(state="normal")
					id_entry.delete(0, END)
				else:
					messagebox.showinfo('Delete Employee', 'Nothing happened!')

			def clear():
				# clear entry boxes
				ename_entry.delete(0, END)
				fname_entry.delete(0, END)
				sname_entry.delete(0, END)
				id_entry.config(state="normal")
				id_entry.delete(0, END)
		
			# Buttons 
			save_button = Button(top, text="Add", command=save)
			save_button.grid(row=4, column=0, columnspan=3, sticky=EW, padx=5, pady=5)

			search_button = Button(top, text="Search", command=search)
			search_button.grid(row=5, column=0, columnspan=3, sticky=EW, padx=5, pady=5)

			update_button = Button(top, text="Update", command=update)
			update_button.grid(row=6, column=0, columnspan=3, sticky=EW, padx=5, pady=5)
			
			delete_button = Button(top, text="Delete", command=delete)
			delete_button.grid(row=7, column=0, columnspan=3, sticky=EW, padx=5, pady=5)

			bulk_button = Button(top, text="Bulk Add", command=pay.bulk_add)
			bulk_button.grid(row=8, column=0, columnspan=3, sticky=EW, padx=5, pady=5)

			clear_button = Button(top, text="Clear", command=clear)
			clear_button.grid(row=9, column=0, columnspan=3, sticky=EW, padx=5, pady=5)
		else:
			os.startfile('Templates')

def program_options(button_id):
	try:
		if button_id == 1:
			f.program_init()
			messagebox.showinfo('Run Wages', 'Program Setup Complete!')
		elif button_id == 2:
			response = messagebox.askyesno('Run Wages Check list', '''
								  Have you Checked the following:
								  1) Badge Number
								  2) Cashier/Baker Times
								  3) Public Holidays
								  4) Rosters
								  5) Template Updated
								  5) Uniclox Files''')
			if response == 1:
				f.wages_time_main_program()
				messagebox.showinfo('Run Wages', 'Wage Hours Completed!')
			else:
				messagebox.showinfo('Run Wages', 'Nothing Happened!')
		elif button_id == 3:
			f.recal_hours()
			messagebox.showinfo('Run Wages', 'Wage Hour Recalculation Complete')       
		elif button_id == 4:
			os.system('start "EXCEL.EXE" "Carwash Times/Carwash Times.xlsx"')
		elif button_id == 5:
			response = messagebox.askyesno('Run Payroll', 'Are you sure you want to run payroll?')
			if response == 1:
				f.run_payroll()               
				messagebox.showinfo('Run Payroll', 'Payroll Completed!')
			else:
				messagebox.showinfo('Run payroll', 'Nothing Happened!') 
		elif button_id == 6:
			shutil.copy2('Wage Times.xlsx', 'Copy Folder/Wage Times.xlsx')
			shutil.copy2('Payroll/Payroll.xlsx', 'Copy Folder/Payroll.xlsx')
			shutil.copy2('Rosters/Attendant_Carwash_Roster.xlsx', 'Copy Folder/Attendant_Carwash_Roster.xlsx')
			shutil.copy2('Rosters/CASHIERS_ROSTER.xlsx', 'Copy Folder/CASHIERS_ROSTER.xlsx')
			shutil.copy2('Carwash Times/Carwash Times.xlsx', 'Copy Folder/Carwash Times.xlsx')
			shutil.copy2('Tax/tax_results.xlsx', 'Copy Folder/tax_results.xlsx')

			os.startfile("Copy Folder")
		elif button_id == 7:
			response = messagebox.askyesno('Calculate Tax', 'Are you sure you want to calculate the tax?')
			if response == 1:
				f.calulate_tax()              
				messagebox.showinfo('Calculate Tax', 'Tax Calculation Completed!')
			else:
				messagebox.showinfo('Calculate Tax', 'Nothing Happened!') 
		elif button_id == 8:
			pay.gen_payslips()
			os.startfile("Payslips")
			 
	except Exception as error:
		messagebox.showerror('Run Wages', error)

# WIDGETS
# Setup buttons
setup_label = Label(root, text='SETUP',borderwidth=1, relief='solid')

setup_button = Button(root, text='Badge Numbers', width=12, command=lambda: setup_options(1))
setup_button2 = Button(root, text='Cashier/Baker', width=12, command=lambda: setup_options(2))
setup_button3 = Button(root, text='Public Holidays', width=12, command=lambda: setup_options(3))
setup_button4 = Button(root, text='Rosters', width=12, command=lambda: setup_options(4))
setup_button5 = Button(root, text='Open Templates', width=12, command=lambda: setup_options(9))
setup_button6 = Button(root, text='Open Uniclox', width=12, command=lambda: setup_options(6))

# Run Program Buttons
program_label = Label(root, text='RUN PROGRAM',borderwidth=1, relief='solid')

program_button = Button(root, text='First Time', width=12, command=lambda: program_options(1))
program_button2 = Button(root, text='Run Wages', width=12, command=lambda: program_options(2))
program_button3 = Button(root, text='Recalulate Wages', width=12, command=lambda: program_options(3))
program_button4 = Button(root, text='Carwash Times', width=12, command=lambda: program_options(4))

# Open Wage Times Excel
open_wage_button = Button(root, text='Open Wage Times', width=12, command=lambda: setup_options(5))

# Open Payroll Folder
payroll_open_folder_button = Button(root, text='Open Payroll Folder', width=12, command=lambda: setup_options(8))

# Open Payroll Excel
payroll_open_button = Button(root, text='Open Payroll Sheet', width=12, command=lambda: setup_options(7))

# Run Payroll
payroll_button = Button(root, text='Run Payroll', width=12, command=lambda: program_options(5))

# Run Tax
calculate_tax_button = Button(root, text='Calculate Tax', width=12, command=lambda: program_options(7))

# Generate Payslips
payslips = Button(root, text='Generate Payslips', width=12, command=lambda: program_options(8))

# Copy Button
copy_button = Button(root, text='Copy Sheets for Saving', width=12, command=lambda: program_options(6))

# BIND WIDGETS
# Setup Buttons
setup_label.grid(row=0, column=0, columnspan=4 ,sticky=W+E, padx=(5,5), pady=(0,10))

setup_button.grid(row=1, column=0, padx=(5,10))
setup_button2.grid(row=1, column=1, padx=(0,10))
setup_button3.grid(row=1, column=2, padx=(0,10))
setup_button4.grid(row=1, column=3, padx=(0,5))
setup_button5.grid(row=2, column=0, columnspan=2 ,sticky=W+E, padx=(5,5) ,pady=(10,10))
setup_button6.grid(row=2, column=2, columnspan=2 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# Run Program Buttons
program_label.grid(row=3, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

program_button.grid(row=4, column=0, padx=(5,10))
program_button2.grid(row=4, column=1, padx=(5,10))
program_button3.grid(row=4, column=2, padx=(5,10))
program_button4.grid(row=4, column=3, padx=(5,10))

# Open Wage Times.xlsx
open_wage_button.grid(row=5, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# Open Payroll Folder
payroll_open_folder_button.grid(row=7, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# Open Payroll Excel
payroll_open_button.grid(row=8, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# Run Payroll
payroll_button.grid(row=9, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# Run Tax
calculate_tax_button.grid(row=10, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# Generate Payslips
payslips.grid(row=11, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# Copy Button
copy_button.grid(row=12, column=0, columnspan=4 ,sticky=W+E, padx=(5,5) ,pady=(10,10))

# ROOT WINDOW CONFIG
root.title('Wage Calculator')
# root.iconbitmap('icons/smoking.ico')
root.geometry('440x490')
# root.columnconfigure(0, weight=1)

# RUN WINDOW
root.mainloop()