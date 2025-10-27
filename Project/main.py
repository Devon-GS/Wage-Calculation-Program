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
            5) Uniclox Files
            ''')
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