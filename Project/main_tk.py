from tkinter import *
import os

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
        

# WIDGETS
# Setup buttons
setup_label = Label(root, text='SETUP',borderwidth=1, relief='solid')

setup_button = Button(root, text='Badge Numbers', width=12, command=lambda: setup_options(1))
setup_button2 = Button(root, text='Cashier/Baker', width=12, command=lambda: setup_options(2))
setup_button3 = Button(root, text='Public Holidays', width=12, command=lambda: setup_options(3))
setup_button4 = Button(root, text='Rosters', width=12, command=lambda: setup_options(4))

# Run Program Buttons
program_label = Label(root, text='RUN PROGRAM',borderwidth=1, relief='solid')

program_button = Button(root, text='First Time', width=12, command=lambda: setup_options(1))
program_button2 = Button(root, text='Run Wages', width=12, command=lambda: setup_options(1))
program_button3 = Button(root, text='Recalulate Wages', width=12, command=lambda: setup_options(1))
program_button4 = Button(root, text='Carwash Times', width=12, command=lambda: setup_options(1))


# BIND WIDGETS
# Setup Buttons
setup_label.grid(row=0, column=0, columnspan=4 ,sticky=W+E ,pady=(0,10))

setup_button.grid(row=1, column=0, padx=(5,10))
setup_button2.grid(row=1, column=1, padx=(0,10))
setup_button3.grid(row=1, column=2, padx=(0,10))
setup_button4.grid(row=1, column=3, padx=(0,5))

# Run Program Buttons
program_label.grid(row=2, column=0, columnspan=4 ,sticky=W+E ,pady=(10,10))

program_button.grid(row=3, column=0, padx=(5,10))
program_button2.grid(row=3, column=1, padx=(5,10))
program_button3.grid(row=3, column=2, padx=(5,10))
program_button4.grid(row=3, column=3, padx=(5,10))


# ROOT WINDOW CONFIG
root.title('Wage Calculator')
# root.iconbitmap('icons/smoking.ico')
root.geometry('420x360')
# root.columnconfigure(0, weight=1)

# RUN WINDOW
root.mainloop()