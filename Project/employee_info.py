import customtkinter as ctk
from tkinter import messagebox
from CTkMessagebox import CTkMessagebox
from database import DatabaseManager


db = DatabaseManager()

def pop_up():
	top = ctk.CTkToplevel()
	top.attributes("-topmost", True)
	top.geometry("400x620")  
	top.title("Employee Management")
	top.configure(fg_color=("#f1f5f9", "#1e293b")) # Theme matching bg

	# Prevent the window from opening behind the main app
	top.after(100, top.lift)
	top.focus()

	# Layout Configuration
	top.grid_columnconfigure(1, weight=1)

	# --- INPUT FIELDS ---
	# Using a consistent font and padding
	label_font = ctk.CTkFont(family="Segoe UI", size=13, weight="bold")

	ctk.CTkLabel(top, text="English Name:", font=label_font).grid(row=0, column=0, padx=20, pady=(25, 5), sticky="w")
	ename_entry = ctk.CTkEntry(top, placeholder_text="e.g. John")
	ename_entry.grid(row=0, column=1, padx=20, pady=(25, 5), sticky="ew") 

	ctk.CTkLabel(top, text="Full Name:", font=label_font).grid(row=1, column=0, padx=20, pady=10, sticky="w")
	fname_entry = ctk.CTkEntry(top, placeholder_text="e.g. Johnathan")
	fname_entry.grid(row=1, column=1, padx=20, pady=10, sticky="ew")

	ctk.CTkLabel(top, text="Surname:", font=label_font).grid(row=2, column=0, padx=20, pady=10, sticky="w")
	sname_entry = ctk.CTkEntry(top, placeholder_text="e.g. Smith")
	sname_entry.grid(row=2, column=1, padx=20, pady=10, sticky="ew")

	ctk.CTkLabel(top, text="ID/Passport:", font=label_font).grid(row=3, column=0, padx=20, pady=10, sticky="w")
	id_entry = ctk.CTkEntry(top, placeholder_text="ID Number")
	id_entry.grid(row=3, column=1, padx=20, pady=10, sticky="ew")

	# --- FUNCTIONS ---

	# Add employee infornation
	def save():
		# Ask if sure
		msg = CTkMessagebox(title="Add Employee", 
				message="Are you sure you want to add an employee?",
				icon="question", 
				option_1="No", 
				option_2="Yes")
		
		# Get response
		response = msg.get()

		if response == "Yes":
			english_name = ename_entry.get().capitalize()
			full_name = fname_entry.get().capitalize()
			surname = sname_entry.get().capitalize()
			id_pass = id_entry.get()

			# Add employee information
			db.add_employees(english_name, full_name, surname, id_pass)

			# Clear entry boxes
			ename_entry.delete(0, ctk.END)
			fname_entry.delete(0, ctk.END)
			sname_entry.delete(0, ctk.END)
			id_entry.configure(state="normal") 
			id_entry.delete(0, ctk.END)
	
		else:
			CTkMessagebox(title="Add Employee", 
				message="Operation Canceled",
				icon="cancel")

	# Search employee
	def search():
		# Clear entry boxes
		ename_entry.delete(0, ctk.END)
		fname_entry.delete(0, ctk.END)
		sname_entry.delete(0, ctk.END)
		id_entry.configure(state="normal")
		id_entry.delete(0, ctk.END)

		# Start pop up
		etop = ctk.CTkToplevel()
		etop.attributes("-topmost", True)
		etop.geometry("250x200")
		etop.title("Edit Employee")
		etop.configure(fg_color=("#f1f5f9", "#1e293b"))

		# Prevent the window from opening behind the main app
		etop.after(100, etop.lift)
		etop.focus()

		# Configure the grid column weight
		etop.columnconfigure(0, weight=1) 

		# Get all Employee names
		results = db.search_employees()
		
		# Make options fro drop down
		options = list(results.keys())

		# Function to handle the selection
		def select_employee():
			try:
				choice = option_menu.get()
				empolyee = db.employee_selected_option(results[choice])

				ename_entry.insert(0, empolyee[0][0])
				fname_entry.insert(0, empolyee[0][1])
				sname_entry.insert(0, empolyee[0][2])
				id_entry.insert(0, empolyee[0][3])
				id_entry.configure(state="readonly") 

				etop.destroy()
			except KeyError:
				CTkMessagebox(title="Error", message='Please Select a Valid Employee', icon="cancel")			

        # OptionMenu
		option_menu = ctk.CTkOptionMenu(etop, 
            values=options, 
            # command=optionmenu_callback,                
            fg_color="#4f46e5",                         
            button_color="#4338ca",                    
            button_hover_color="#3730a3"               
        )

		option_menu.grid(row=0, column=0, sticky="ew", padx=20, pady=(20, 5))

        # Default starting value
		option_menu.set("Please Select") 

		ctk.CTkButton(etop, text="Yes", fg_color="#10b981", hover_color="#059669", 
					font=label_font, command=select_employee).grid(row=2, column=0, columnspan=2, sticky="ew", padx=20, pady=(20, 5))
		
		
		ctk.CTkButton(etop, text="Exit", fg_color="#ef4444", hover_color="#b91c1c", 
					font=label_font, command=etop.destroy).grid(row=3, column=0, columnspan=2, sticky="ew", padx=20, pady=(20, 5))

	def update():
		# Ask if sure
		msg = CTkMessagebox(title="Update Employee", 
				message="Are you sure you want to update the employee?",
				icon="question", 
				option_1="No", 
				option_2="Yes")
		
		# Get response
		response = msg.get()


		if response == 'Yes':
			english_name = ename_entry.get().capitalize()
			full_name = fname_entry.get().capitalize()
			surname = sname_entry.get().capitalize()
			id_pass = id_entry.get()

			db.update_employees(english_name, full_name, surname, id_pass)

			ename_entry.delete(0, ctk.END)
			fname_entry.delete(0, ctk.END)
			sname_entry.delete(0, ctk.END)
			id_entry.configure(state="normal")
			id_entry.delete(0, ctk.END)
		else:
			CTkMessagebox(title="Update Employee", 
				message="Operation Canceled",
				icon="cancel")

	def delete():
		# Ask if sure
		msg = CTkMessagebox(title="Delete Employee", 
				message="Are you sure you want to delete the employee?",
				icon="question", 
				option_1="No", 
				option_2="Yes")
		
		# Get response
		response = msg.get()

		if response == 'Yes':
			id_pass = id_entry.get()

			db.delete_employees(id_pass)

			ename_entry.delete(0, ctk.END)
			fname_entry.delete(0, ctk.END)
			sname_entry.delete(0, ctk.END)
			id_entry.configure(state="normal")
			id_entry.delete(0, ctk.END)
		else:
			CTkMessagebox(title="Delete Employee", 
				message="Operation Canceled",
				icon="cancel")

	def clear():
		ename_entry.delete(0, ctk.END)
		fname_entry.delete(0, ctk.END)
		sname_entry.delete(0, ctk.END)
		id_entry.configure(state="normal")
		id_entry.delete(0, ctk.END)

	def bulk_add():
		# Ask if sure
		msg = CTkMessagebox(title="Add Bulk Employees", 
				message="Are you sure you want add bulk employees?",
				icon="question", 
				option_1="No", 
				option_2="Yes")
		
		# Get response
		response = msg.get()

		if response == 'Yes':
			db.bulk_add_employees()
		else:
			CTkMessagebox(title="Bulk Add Employees", 
				message="Operation Canceled",
				icon="cancel")
			
	# --- ACTION BUTTONS ---
	ctk.CTkButton(top, text="Add New Employee", fg_color="#10b981", hover_color="#059669", 
					font=label_font, command=save).grid(row=4, column=0, columnspan=2, sticky="ew", padx=20, pady=(20,30))

	ctk.CTkButton(top, text="Search by English Name", fg_color="#4f46e5", hover_color="#4338ca", 
					command=search).grid(row=5, column=0, columnspan=2, sticky="ew", padx=20, pady=5)

	ctk.CTkButton(top, text="Update Employee Details", fg_color="#4f46e5", hover_color="#4338ca", 
					command=update).grid(row=6, column=0, columnspan=2, sticky="ew", padx=20, pady=(10,30))

	ctk.CTkButton(top, text="Delete Employee", fg_color="#ef4444", hover_color="#b91c1c", 
					command=delete).grid(row=7, column=0, columnspan=2, sticky="ew", padx=20, pady=(5,30))

	ctk.CTkButton(top, text="Bulk Add (CSV)", fg_color="transparent", border_width=1,
			   		command=bulk_add).grid(row=8, column=0, columnspan=2, sticky="ew", padx=20, pady=(5, 20))

	ctk.CTkButton(top, text="Clear Form", fg_color="transparent", border_width=1, text_color=("#1e293b", "#cbd5e1"),
					command=clear).grid(row=9, column=0, columnspan=2, sticky="ew", padx=20, pady=(5, 20))