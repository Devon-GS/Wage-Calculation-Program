# The GUI that connects everything


import customtkinter as ctk
from tkinter import messagebox
from CTkMessagebox import CTkMessagebox
import os
import traceback

# Import your logic classes
from database import DatabaseManager
from processor import WageProcessor
from payroll_logic import PayrollManager
from payslips import PayslipGenerator

# Set the visual theme
ctk.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
ctk.set_default_color_theme("blue")  # Themes: "blue" (standard), "green", "dark-blue"

class WageApp(ctk.CTk):
	def __init__(self):
		super().__init__()

		# --- Configuration ---
		self.title("Bracken Hill Fuel Wages Caclulator v2.0")
		self.geometry("900x600")
		
		# Initialize Logic
		self.db = DatabaseManager()
		self.processor = WageProcessor(self.db)
		self.payroll = PayrollManager(self.db)
		self.payslips = PayslipGenerator(self.db)

		# Create Layout
		self.grid_columnconfigure(1, weight=1)
		self.grid_rowconfigure(0, weight=1)

		# --- Sidebar ---
		self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
		self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
		
		self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="WAGE ENGINE", font=ctk.CTkFont(size=20, weight="bold"))
		self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

		# self.btn_home = ctk.CTkButton(self.sidebar_frame, text="Dashboard", command=self.show_dashboard)
		# self.btn_home.grid(row=1, column=0, padx=20, pady=10)

		self.btn_files = ctk.CTkButton(self.sidebar_frame, text="Open Folder", fg_color="transparent", border_width=1, command=lambda: os.startfile("."))
		self.btn_files.grid(row=2, column=0, padx=20, pady=10)

		self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance:", anchor="w")
		self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(150, 0))
		self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark", "System"], command=self.change_appearance_mode)
		self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))
		self.appearance_mode_optionemenu.set("Dark")

		# --- Main Content Area ---
		self.main_container = ctk.CTkScrollableFrame(self, fg_color="transparent")
		self.main_container.grid(row=0, column=1, padx=20, pady=20, sticky="nsew")
		
		self.show_dashboard()

	def show_dashboard(self):
		# Clear container
		for widget in self.main_container.winfo_children():
			widget.destroy()

		# # Welcome Header
		# self.header = ctk.CTkLabel(self.main_container, text="Wage Options", font=ctk.CTkFont(size=24, weight="bold"))
		# self.header.pack(anchor="w", pady=(0, 20))

		# --- Card 1: Setup ---
		self.setup_card = ctk.CTkFrame(self.main_container)
		self.setup_card.pack(fill="x", pady=10)
		
		ctk.CTkLabel(self.setup_card, text="Configuration Files", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=20, pady=10, sticky="w")

		# files_frame = ctk.CTkFrame(self.setup_card, fg_color="transparent")
		# files_frame.grid(row=1, column=0, padx=10, pady=10)

		# Force columns 0 and 1 to be equal width
		self.setup_card.grid_columnconfigure(0, weight=1)
		self.setup_card.grid_columnconfigure(1, weight=1)

		# MAIN SETUP BUTTONS
		# Bages button
		ctk.CTkButton(self.setup_card, text="Bage Numbers", fg_color="#4f46e5", hover_color="#4338ca", command=lambda: os.startfile(config.BADGE_NUMBER_FILE)).grid(row=1, column=0, padx=5, pady=(0, 15), sticky="ew")

		# Public holiday button
		ctk.CTkButton(self.setup_card, text="Public Holidays", fg_color="#4f46e5", hover_color="#4338ca", command=lambda: os.startfile(config.PUBLIC_HOILIDAY_FILE)).grid(row=1, column=1, padx=5, pady=(0, 15), sticky="ew")
		
		# Init satabase
		# ctk.CTkButton(self.setup_card, text="Initialize Database", fg_color="#4f46e5", hover_color="#4338ca", command=self.init_sys).grid(row=2, column=0, padx=20, pady=(0, 15), sticky="ew")
		
		# Rosters button
		ctk.CTkButton(self.setup_card, text="Rosters", fg_color="#4f46e5", hover_color="#4338ca", command=lambda: os.startfile(config.ROSTER_FOLDER)).grid(row=2, column=0, padx=5, pady=(0, 15), sticky="ew")
		
		# Baker button
		ctk.CTkButton(self.setup_card, text="Baker Cashier", fg_color="#4f46e5", hover_color="#4338ca", command=lambda: os.startfile(config.BAKER_CASHIER_FILE)).grid(row=2, column=1, padx=5, pady=(0, 15), sticky="ew")
		
		# Template (Employee Managemant) button
		def templates():
			top = ctk.CTkToplevel()
			top.attributes("-topmost", True)
			top.geometry("400x620")  # Slightly wider/taller for better spacing
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
					self.db.add_employees(english_name, full_name, surname, id_pass)

					# Clear entry boxes
					ename_entry.delete(0, ctk.END)
					fname_entry.delete(0, ctk.END)
					sname_entry.delete(0, ctk.END)
					id_entry.delete(0, ctk.END)
			
				else:
					CTkMessagebox(title="Add Employee", 
                        message="Operation Canceled",
                        icon="cancel")

			def search():
				english_name = ename_entry.get()
				ename_entry.delete(0, ctk.END)

				search_results = self.db.search_employees(english_name)

				# Loop through results
				for x in search_results:
					# Ask if right employee
					response = CTkMessagebox(title="Employee Search", 
                        message=f'{x} : Is this the right employee?',
                        icon="question", 
                        option_1="No", 
                        option_2="Yes")
					
					# Logic if yes or no
					if response == 'Yes':
						ename_entry.insert(0, x[0])
						fname_entry.insert(0, x[1])
						sname_entry.insert(0, x[2])
						id_entry.insert(0, x[3])
						id_entry.configure(state="readonly") # Correct CTK attribute
						break

			def update():
				pass
			# 	response = messagebox.askyesno('Update Employee', 'Are you sure you want to update the employee?')
			# 	if response == 1:
			# 		english_name = ename_entry.get().capitalize()
			# 		full_name = fname_entry.get().capitalize()
			# 		surname = sname_entry.get().capitalize()
			# 		id_pass = id_entry.get()

			# 		self.payslips.update_employees(english_name, full_name, surname, id_pass)

			# 		ename_entry.delete(0, ctk.END)
			# 		fname_entry.delete(0, ctk.END)
			# 		sname_entry.delete(0, ctk.END)
			# 		id_entry.configure(state="normal")
			# 		id_entry.delete(0, ctk.END)
			# 	else:
			# 		messagebox.showinfo('Update Employee', 'Nothing happened!')

			def delete():
				self.db.employee_management()
				# db_manager.employee_management()
				# response = messagebox.askyesno('Delete Employee', 'Are you sure you want to delete the employee?')
				# if response == 1:
				# 	id_pass = id_entry.get()
				# 	# self.payslips.delete_employees(id_pass)

				# 	ename_entry.delete(0, ctk.END)
				# 	fname_entry.delete(0, ctk.END)
				# 	sname_entry.delete(0, ctk.END)
				# 	id_entry.configure(state="normal")
				# 	id_entry.delete(0, ctk.END)
				# else:
				# 	messagebox.showinfo('Delete Employee', 'Nothing happened!')

			def clear():
				ename_entry.delete(0, ctk.END)
				fname_entry.delete(0, ctk.END)
				sname_entry.delete(0, ctk.END)
				id_entry.configure(state="normal")
				id_entry.delete(0, ctk.END)

			# --- ACTION BUTTONS ---
			# Primary action (Add) uses the Success Green
			ctk.CTkButton(top, text="Add New Employee", fg_color="#10b981", hover_color="#059669", 
							font=label_font, command=save).grid(row=4, column=0, columnspan=2, sticky="ew", padx=20, pady=(20, 5))

			# Search and Update use Indigo
			ctk.CTkButton(top, text="Search by English Name", fg_color="#4f46e5", hover_color="#4338ca", 
							command=search).grid(row=5, column=0, columnspan=2, sticky="ew", padx=20, pady=5)

			ctk.CTkButton(top, text="Update Employee Details", fg_color="#4f46e5", hover_color="#4338ca", 
							command=update).grid(row=6, column=0, columnspan=2, sticky="ew", padx=20, pady=5)

			# Delete uses a warning color
			ctk.CTkButton(top, text="Delete Employee", fg_color="#ef4444", hover_color="#b91c1c", 
							command=delete).grid(row=7, column=0, columnspan=2, sticky="ew", padx=20, pady=5)

			# Bulk add and clear use transparent/outlined styles
			ctk.CTkButton(top, text="Bulk Add (CSV)", fg_color="transparent", border_width=1,).grid(row=8, column=0, columnspan=2, sticky="ew", padx=20, pady=5)

			ctk.CTkButton(top, text="Clear Form", fg_color="transparent", border_width=1, text_color=("#1e293b", "#cbd5e1"),
							command=clear).grid(row=9, column=0, columnspan=2, sticky="ew", padx=20, pady=(5, 20))

		ctk.CTkButton(self.setup_card, text="Template", fg_color="#4f46e5", hover_color="#4338ca", command=templates).grid(row=3, column=0, columnspan=2, padx=5, pady=(0, 15), sticky="ew")

		# file_btns = [
		# 	("Badge Numbers", "Badges.xlsx"), 
		# 	("Public Holidays", "Public Holidays.xlsx")
		# ]
			
		# # Get the directory where main.py is located
		# BASE_DIR = os.path.dirname(os.path.abspath(__file__))
		
		# # Loop through and display buttons
		# for i, (name, path) in enumerate(file_btns):
		# 	if name == 'Rosters':
		# 		full_path = "Rosters"
		# 	else:
		# 		full_path = os.path.join(BASE_DIR, name, path)
			
		# 	ctk.CTkButton(files_frame, text=name, width=100, command=lambda p=full_path: os.startfile(p)).grid(row=0, column=i, padx=(12,0))

		

		
		
		
		
		
		
		
		
		
		
		
		
		
		
		# --- Card 2: Processing ---
		self.ops_card = ctk.CTkFrame(self.main_container)
		self.ops_card.pack(fill="x", pady=10)
		
		ctk.CTkLabel(self.ops_card, text="2. Data Processing", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=10)
		
		ctk.CTkButton(self.ops_card, text="RUN MAIN WAGE PROGRAM", height=40, font=ctk.CTkFont(weight="bold"), fg_color="#10b981", hover_color="#059669", command=self.run_wages).pack(fill="x", padx=20, pady=5)
		
		ctk.CTkButton(self.ops_card, text="Recalculate Hours", command=self.run_recal).pack(fill="x", padx=20, pady=5)
		ctk.CTkButton(self.ops_card, text="Open Hours Sheet", fg_color="transparent", border_width=1, command=lambda: os.startfile("Wage Times.xlsx")).pack(fill="x", padx=20, pady=(5, 15))

		# --- Card 3: Finalization ---
		self.final_card = ctk.CTkFrame(self.main_container)
		self.final_card.pack(fill="x", pady=10)
		
		ctk.CTkLabel(self.final_card, text="3. Payroll & Payslips", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=10)
		
		self.final_grid = ctk.CTkFrame(self.final_card, fg_color="transparent")
		self.final_grid.pack(fill="x", padx=20, pady=(0, 15))
		
		ctk.CTkButton(self.final_grid, text="Calculate Tax", command=self.run_tax).grid(row=0, column=0, padx=(0, 5), sticky="ew")
		ctk.CTkButton(self.final_grid, text="Generate Slips", fg_color="#4f46e5", command=self.run_slips).grid(row=0, column=1, padx=(5, 0), sticky="ew")
		self.final_grid.grid_columnconfigure((0, 1), weight=1)

	# --- Logic Wrappers ---
	def init_sys(self):
		self.db.initialize_tables()
		messagebox.showinfo("Success", "System Database Ready")

	def run_wages(self):
		try:
			# Add your specific workflow here
			self.db.clear_session_data()
			self.processor.collect_clock_times("Att")
			self.processor.calculate_sheet_hours("Att Week One", "Att")
			messagebox.showinfo("Success", "Wage program finished successfully")
		except Exception:
			messagebox.showerror("Error", traceback.format_exc())

	def run_recal(self):
		self.processor.calculate_sheet_hours("Att Week One", "Att")
		messagebox.showinfo("Recal", "Hours updated.")

	def run_tax(self):
		self.payroll.calculate_tax()
		messagebox.showinfo("Tax", "Tax logic finished.")

	def run_slips(self):
		self.payslips.generate_all()
		messagebox.showinfo("Payslips", "Payslips generated in /Payslips.")

	def change_appearance_mode(self, new_mode):
		ctk.set_appearance_mode(new_mode)

if __name__ == "__main__":
	app = WageApp()
	app.mainloop()