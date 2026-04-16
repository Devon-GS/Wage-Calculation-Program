import customtkinter as ctk
from employee_info import pop_up
from tkinter import messagebox
from CTkMessagebox import CTkMessagebox
import config
import os
import traceback

# Import logic files
import database as db
import processor as processor
import payroll_logic as payroll_manager
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
		# self.db = DatabaseManager()
		# self.processor = WageProcessor(self.db)
		# self.payroll = PayrollManager(self.db)
		# self.payslips = PayslipGenerator(self.db)

		# Create Layout
		self.grid_columnconfigure(1, weight=1)
		self.grid_rowconfigure(0, weight=1)

		# --- Sidebar ---
		self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0)
		self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
		
		self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="WAGE ENGINE", font=ctk.CTkFont(size=20, weight="bold"))
		self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

		self.btn_init = ctk.CTkButton(self.sidebar_frame, text="Initialize Database", command=self.init_sys)
		self.btn_init.grid(row=1, column=0, padx=20, pady=10)

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
		ctk.CTkButton(self.setup_card, text="Bage Numbers", fg_color="#4f46e5", hover_color="#4338ca", 
				command=lambda: os.startfile(config.BADGE_NUMBER_FILE)).grid(row=1, column=0, padx=5, pady=(0, 15), sticky="ew")

		# Public holiday button
		ctk.CTkButton(self.setup_card, text="Public Holidays", fg_color="#4f46e5", hover_color="#4338ca", 
				command=self.public_holidays).grid(row=1, column=1, padx=5, pady=(0, 15), sticky="ew")
		
		# Rosters button
		ctk.CTkButton(self.setup_card, text="Rosters", fg_color="#4f46e5", hover_color="#4338ca", 
				command=lambda: os.startfile(config.ROSTER_FOLDER)).grid(row=2, column=0, padx=5, pady=(0, 15), sticky="ew")
		
		# Baker button
		ctk.CTkButton(self.setup_card, text="Baker Cashier", fg_color="#4f46e5", hover_color="#4338ca", 
				command=lambda: os.startfile(config.BAKER_CASHIER_FILE)).grid(row=2, column=1, padx=5, pady=(0, 15), sticky="ew")
		
		# Employee infomation
		ctk.CTkButton(self.setup_card, text="Employee Infomation", fg_color="#4f46e5", hover_color="#4338ca",
				command=pop_up).grid(row=3, column=0, columnspan=2, padx=5, pady=(0, 15), sticky="ew")
		
		# Uniclox button
		ctk.CTkButton(self.setup_card, text="Open Uniclox", fg_color="#4f46e5", hover_color="#4338ca", 
				command=lambda: os.startfile(config.UNICLOX_FOLDER)).grid(row=4, column=0, columnspan=2, padx=5, pady=(0, 15), sticky="ew")

		# --- Card 2: Processing ---
		self.ops_card = ctk.CTkFrame(self.main_container)
		self.ops_card.pack(fill="x", pady=10)
		
		ctk.CTkLabel(self.ops_card, text="Data Processing", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=10)
		
		ctk.CTkButton(self.ops_card, text="RUN WAGE TIME CALCULATION", height=40, font=ctk.CTkFont(weight="bold"), fg_color="#10b981", hover_color="#059669", 
				command=self.run_wages).pack(fill="x", padx=20, pady=5)
		
		ctk.CTkButton(self.ops_card, text="Recalculate Hours", command=self.run_recal).pack(fill="x", padx=20, pady=5)

		ctk.CTkButton(self.ops_card, text="Open Wage Times Sheet", fg_color="transparent", border_width=1, 
				command=lambda: os.startfile(config.WAGE_TIMES_FILE)).pack(fill="x", padx=20, pady=(5, 15))

		# --- Card 3: Finalization ---
		self.final_card = ctk.CTkFrame(self.main_container)
		self.final_card.pack(fill="x", pady=10)
		
		ctk.CTkLabel(self.final_card, text="Payroll & Payslips", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=20, pady=10)
		
		self.final_grid = ctk.CTkFrame(self.final_card, fg_color="transparent")
		self.final_grid.pack(fill="x", padx=20, pady=(0, 15))
		
		ctk.CTkButton(self.final_grid, text="RUN PAYROLL", height=40, font=ctk.CTkFont(weight="bold"), fg_color="#10b981", hover_color="#059669", 
				command=self.run_payroll).grid(row=0, column=0, columnspan=2, padx=(5), pady=(0, 15) ,sticky="ew")
		
		ctk.CTkButton(self.final_grid, text="Open Payroll File", fg_color="transparent", border_width=1,
				command=lambda: os.startfile(config.PAYROLL_FILE)).grid(row=1, column=0,columnspan=2, padx=(5), pady=(0, 15) ,sticky="ew")
		
		ctk.CTkButton(self.final_grid, text="Calculate Tax", command=self.run_tax).grid(row=2, column=0, padx=(0, 5), sticky="ew")
		ctk.CTkButton(self.final_grid, text="Generate Slips", fg_color="#4f46e5", 
				command=self.run_slips).grid(row=2, column=1, padx=(5, 0), sticky="ew")
		
		self.final_grid.grid_columnconfigure((0, 1), weight=1)

	# --- Logic Wrappers ---
	def init_sys(self):
		msg = CTkMessagebox(title="Initialize Database", 
				message="Are you sure you want Initialize the database?",
				icon="question", 
				option_1="No", 
				option_2="Yes")
		
		# Get response
		response = msg.get()

		if response == 'Yes':
			db.initialize_tables()
		else:
			CTkMessagebox(title="Initialize Database", 
				message="Operation Canceled",
				icon="cancel")
			
	def public_holidays(self):
		msg = CTkMessagebox(title="Public Holidays", 
				message="Do you want edit public holidays or update database?",
				icon="question", 
				option_1="Update", 
				option_2="Edit")
		
		# Get response
		response = msg.get()

		if response == 'Edit':
			os.startfile(config.PUBLIC_HOILIDAY_FILE)
		else:
			processor.collect_public_holidays()

	def run_wages(self):
		try:
			# - Clear Excel -
			processor.clear_excel()

			# - Load Workbook -
			wb = processor.load_excel()

			# - Clear database -
			db.clear_session_data()

			# - Send roster shift to db -
			processor.roster_shift_to_db("Attendant", "WeekOne")
			processor.roster_shift_to_db("Attendant", "WeekTwo")
			processor.roster_shift_to_db("Cashier", "WeekOne")
			processor.roster_shift_to_db("Cashier", "WeekTwo")

				# - Collect Clocks -
			processor.collect_clock_times()

			# - Shifts -
			processor.sync_shifts_to_excel(wb, 'Att Week One')
			processor.sync_shifts_to_excel(wb, 'Att Week Two')
			processor.sync_shifts_to_excel(wb, 'Cashier Week One')
			processor.sync_shifts_to_excel(wb, 'Cashier Week Two')


			# - Clock -
			processor.sync_clocks_to_excel(wb, 'Att Week One')
			processor.sync_clocks_to_excel(wb, 'Att Week Two')
			processor.sync_clocks_to_excel(wb, 'Cashier Week One')
			processor.sync_clocks_to_excel(wb, 'Cashier Week Two')

			# - Calculate Hours -
			processor.calculate_hours(wb, 'Att Week One')
			processor.calculate_hours(wb, 'Att Week Two')
			processor.calculate_hours(wb, 'Cashier Week One')
			processor.calculate_hours(wb, 'Cashier Week Two')

			# - Calculate Total Hours -
			processor.cal_total_hours(wb)
			processor.cal_total_hours(wb, "Cashiers")

			#  - Format Excel -
			processor.format_excel(wb)

			# - Save Workbook -
			processor.save_workbook(wb)

			#  - Carwash Times -
			processor.carwash_hours()

			messagebox.showinfo("Success", "Wage program finished successfully")
		except Exception:
			messagebox.showerror("Error", traceback.format_exc())

	def run_recal(self):
		try:
			# - Load Workbook -
			wb = processor.load_excel()
			# - Calculate Hours -
			processor.calculate_hours(wb, 'Att Week One')
			processor.calculate_hours(wb, 'Att Week Two')
			processor.calculate_hours(wb, 'Cashier Week One')
			processor.calculate_hours(wb, 'Cashier Week Two')

			# - Calculate Total Hours -
			processor.cal_total_hours(wb)
			processor.cal_total_hours(wb, "Cashiers")

			# - Save Workbook -
			processor.save_workbook(wb)

			messagebox.showinfo("Success", "Recalculation finished successfully")
		except Exception:
			messagebox.showerror("Error", traceback.format_exc())

	def run_payroll(self):
		payroll_manager.run_payroll()
		messagebox.showinfo("Payroll", "Payroll Run Finished")

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