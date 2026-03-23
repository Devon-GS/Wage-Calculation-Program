# The GUI that connects everything


# import tkinter as tk
# from tkinter import messagebox, Toplevel
# import traceback
# import os

# from database import DatabaseManager
# from processor import WageProcessor
# from payroll_logic import PayrollManager
# from payslips import PayslipGenerator

# class WageApp:
#     def __init__(self, root):
#         self.root = root
#         self.root.title("Wage System v2.0")
        
#         # Initialize Managers
#         self.db = DatabaseManager()
#         self.processor = WageProcessor(self.db)
#         self.payroll = PayrollManager(self.db)
#         self.payslips = PayslipGenerator(self.db)

#         self.setup_ui()

#     def setup_ui(self):
#         # Setup Buttons
#         tk.Label(self.root, text="WAGE SYSTEM CONTROL", font=("Arial", 12, "bold")).pack(pady=10)
        
#         btn_config = {"width": 25, "pady": 5}
        
#         tk.Button(self.root, text="1. Initialize System", command=self.init_sys, **btn_config).pack()
#         tk.Button(self.root, text="2. Run Wage Hours", command=self.run_wages, **btn_config).pack()
#         tk.Button(self.root, text="3. Calculate Tax", command=self.run_tax, **btn_config).pack()
#         tk.Button(self.root, text="4. Generate Payslips", command=self.run_slips, **btn_config).pack()
#         tk.Button(self.root, text="Open Payroll Folder", command=lambda: os.startfile("Payroll"), **btn_config).pack(pady=20)

#     def init_sys(self):
#         try:
#             self.db.initialize_tables()
#             messagebox.showinfo("Success", "Database and Tables ready.")
#         except Exception as e:
#             messagebox.showerror("Error", str(e))

#     def run_wages(self):
#         try:
#             self.db.clear_session_data()
#             self.processor.collect_clock_times("Att")
#             # Call calculations...
#             messagebox.showinfo("Success", "Wage processing complete.")
#         except Exception:
#             messagebox.showerror("Error", traceback.format_exc())

#     def run_tax(self):
#         self.payroll.calculate_tax()
#         messagebox.showinfo("Success", "Tax calculation finished.")

#     def run_slips(self):
#         self.payslips.generate_all()
#         messagebox.showinfo("Success", "Payslips saved to /Payslips folder.")

# if __name__ == "__main__":
#     root = tk.Tk()
#     root.geometry("300x400")
#     app = WageApp(root)
#     root.mainloop()

import customtkinter as ctk
from tkinter import messagebox
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

        self.btn_home = ctk.CTkButton(self.sidebar_frame, text="Dashboard", command=self.show_dashboard)
        self.btn_home.grid(row=1, column=0, padx=20, pady=10)

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

        # Welcome Header
        self.header = ctk.CTkLabel(self.main_container, text="System Dashboard", font=ctk.CTkFont(size=24, weight="bold"))
        self.header.pack(anchor="w", pady=(0, 20))

        # --- Card 1: Setup ---
        self.setup_card = ctk.CTkFrame(self.main_container)
        self.setup_card.pack(fill="x", pady=10)
        
        ctk.CTkLabel(self.setup_card, text="1. Configuration Files", font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, padx=20, pady=10, sticky="w")
        
        files_frame = ctk.CTkFrame(self.setup_card, fg_color="transparent")
        files_frame.grid(row=1, column=0, padx=10, pady=10)

        file_btns = [
            ("Badges", "Badge Numbers"), ("Rosters", "Rosters"), 
            ("Holidays", "Public Holidays"), ("Baker/Cas", "Baker Cashier")
        ]
        
        for i, (name, path) in enumerate(file_btns):
            ctk.CTkButton(files_frame, text=name, width=100, command=lambda p=path: os.startfile(p)).grid(row=0, column=i, padx=5)

        ctk.CTkButton(self.setup_card, text="Initialize Database", fg_color="#4f46e5", hover_color="#4338ca", command=self.init_sys).grid(row=2, column=0, padx=20, pady=(0, 15), sticky="ew")

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