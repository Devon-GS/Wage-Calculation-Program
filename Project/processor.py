# Contains the Roster processing and Hour calculations


# import os
# import re
# from datetime import datetime, time
# from openpyxl import load_workbook, Workbook
# from config import WAGE_TIMES_FILE

# class WageProcessor:
#     def __init__(self, db_manager):
#         self.db = db_manager

#     def get_public_holidays(self):
#         holidays = []
#         try:
#             wb = load_workbook("Public Holidays/Public Holidays.xlsx", data_only=True)
#             for row in wb.active.iter_rows(min_row=2, max_col=1, values_only=True):
#                 if row[0]: holidays.append(row[0].strftime('%d/%m/%y'))
#             wb.close()
#         except: pass
#         return holidays

#     def collect_clock_times(self, role):
#         table = "ClockTimeAttendent" if role == "Att" else "ClockTimeCashier"
#         clock_data = []
#         if not os.path.exists('Uniclox'): return

#         files = sorted([f for f in os.listdir('Uniclox/') if 'TL' in f])[-20:]
#         for file in files:
#             with open(f'Uniclox/{file}', 'r') as f:
#                 for line in f:
#                     p = line.strip().split(',')
#                     dt = datetime.strptime(p[1], '%Y-%m-%d %H:%M:%S')
#                     clock_data.append((p[0], dt.strftime("%d/%m/%y"), dt.strftime("%H:%M")))

#         with self.db.get_connection() as con:
#             con.cursor().executemany(f"INSERT INTO {table} (badge, date, time) VALUES (?, ?, ?)", clock_data)

#     def calculate_sheet_hours(self, sheet_name, role):
#         wb = load_workbook(WAGE_TIMES_FILE)
#         ws = wb[sheet_name]
#         holidays = self.get_public_holidays()

#         for i in range(2, ws.max_row + 1):
#             name = ws.cell(row=i, column=1).value
#             if not name or 'Total' in name: continue

#             # Core calculation logic
#             ti, to = ws.cell(row=i, column=5).value, ws.cell(row=i, column=6).value
#             ci, co = ws.cell(row=i, column=7).value, ws.cell(row=i, column=8).value
#             date_val = ws.cell(row=i, column=4).value
#             is_sunday = ws.cell(row=i, column=3).value == "Sunday"

#             if (ti and ti > 0 and not ci) or (to and to > 0 and not co):
#                 ws.cell(row=i, column=12, value="No Clock")
#                 continue

#             # Night Shift Logic (Simplified for modularity)
#             hours = 0
#             if ti == 18:
#                 # Add your specific night shift logic here
#                 hours = 12 # Example
#             elif ti and to:
#                 # Add your normal day logic here
#                 hours = float(to) - float(ti)

#             # Assign to correct column
#             if date_val in holidays: ws.cell(row=i, column=11, value=hours)
#             elif is_sunday: ws.cell(row=i, column=10, value=hours)
#             else: ws.cell(row=i, column=9, value=hours)

#         wb.save(WAGE_TIMES_FILE)


import os
import re
from datetime import datetime, time
from openpyxl import load_workbook
from config import WAGE_TIMES_FILE

class WageProcessor:
    def __init__(self, db_manager):
        self.db = db_manager

    def get_public_holidays(self):
        holidays = []
        try:
            # Note: adjust path if necessary based on your folder structure
            wb = load_workbook("Public Holidays/Public Holidays.xlsx", data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, max_col=1, max_row=20, values_only=True):
                if row[0]:
                    # Standardize format to d/m/y to match roster dates
                    holidays.append(row[0].strftime('%d/%m/%y'))
            wb.close()
        except Exception as e:
            print(f"Error loading holidays: {e}")
        return holidays

    def _adjust_time(self, clock_val, roster_h, is_clock_in):
        """
        Implements the 15-minute rounding logic from your original script.
        clock_val: string "HH:MM"
        roster_h: int (e.g., 8 or 18)
        is_clock_in: Boolean (True for clock in, False for clock out)
        """
        if not clock_val:
            return float(roster_h)

        # Ensure clock_val is HH:MM string
        if not isinstance(clock_val, str):
            clock_val = clock_val.strftime("%H:%M")

        h = int(clock_val.split(":")[0])
        m = int(clock_val.split(":")[1])
        roster_time_str = f"{int(roster_h):02d}:00"

        if is_clock_in:
            # Logic: If clocked in LATER than rostered time, round UP
            if clock_val > roster_time_str:
                if m <= 15: return h + 0.25
                elif m <= 30: return h + 0.50
                elif m <= 45: return h + 0.75
                else: return float(h + 1)
            else:
                return float(roster_h)
        else:
            # Logic: If clocked out EARLIER than rostered time, round DOWN
            if clock_val < roster_time_str:
                if m <= 4: return float(h)
                elif m <= 15: return float(h)
                elif m <= 30: return h + 0.25
                elif m <= 45: return h + 0.50
                else: return h + 0.75
            else:
                return float(roster_h)

    def calculate_sheet_hours(self, sheet_name, role):
        wb = load_workbook(WAGE_TIMES_FILE)
        if sheet_name not in wb.sheetnames:
            return
            
        ws = wb[sheet_name]
        holidays = self.get_public_holidays()

        for i in range(2, ws.max_row + 1):
            name = ws.cell(row=i, column=1).value
            # Skip empty rows or summary rows
            if not name or 'Total' in name:
                continue

            # Load values from Excel
            day = ws.cell(row=i, column=3).value
            date_str = ws.cell(row=i, column=4).value
            ti = ws.cell(row=i, column=5).value  # Roster In (int)
            to = ws.cell(row=i, column=6).value  # Roster Out (int)
            ci = ws.cell(row=i, column=7).value  # Clock In (str/time)
            co = ws.cell(row=i, column=8).value  # Clock Out (str/time)

            # 1. No Clock check
            if (ti and ti > 0 and not ci) or (to and to > 0 and not co):
                ws.cell(row=i, column=12, value="No Clock")
                continue

            # 2. Calculate adjusted hours
            hours_worked = 0.0

            if ti == 18:
                # Night Shift Logic: Only calculate hours from Clock In until Midnight (24:00)
                # The remaining hours (Midnight to 06:00) are usually handled by the next day's row
                tti = self._adjust_time(ci, ti, is_clock_in=True)
                hours_worked = 24.0 - tti
            
            elif ti == 0 and to > 0:
                # Part of Night Shift (Morning finish): 00:00 to clock out
                tto = self._adjust_time(co, to, is_clock_in=False)
                hours_worked = tto # Hours since midnight
                
            elif ti is not None and to is not None:
                # Normal Day Shift logic
                tti = self._adjust_time(ci, ti, is_clock_in=True)
                tto = self._adjust_time(co, to, is_clock_in=False)
                hours_worked = tto - tti

            # 3. Prevent negative hours (just in case)
            hours_worked = max(0, hours_worked)

            # 4. Assign to correct column (9: Normal, 10: Sunday, 11: Public Holiday)
            # Column 11: Public Holiday
            if date_str in holidays:
                ws.cell(row=i, column=11, value=hours_worked)
                ws.cell(row=i, column=9, value=None) # Clear normal column if it was filled
            # Column 10: Sunday
            elif day == "Sunday":
                ws.cell(row=i, column=10, value=hours_worked)
            # Column 9: Normal Weekday
            else:
                ws.cell(row=i, column=9, value=hours_worked)

        wb.save(WAGE_TIMES_FILE)
        wb.close()