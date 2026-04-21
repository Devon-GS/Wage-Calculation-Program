### Build Out

- [x] 1. Get roster shifts
- [x] 2. Get actual clock times
- [x] 3. Put into excel sheet `Wage Times.xlsx`
- [x] 4. Calculate hour worked between shifts and clock time in excel sheet
- [x] 5. Get total normal and sunday hours and add to total sheet
- [x] 6. Send total hour to payroll workbook
- [x] 7. Send carwash times to payroll workbook
- [x] 8. Calculate tax
- [x] 9. Generate payslips
- [x] 10. Copy files for saving

# TODO

- [x] Add no leeway on public holidays when calculating hours
- [x] Add feature: If an employee who works both baker and cashier positions works as a cashier on a Sunday or public holiday, automatically calculate total hours and add to dict under labels `c_pub` and `c_sun`
- [x] Add/change headings in cashier total sheet:
    - [x] Change "Baker's Cashier Hours" -> "B - Cashier Hours"
    - [x] Add "B - Cashier Sunday Hours"
    - [x] Add "B - Cashier Public Holiday Hours"
- [x] Fix: Total wages (normal, Sunday, public) for cashier/baker not calculating
- [x] Format total sheet
- [x] Carwash time to database for export to payroll
- [x] Remove print statements
- [x] Recalculate wages function
- [x] Add run payroll
- [x] Payroll can make sheet any name
- [x] Tax
- [x] Payslips 
- [x] Copy for backup/keep
- [x] Automate Carwash hours
	- [x] File can be any name
	- [x] Add button to open file on side
	- [x] Total and extra align center
	- [x] Update carwash time function for new positions of data
	- [x] Add new functions to payroll run
- [x] Sort out payroll file errors if more that one file ing folder [make function that is call once in payroll.py file]
- [ ] Add error handling to processor functions
	- [ ] get employee info function if return empty dic handel it (database.py)
	- [ ] Log traceback error messages (new folder Logs)
	- [ ] Check messageboxs are all CTkMessagebox
	- [ ] Push errors up to main
- [ ] Change program openning size
- [ ] Clean up
- [ ] Clean up unsed branchs (git)
- [x] Changes branchs (main to orginal-program) and (refactor to main). Refactor now default branch gtihub
- [ ] Tag code V2.0 before deployment

# Errors
- Carwash times file must add formulas to extra time in config
- Must add section to carwash hours for extra time worked
- update functions to extract extra time from carwash hours and write to carwash times

---