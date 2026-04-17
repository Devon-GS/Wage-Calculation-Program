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
- [ ] 10. Copy files for saving

# Working and Errors

- [x] Add no leeway on public holidays when calculating hours
- [ ] Add feature: If an employee who works both baker and cashier positions works as a cashier on a Sunday or public holiday, automatically calculate total hours and add to dict under labels `c_pub` and `c_sun`
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
- [ ] Copy for backup/keep
- [ ] Check messageboxs are all CTkMessagebox
- [ ] Change program openning size
- [x] Sort out payroll file errors if more that one file ing folder [make function that is call once in payroll.py file]
- [ ] Add error handling to processor functions
	-[ ] get imployee info function if return empty dic handel it (database.py)
- [ ] Clean up
- [ ] Clean up unsed branchs (git)
- [x] Changes branchs (main to orginal-program) and (refactor to main). Refactor now default branch gtihub

---