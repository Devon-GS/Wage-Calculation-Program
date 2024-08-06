from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.styles.borders import Border, Side
import pandas as pd

def gen_payslips():
    # Read in wxcel workbook
    df = pd.read_excel('Payroll/payroll.xlsx')

    # Get column names
    col_names = df.columns.tolist()

    # Get date of wages
    date = col_names[0]

    # # # Get info of first column headings
    headings = df[df.columns[0]].tolist()
    headings[:0] = ['name']

    # Get names of employees
    names = col_names[2:-1]  # Filter list remove first 2 elements and last element

    employees = []
    for e in names:
        if 'Null' not in e:
            employees.append(e)

    # Get employee info and loop through and create payslips
    for name in employees:
        index = col_names.index(name)
        emp_info = df.get(col_names[index]).tolist()
        emp_info[:0] = [name]

        # # Zip info into dict
        pay_info = dict(zip(headings, emp_info))

        # Create Workbooks with payslips
        wb = Workbook()
        ws = wb.active

        # Title of Payslip
        ws['A1'] = 'WAGE ADVICE:'
        ws['B1'] = 'SASOL DE BRON'

        # Name of employee
        ws['A3'] = 'Name'
        ws['B3'] = pay_info['name']

        # Occupation
        ws['A4'] = 'Occupation'
        ws['B4'] = pay_info['Occupation']

        # Fortnight ending
        ws['A5'] = 'Fornight Ending'
        ws['B5'] = date

        # Hour, Rate, Amount
        ws['B7'] = 'Hours'
        ws['C7'] = 'Rate'
        ws['D7'] = 'Amount'

        # Ordinary Time
        ws['A8'] = 'Ordinary Time'
        ws['B8'] = pay_info['HRS']
        ws['C8'] = pay_info['N_RATE']
        ws['D8'] = '=B8*C8'

        # Overtime
        ws['A9'] = 'Overtime'
        ws['B9'] = pay_info['O/T HRS']
        ws['C9'] = pay_info['OT_RATE']
        ws['D9'] = '=B9*C9'

        # Sunday Times
        ws['A10'] = 'Sunday Times'
        ws['B10'] = pay_info['SUNDAY']
        ws['C10'] = pay_info['S_RATE']
        ws['D10'] = '=B10*C10'

        # Public Holiday
        ws['A11'] = 'Public Holidays'
        ws['B11'] = pay_info['PUB HOL HRS']
        ws['C11'] = pay_info['PH_RATE']
        ws['D11'] = '=B11*C11'

        # Bonus
        ws['A12'] = 'Bonus'
        ws['D12'] = pay_info['BONUS']

        # Sick Leave / Leave
        ws['A13'] = 'Sick Leave / Leave'
        ws['D13'] = pay_info['SICK / LEAVE']

        # Standard Hours
        ws['A14'] = 'Standard Hours'
        ws['D14'] = pay_info['STANDARD HRS']

        # Total Wage
        ws['A15'] = 'Total Wage'
        ws['D15'] = pay_info['TOTAL WAGE']

        # Deduction Heading
        ws['A17'] = 'Deductions'

        # UIF
        ws['A18'] = 'UIF'
        ws['D18'] = pay_info['UIF']

        # Mibco
        ws['A19'] = 'Mibco'
        ws['D19'] = pay_info['MIBCO']

        # Uniforms
        ws['A20'] = 'Uniforms'
        ws['D20'] = pay_info['UNIFORMS']

        # Union
        ws['A21'] = 'Union'
        ws['D21'] = pay_info['UNION']

        # Advance
        ws['A22'] = 'Advance'
        ws['D22'] = pay_info['ADVANCES']

        # Prov Fund
        ws['A23'] = 'Prov Fund'
        ws['D23'] = pay_info['PROV FUND']

        # PAYE
        ws['A24'] = 'PAYE'
        ws['D24'] = pay_info['PAYE']

        # PAYE REPAY
        ws['A25'] = 'PAYE REPAY'
        ws['D25'] = pay_info['PAYE REPAYME']

        # Shortages
        ws['A26'] = 'Shortages'
        ws['D26'] = pay_info['SHORTAGES']

        # Net Wage
        ws['A28'] = 'NET WAGE'
        ws['D28'] = pay_info['NET WAGE']

        # FORMATING
        
        # Border         
        thin = Side(border_style="thin")
        thin_border = Border(top=thin, left=thin, right=thin, bottom=thin)
        
        # Font for headings    
        font_headers = Font(name='Arial', size=10, bold=True)
        
        # Title, name, occupation and week ending
        ws.column_dimensions['A'].width = 17.04
        ws['A1'].font = font_headers
        
        # Merge cells
        ws.merge_cells('B1:D1')
        ws["B1"].alignment = Alignment(horizontal="center")
        ws['B1'].font = font_headers
        
        ws.merge_cells('B3:D3')
        ws["B3"].alignment = Alignment(horizontal="center")
        
        ws.merge_cells('B4:D4')
        ws["B4"].alignment = Alignment(horizontal="center")
        
        ws.merge_cells('B5:D5')
        ws["B5"].alignment = Alignment(horizontal="center")
        
        # Bold Headings
        ws['A15'].font = font_headers
        ws['A17'].font = font_headers
        ws['A28'].font = font_headers

        for row in range(3,29):
            for col in range(1,5):
                ws.cell(row,col).border = thin_border
        
        ws.column_dimensions['D'].width = 9.65
        
        for row in range(8,29):
            for col in range(3,5):
                ws.cell(row,col).number_format = 'R #,##0.00'      
        
        wb.save(f'Payslips/{name}.xlsx')
