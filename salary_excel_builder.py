"""
Salary Entry and Complete Accounting Excel Module
Generates comprehensive salary processing Excel workbook with all calculations and journal entries
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
from decimal import Decimal


class SalaryExcelBuilder:
    """
    Generates comprehensive salary processing and accounting Excel workbook
    Includes: Salary calculations, journal entries, ledger heads, and payroll summaries
    """
    
    def __init__(self, filename="Salary_Processing.xlsx", company_name="ABC Trading Company", 
                 financial_month="September 2024"):
        self.filename = filename
        self.company_name = company_name
        self.financial_month = financial_month
        self.workbook = openpyxl.Workbook()
        self.workbook.remove(self.workbook.active)
        self.employees = []
        self.setup_styles()
    
    def setup_styles(self):
        """Setup reusable cell styles"""
        # Headers
        self.main_header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        self.main_header_font = Font(bold=True, color="FFFFFF", size=14)
        
        # Subheaders
        self.subheader_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        self.subheader_font = Font(bold=True, color="FFFFFF", size=11)
        
        # Section headers
        self.section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        self.section_font = Font(bold=True, size=10)
        
        # Totals
        self.total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        self.total_font = Font(bold=True, size=10)
        
        # Debit/Credit highlights
        self.debit_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        self.credit_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        
        # Borders
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Alignments
        self.center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        self.right_alignment = Alignment(horizontal='right', vertical='center')
        self.currency_format = '_("₹"* #,##0.00_);_("₹"* (#,##0.00);_("₹"* "-"??_);_(@_)'
    
    def add_employee(self, employee_id, employee_name, gross_salary):
        """Add employee and calculate all salary components"""
        employee = {
            'id': employee_id,
            'name': employee_name,
            'gross_salary': gross_salary,
            'employee_epf': gross_salary * 0.12,
            'employer_epf': gross_salary * 0.12,
            'employee_esi': gross_salary * 0.0075 if gross_salary <= 21000 else 0,
            'employer_esi': gross_salary * 0.0325 if gross_salary <= 21000 else 0,
            'tds': gross_salary * 0.05
        }
        employee['net_salary'] = (gross_salary - employee['employee_epf'] - 
                                 employee['employee_esi'] - employee['tds'])
        self.employees.append(employee)
        return employee
    
    def create_salary_summary_sheet(self):
        """Create Salary Summary Sheet"""
        ws = self.workbook.create_sheet("Salary Summary")
        
        # Title
        ws.merge_cells('A1:H1')
        title = ws['A1']
        title.value = "SALARY PROCESSING SUMMARY"
        title.font = self.main_header_font
        title.fill = self.main_header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        # Company and Month Info
        ws['A3'] = "Company Name:"
        ws['B3'] = self.company_name
        ws['A4'] = "Financial Month:"
        ws['B4'] = self.financial_month
        ws['A5'] = "Prepared Date:"
        ws['B5'] = datetime.now().strftime("%d-%b-%Y")
        ws['A6'] = "Total Employees:"
        ws['B6'] = len(self.employees)
        
        # Summary Table Header
        headers = ["S.No.", "Employee ID", "Employee Name", "Gross Salary", 
                  "Employee EPF", "Employee ESI", "TDS", "Net Salary"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=8, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.thin_border
        
        # Employee data
        for idx, emp in enumerate(self.employees, 9):
            ws.cell(row=idx, column=1).value = idx - 8
            ws.cell(row=idx, column=2).value = emp['id']
            ws.cell(row=idx, column=3).value = emp['name']
            ws.cell(row=idx, column=4).value = emp['gross_salary']
            ws.cell(row=idx, column=5).value = emp['employee_epf']
            ws.cell(row=idx, column=6).value = emp['employee_esi']
            ws.cell(row=idx, column=7).value = emp['tds']
            ws.cell(row=idx, column=8).value = emp['net_salary']
            
            for col in range(1, 9):
                cell = ws.cell(row=idx, column=col)
                cell.border = self.thin_border
                if col == 1:
                    cell.alignment = self.center_alignment
                elif col in [4, 5, 6, 7, 8]:
                    cell.number_format = self.currency_format
                    cell.alignment = self.right_alignment
                else:
                    cell.alignment = self.left_alignment
        
        # Total row
        total_row = 9 + len(self.employees)
        ws.merge_cells(f'A{total_row}:C{total_row}')
        ws[f'A{total_row}'].value = "TOTAL"
        ws[f'A{total_row}'].font = self.total_font
        ws[f'A{total_row}'].fill = self.total_fill
        ws[f'A{total_row}'].border = self.thin_border
        
        for col in range(4, 9):
            cell = ws.cell(row=total_row, column=col)
            cell.value = f"=SUM({get_column_letter(col)}9:{get_column_letter(col)}{total_row-1})"
            cell.font = self.total_font
            cell.fill = self.total_fill
            cell.border = self.thin_border
            cell.number_format = self.currency_format
            cell.alignment = self.right_alignment
        
        # Adjust column widths
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 15
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 18
        ws.column_dimensions['E'].width = 16
        ws.column_dimensions['F'].width = 16
        ws.column_dimensions['G'].width = 12
        ws.column_dimensions['H'].width = 16
        
        return ws
    
    def create_individual_salary_slips(self):
        """Create Individual Salary Slips for each employee"""
        for emp_num, emp in enumerate(self.employees, 1):
            ws = self.workbook.create_sheet(f"Slip_{emp['id']}")
            
            # Title
            ws.merge_cells('A1:D1')
            title = ws['A1']
            title.value = "SALARY SLIP"
            title.font = self.main_header_font
            title.fill = self.main_header_fill
            title.alignment = self.center_alignment
            ws.row_dimensions[1].height = 25
            
            # Company Info
            ws.merge_cells('A2:D2')
            company = ws['A2']
            company.value = self.company_name
            company.font = Font(bold=True, size=11)
            company.alignment = self.center_alignment
            
            # Employee Details
            ws['A4'] = "Employee ID:"
            ws['B4'] = emp['id']
            ws['A5'] = "Employee Name:"
            ws['B5'] = emp['name']
            ws['A6'] = "Month:"
            ws['B6'] = self.financial_month
            ws['A7'] = "Prepared Date:"
            ws['B7'] = datetime.now().strftime("%d-%b-%Y")
            
            # Earnings Section
            ws['A9'] = "EARNINGS"
            ws['A9'].font = self.section_font
            ws['A9'].fill = self.section_fill
            
            ws['A10'] = "Gross Salary"
            ws['B10'] = emp['gross_salary']
            ws['B10'].number_format = self.currency_format
            
            # Deductions Section
            ws['A12'] = "DEDUCTIONS"
            ws['A12'].font = self.section_font
            ws['A12'].fill = self.section_fill
            
            deductions = [
                ("Employee EPF", emp['employee_epf']),
                ("Employee ESI", emp['employee_esi']),
                ("TDS (Tax Deducted at Source)", emp['tds'])
            ]
            
            for idx, (desc, amount) in enumerate(deductions, 13):
                ws[f'A{idx}'] = desc
                ws[f'B{idx}'] = amount
                ws[f'B{idx}'].number_format = self.currency_format
            
            # Total Deductions
            ws['A16'] = "Total Deductions"
            ws['A16'].font = Font(bold=True)
            ws['B16'] = f"=SUM(B13:B15)"
            ws['B16'].number_format = self.currency_format
            ws['B16'].font = Font(bold=True)
            
            # Net Salary
            ws['A18'] = "NET SALARY PAYABLE"
            ws['A18'].font = self.total_font
            ws['A18'].fill = self.total_fill
            ws['B18'] = emp['net_salary']
            ws['B18'].number_format = self.currency_format
            ws['B18'].font = self.total_font
            ws['B18'].fill = self.total_fill
            
            # Employer Contribution (for reference)
            ws['A20'] = "EMPLOYER CONTRIBUTION (Additional Expense)"
            ws['A20'].font = self.section_font
            ws['A20'].fill = self.section_fill
            
            ws['A21'] = "Employer EPF"
            ws['B21'] = emp['employer_epf']
            ws['B21'].number_format = self.currency_format
            
            ws['A22'] = "Employer ESI"
            ws['B22'] = emp['employer_esi']
            ws['B22'].number_format = self.currency_format
            
            ws['A23'] = "Total Employer Contribution"
            ws['B23'] = f"=B21+B22"
            ws['B23'].number_format = self.currency_format
            ws['B23'].font = Font(bold=True)
            
            # Adjust column widths
            ws.column_dimensions['A'].width = 35
            ws.column_dimensions['B'].width = 18
        
        return len(self.employees)
    
    def create_journal_entries_sheet(self):
        """Create Journal Entries Sheet with all 5 entries"""
        ws = self.workbook.create_sheet("Journal Entries")
        
        # Title
        ws.merge_cells('A1:D1')
        title = ws['A1']
        title.value = "JOURNAL ENTRIES - SALARY PROCESSING"
        title.font = self.main_header_font
        title.fill = self.main_header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        ws['A2'] = f"Company: {self.company_name}"
        ws['A3'] = f"Month: {self.financial_month}"
        
        current_row = 5
        
        # Calculate totals
        total_gross = sum(emp['gross_salary'] for emp in self.employees)
        total_emp_epf = sum(emp['employee_epf'] for emp in self.employees)
        total_emp_esi = sum(emp['employee_esi'] for emp in self.employees)
        total_emp_tds = sum(emp['tds'] for emp in self.employees)
        total_net = sum(emp['net_salary'] for emp in self.employees)
        total_employer_epf = sum(emp['employer_epf'] for emp in self.employees)
        total_employer_esi = sum(emp['employer_esi'] for emp in self.employees)
        
        # Entry 1: Salary Provision Entry
        ws[f'A{current_row}'] = "ENTRY 1: SALARY PROVISION ENTRY"
        ws[f'A{current_row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'A{current_row}'].fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
        ws.merge_cells(f'A{current_row}:D{current_row}')
        current_row += 1
        
        # Headers
        ws[f'A{current_row}'] = "Account"
        ws[f'B{current_row}'] = "Debit (₹)"
        ws[f'C{current_row}'] = "Credit (₹)"
        ws[f'D{current_row}'] = "Description"
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{current_row}'].font = self.subheader_font
            ws[f'{col}{current_row}'].fill = self.subheader_fill
            ws[f'{col}{current_row}'].border = self.thin_border
        current_row += 1
        
        # Entry 1 Lines
        entry1_lines = [
            ("Salary Expense A/c", total_gross, None, "Employee salaries"),
            ("Employer EPF Contribution A/c", total_employer_epf, None, "Employer EPF (12%)"),
            ("Employer ESI Contribution A/c", total_employer_esi, None, "Employer ESI (3.25%)"),
            ("    To Employee EPF Payable A/c", None, total_emp_epf, "Employee EPF payable"),
            ("    To Employee ESI Payable A/c", None, total_emp_esi, "Employee ESI payable"),
            ("    To TDS Payable A/c", None, total_emp_tds, "TDS payable"),
            ("    To Salary Payable A/c", None, total_net, "Net salary payable"),
        ]
        
        for account, debit, credit, desc in entry1_lines:
            ws[f'A{current_row}'] = account
            ws[f'B{current_row}'] = debit if debit else ""
            ws[f'C{current_row}'] = credit if credit else ""
            ws[f'D{current_row}'] = desc
            
            if debit:
                ws[f'B{current_row}'].number_format = self.currency_format
                ws[f'B{current_row}'].fill = self.debit_fill
            if credit:
                ws[f'C{current_row}'].number_format = self.currency_format
                ws[f'C{current_row}'].fill = self.credit_fill
            
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{current_row}'].border = self.thin_border
            
            current_row += 1
        
        # Total line for Entry 1
        ws[f'A{current_row}'] = "TOTAL"
        ws[f'A{current_row}'].font = Font(bold=True)
        ws[f'B{current_row}'] = f"=SUM(B7:B9)"
        ws[f'C{current_row}'] = f"=SUM(C10:C13)"
        ws[f'B{current_row}'].font = Font(bold=True)
        ws[f'C{current_row}'].font = Font(bold=True)
        for col in ['A', 'B', 'C']:
            ws[f'{col}{current_row}'].fill = self.total_fill
            ws[f'{col}{current_row}'].border = self.thin_border
            ws[f'{col}{current_row}'].number_format = self.currency_format
        
        current_row += 3
        
        # Entry 2: Net Salary Payment
        ws[f'A{current_row}'] = "ENTRY 2: NET SALARY PAYMENT TO EMPLOYEES"
        ws[f'A{current_row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'A{current_row}'].fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
        ws.merge_cells(f'A{current_row}:D{current_row}')
        current_row += 1
        
        # Headers
        for col_letter, col_text in [('A', 'Account'), ('B', 'Debit (₹)'), ('C', 'Credit (₹)'), ('D', 'Description')]:
            ws[f'{col_letter}{current_row}'].value = col_text
            ws[f'{col_letter}{current_row}'].font = self.subheader_font
            ws[f'{col_letter}{current_row}'].fill = self.subheader_fill
            ws[f'{col_letter}{current_row}'].border = self.thin_border
        current_row += 1
        
        # Entry 2 Lines
        ws[f'A{current_row}'] = "Salary Payable A/c"
        ws[f'B{current_row}'] = total_net
        ws[f'B{current_row}'].number_format = self.currency_format
        ws[f'B{current_row}'].fill = self.debit_fill
        ws[f'D{current_row}'] = "Salary paid to employees"
        current_row += 1
        
        ws[f'A{current_row}'] = "    To Bank A/c"
        ws[f'C{current_row}'] = total_net
        ws[f'C{current_row}'].number_format = self.currency_format
        ws[f'C{current_row}'].fill = self.credit_fill
        current_row += 1
        
        # Total
        ws[f'A{current_row}'] = "TOTAL"
        ws[f'A{current_row}'].font = Font(bold=True)
        ws[f'B{current_row}'] = total_net
        ws[f'C{current_row}'] = total_net
        for col in ['A', 'B', 'C']:
            ws[f'{col}{current_row}'].fill = self.total_fill
            ws[f'{col}{current_row}'].border = self.thin_border
            ws[f'{col}{current_row}'].font = Font(bold=True)
            ws[f'{col}{current_row}'].number_format = self.currency_format
        
        current_row += 3
        
        # Entry 3: EPF Deposit
        ws[f'A{current_row}'] = "ENTRY 3: EPF DEPOSIT"
        ws[f'A{current_row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'A{current_row}'].fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
        ws.merge_cells(f'A{current_row}:D{current_row}')
        current_row += 1
        
        # Headers
        for col_letter, col_text in [('A', 'Account'), ('B', 'Debit (₹)'), ('C', 'Credit (₹)'), ('D', 'Description')]:
            ws[f'{col_letter}{current_row}'].value = col_text
            ws[f'{col_letter}{current_row}'].font = self.subheader_font
            ws[f'{col_letter}{current_row}'].fill = self.subheader_fill
            ws[f'{col_letter}{current_row}'].border = self.thin_border
        current_row += 1
        
        # Entry 3 Lines
        ws[f'A{current_row}'] = "Employee EPF Payable A/c"
        ws[f'B{current_row}'] = total_emp_epf
        ws[f'B{current_row}'].number_format = self.currency_format
        ws[f'B{current_row}'].fill = self.debit_fill
        ws[f'D{current_row}'] = "Employee contribution"
        current_row += 1
        
        ws[f'A{current_row}'] = "Employer EPF Payable A/c"
        ws[f'B{current_row}'] = total_employer_epf
        ws[f'B{current_row}'].number_format = self.currency_format
        ws[f'B{current_row}'].fill = self.debit_fill
        ws[f'D{current_row}'] = "Employer contribution"
        current_row += 1
        
        ws[f'A{current_row}'] = "    To Bank A/c"
        ws[f'C{current_row}'] = total_emp_epf + total_employer_epf
        ws[f'C{current_row}'].number_format = self.currency_format
        ws[f'C{current_row}'].fill = self.credit_fill
        ws[f'D{current_row}'] = "Total EPF deposited"
        current_row += 1
        
        # Total
        ws[f'A{current_row}'] = "TOTAL"
        ws[f'A{current_row}'].font = Font(bold=True)
        ws[f'B{current_row}'] = f"=SUM(B{current_row-2}:B{current_row-1})"
        ws[f'C{current_row}'] = total_emp_epf + total_employer_epf
        for col in ['A', 'B', 'C']:
            ws[f'{col}{current_row}'].fill = self.total_fill
            ws[f'{col}{current_row}'].border = self.thin_border
            ws[f'{col}{current_row}'].font = Font(bold=True)
            ws[f'{col}{current_row}'].number_format = self.currency_format
        
        current_row += 3
        
        # Entry 4: ESI Deposit
        ws[f'A{current_row}'] = "ENTRY 4: ESI DEPOSIT"
        ws[f'A{current_row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'A{current_row}'].fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
        ws.merge_cells(f'A{current_row}:D{current_row}')
        current_row += 1
        
        # Headers
        for col_letter, col_text in [('A', 'Account'), ('B', 'Debit (₹)'), ('C', 'Credit (₹)'), ('D', 'Description')]:
            ws[f'{col_letter}{current_row}'].value = col_text
            ws[f'{col_letter}{current_row}'].font = self.subheader_font
            ws[f'{col_letter}{current_row}'].fill = self.subheader_fill
            ws[f'{col_letter}{current_row}'].border = self.thin_border
        current_row += 1
        
        # Entry 4 Lines
        ws[f'A{current_row}'] = "Employee ESI Payable A/c"
        ws[f'B{current_row}'] = total_emp_esi
        ws[f'B{current_row}'].number_format = self.currency_format
        ws[f'B{current_row}'].fill = self.debit_fill
        ws[f'D{current_row}'] = "Employee contribution"
        current_row += 1
        
        ws[f'A{current_row}'] = "Employer ESI Payable A/c"
        ws[f'B{current_row}'] = total_employer_esi
        ws[f'B{current_row}'].number_format = self.currency_format
        ws[f'B{current_row}'].fill = self.debit_fill
        ws[f'D{current_row}'] = "Employer contribution"
        current_row += 1
        
        ws[f'A{current_row}'] = "    To Bank A/c"
        ws[f'C{current_row}'] = total_emp_esi + total_employer_esi
        ws[f'C{current_row}'].number_format = self.currency_format
        ws[f'C{current_row}'].fill = self.credit_fill
        ws[f'D{current_row}'] = "Total ESI deposited"
        current_row += 1
        
        # Total
        ws[f'A{current_row}'] = "TOTAL"
        ws[f'A{current_row}'].font = Font(bold=True)
        ws[f'B{current_row}'] = f"=SUM(B{current_row-2}:B{current_row-1})"
        ws[f'C{current_row}'] = total_emp_esi + total_employer_esi
        for col in ['A', 'B', 'C']:
            ws[f'{col}{current_row}'].fill = self.total_fill
            ws[f'{col}{current_row}'].border = self.thin_border
            ws[f'{col}{current_row}'].font = Font(bold=True)
            ws[f'{col}{current_row}'].number_format = self.currency_format
        
        current_row += 3
        
        # Entry 5: TDS Deposit
        ws[f'A{current_row}'] = "ENTRY 5: TDS DEPOSIT TO GOVERNMENT"
        ws[f'A{current_row}'].font = Font(bold=True, size=12, color="FFFFFF")
        ws[f'A{current_row}'].fill = PatternFill(start_color="203864", end_color="203864", fill_type="solid")
        ws.merge_cells(f'A{current_row}:D{current_row}')
        current_row += 1
        
        # Headers
        for col_letter, col_text in [('A', 'Account'), ('B', 'Debit (₹)'), ('C', 'Credit (₹)'), ('D', 'Description')]:
            ws[f'{col_letter}{current_row}'].value = col_text
            ws[f'{col_letter}{current_row}'].font = self.subheader_font
            ws[f'{col_letter}{current_row}'].fill = self.subheader_fill
            ws[f'{col_letter}{current_row}'].border = self.thin_border
        current_row += 1
        
        # Entry 5 Lines
        ws[f'A{current_row}'] = "TDS Payable A/c"
        ws[f'B{current_row}'] = total_emp_tds
        ws[f'B{current_row}'].number_format = self.currency_format
        ws[f'B{current_row}'].fill = self.debit_fill
        ws[f'D{current_row}'] = "TDS deducted from salary"
        current_row += 1
        
        ws[f'A{current_row}'] = "    To Bank A/c"
        ws[f'C{current_row}'] = total_emp_tds
        ws[f'C{current_row}'].number_format = self.currency_format
        ws[f'C{current_row}'].fill = self.credit_fill
        ws[f'D{current_row}'] = "TDS deposited to government"
        current_row += 1
        
        # Total
        ws[f'A{current_row}'] = "TOTAL"
        ws[f'A{current_row}'].font = Font(bold=True)
        ws[f'B{current_row}'] = total_emp_tds
        ws[f'C{current_row}'] = total_emp_tds
        for col in ['A', 'B', 'C']:
            ws[f'{col}{current_row}'].fill = self.total_fill
            ws[f'{col}{current_row}'].border = self.thin_border
            ws[f'{col}{current_row}'].font = Font(bold=True)
            ws[f'{col}{current_row}'].number_format = self.currency_format
        
        # Set column widths
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 18
        ws.column_dimensions['C'].width = 18
        ws.column_dimensions['D'].width = 30
        
        return ws
    
    def create_ledger_heads_sheet(self):
        """Create Ledger Heads Classification Sheet"""
        ws = self.workbook.create_sheet("Ledger Heads")
        
        # Title
        ws.merge_cells('A1:C1')
        title = ws['A1']
        title.value = "LEDGER HEADS CLASSIFICATION"
        title.font = self.main_header_font
        title.fill = self.main_header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        current_row = 3
        
        ledger_data = {
            'ASSETS': [
                'Bank A/c'
            ],
            'EXPENSES': [
                'Salary Expense A/c',
                'Employer EPF Contribution A/c',
                'Employer ESI Contribution A/c'
            ],
            'LIABILITIES': [
                'Salary Payable A/c',
                'Employee EPF Payable A/c',
                'Employee ESI Payable A/c',
                'Employer EPF Payable A/c',
                'Employer ESI Payable A/c',
                'TDS Payable A/c'
            ],
            'INDIRECT EXPENSES': [
                'Employee Benefits Expense'
            ]
        }
        
        for category, accounts in ledger_data.items():
            # Category Header
            ws.merge_cells(f'A{current_row}:C{current_row}')
            header = ws[f'A{current_row}']
            header.value = category
            header.font = self.section_font
            header.fill = self.section_fill
            header.border = self.thin_border
            current_row += 1
            
            # Column headers
            for col_letter, col_text in [('A', 'S.No.'), ('B', 'Ledger Head'), ('C', 'Category')]:
                ws[f'{col_letter}{current_row}'].value = col_text
                ws[f'{col_letter}{current_row}'].font = self.subheader_font
                ws[f'{col_letter}{current_row}'].fill = self.subheader_fill
                ws[f'{col_letter}{current_row}'].border = self.thin_border
            current_row += 1
            
            # Accounts
            for idx, account in enumerate(accounts, 1):
                ws[f'A{current_row}'] = idx
                ws[f'B{current_row}'] = account
                ws[f'C{current_row}'] = category
                
                for col_letter in ['A', 'B', 'C']:
                    ws[f'{col_letter}{current_row}'].border = self.thin_border
                    ws[f'{col_letter}{current_row}'].alignment = self.left_alignment
                
                current_row += 1
            
            current_row += 1
        
        # Set column widths
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 25
        
        return ws
    
    def create_payroll_summary_sheet(self):
        """Create Detailed Payroll Summary Sheet"""
        ws = self.workbook.create_sheet("Payroll Summary")
        
        # Title
        ws.merge_cells('A1:D1')
        title = ws['A1']
        title.value = "DETAILED PAYROLL SUMMARY"
        title.font = self.main_header_font
        title.fill = self.main_header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        # Summary Info
        ws['A3'] = "Company:"
        ws['B3'] = self.company_name
        ws['A4'] = "Month:"
        ws['B4'] = self.financial_month
        ws['A5'] = "Total Employees:"
        ws['B5'] = len(self.employees)
        
        # Calculate totals
        total_gross = sum(emp['gross_salary'] for emp in self.employees)
        total_emp_epf = sum(emp['employee_epf'] for emp in self.employees)
        total_emp_esi = sum(emp['employee_esi'] for emp in self.employees)
        total_tds = sum(emp['tds'] for emp in self.employees)
        total_net = sum(emp['net_salary'] for emp in self.employees)
        total_employer_epf = sum(emp['employer_epf'] for emp in self.employees)
        total_employer_esi = sum(emp['employer_esi'] for emp in self.employees)
        total_deductions = total_emp_epf + total_emp_esi + total_tds
        
        # Summary Table
        current_row = 7
        
        ws.merge_cells(f'A{current_row}:B{current_row}')
        ws[f'A{current_row}'].value = "SUMMARY DETAILS"
        ws[f'A{current_row}'].font = self.section_font
        ws[f'A{current_row}'].fill = self.section_fill
        current_row += 1
        
        summary_items = [
            ("Total Gross Salary", total_gross),
            ("", ""),
            ("Employee Deductions:", ""),
            ("  Employee EPF (12%)", total_emp_epf),
            ("  Employee ESI (0.75%)", total_emp_esi),
            ("  TDS (5%)", total_tds),
            ("Total Employee Deductions", total_deductions),
            ("", ""),
            ("Net Salary Payable to Employees", total_net),
            ("", ""),
            ("Employer Contributions (Additional Expense):", ""),
            ("  Employer EPF (12%)", total_employer_epf),
            ("  Employer ESI (3.25%)", total_employer_esi),
            ("Total Employer Contribution", total_employer_epf + total_employer_esi),
            ("", ""),
            ("Total Salary Cost (Gross + Employer Contribution)", total_gross + total_employer_epf + total_employer_esi),
        ]
        
        for desc, amount in summary_items:
            ws[f'A{current_row}'] = desc
            if amount != "":
                ws[f'B{current_row}'] = amount
                ws[f'B{current_row}'].number_format = self.currency_format
                ws[f'B{current_row}'].alignment = self.right_alignment
            
            if desc in ["Total Gross Salary", "Total Employee Deductions", "Net Salary Payable to Employees", 
                       "Total Employer Contribution", "Total Salary Cost (Gross + Employer Contribution)"]:
                ws[f'A{current_row}'].font = self.total_font
                ws[f'A{current_row}'].fill = self.total_fill
                if amount != "":
                    ws[f'B{current_row}'].font = self.total_font
                    ws[f'B{current_row}'].fill = self.total_fill
            
            for col_letter in ['A', 'B']:
                ws[f'{col_letter}{current_row}'].border = self.thin_border
            
            current_row += 1
        
        # Set column widths
        ws.column_dimensions['A'].width = 50
        ws.column_dimensions['B'].width = 18
        
        return ws
    
    def save(self):
        """Save the Excel workbook"""
        self.workbook.save(self.filename)
        print(f"\n✓ Salary Processing Excel created successfully: {self.filename}")
        return self.filename


def main():
    """Main function to generate Salary Processing Excel"""
    
    # Initialize builder
    builder = SalaryExcelBuilder(
        "Salary_Processing_Sep2024.xlsx",
        "ABC Trading Company",
        "September 2024"
    )
    
    # Add employees
    print("Processing Employees...\n")
    builder.add_employee("EMP001", "Rajesh Kumar", 30000)
    builder.add_employee("EMP002", "Priya Singh", 25000)
    builder.add_employee("EMP003", "Amit Patel", 18000)
    builder.add_employee("EMP004", "Neha Gupta", 22000)
    builder.add_employee("EMP005", "Vikram Sharma", 28000)
    
    # Create all sheets
    print("Creating Excel Sheets...")
    builder.create_salary_summary_sheet()
    builder.create_individual_salary_slips()
    builder.create_journal_entries_sheet()
    builder.create_ledger_heads_sheet()
    builder.create_payroll_summary_sheet()
    
    # Save workbook
    builder.save()
    
    print("\nSheets created:")
    print("1. ✓ Salary Summary - Overview of all employees")
    print("2. ✓ Individual Salary Slips - Separate slip for each employee")
    print("3. ✓ Journal Entries - All 5 accounting entries")
    print("4. ✓ Ledger Heads - Classification of accounts")
    print("5. ✓ Payroll Summary - Detailed payroll analysis")


if __name__ == "__main__":
    main()
