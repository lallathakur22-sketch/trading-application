import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import json

class ITRCompositionExcelBuilder:
    """
    Automatic ITR (Income Tax Return) Composition Excel Builder
    Generates comprehensive ITR composition sheets for trading applications
    """
    
    def __init__(self, filename="ITR_Composition.xlsx"):
        self.filename = filename
        self.workbook = openpyxl.Workbook()
        self.workbook.remove(self.workbook.active)  # Remove default sheet
        self.setup_styles()
    
    def setup_styles(self):
        """Setup reusable cell styles"""
        self.header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        self.header_font = Font(bold=True, color="FFFFFF", size=12)
        self.subheader_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
        self.subheader_font = Font(bold=True, color="FFFFFF", size=11)
        self.section_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
        self.section_font = Font(bold=True, size=10)
        self.total_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        self.total_font = Font(bold=True, size=10)
        self.border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        self.center_alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        self.left_alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        self.right_alignment = Alignment(horizontal='right', vertical='center')
    
    def create_summary_sheet(self):
        """Create ITR Summary Sheet"""
        ws = self.workbook.create_sheet("ITR Summary")
        
        # Title
        ws.merge_cells('A1:D1')
        title = ws['A1']
        title.value = "ITR (INCOME TAX RETURN) COMPOSITION SUMMARY"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        # Assessment Year
        ws['A3'] = "Assessment Year:"
        ws['B3'] = "FY 2024-25"
        ws['A4'] = "Financial Year:"
        ws['B4'] = "01-Apr-2023 to 31-Mar-2024"
        ws['A5'] = "PAN:"
        ws['B5'] = "[PAN NUMBER]"
        ws['A6'] = "Prepared Date:"
        ws['B6'] = datetime.now().strftime("%d-%b-%Y")
        
        # Summary Table Header
        row = 8
        headers = ["S.No.", "Description", "Amount (₹)", "Remarks"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=row, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.border
        
        # Summary data
        summary_data = [
            [1, "Gross Income", "=SUM('Income Details'!D:D)", "Total of all income sources"],
            [2, "Income from Business/Profession", "=SUM('Business Income'!D:D)", "Trading business income"],
            [3, "Income from Investments", "=SUM('Investment Income'!D:D)", "Dividend, Interest, Capital Gains"],
            [4, "Income from Other Sources", "=SUM('Other Income'!D:D)", "Miscellaneous income"],
            [5, "Total Income Before Deduction", "=B9+B10+B11+B12", ""],
            [6, "Deductions under Section 80", "=SUM('Deductions'!D:D)", "Section 80C, 80D, etc."],
            [7, "Taxable Income", "=B14-B15", "After deductions"],
            [8, "Tax Calculation", "=SUM('Tax Calculation'!D:D)", "Income tax + Cess"],
            [9, "Total Tax Payable", "=B17", ""],
        ]
        
        for idx, data in enumerate(summary_data, 9):
            row = 8 + idx
            cells = [ws.cell(row=row, column=col) for col in range(1, 5)]
            cells[0].value = data[0]
            cells[1].value = data[1]
            cells[2].value = data[2]
            cells[3].value = data[3]
            
            for cell in cells:
                cell.border = self.border
                cell.alignment = self.left_alignment if cell.column != 3 else self.right_alignment
                if idx > 5:
                    cell.fill = self.total_fill
                    cell.font = self.total_font
        
        ws.column_dimensions['A'].width = 5
        ws.column_dimensions['B'].width = 30
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 25
        
        return ws
    
    def create_income_details_sheet(self):
        """Create Income Details Sheet"""
        ws = self.workbook.create_sheet("Income Details")
        
        # Header
        ws.merge_cells('A1:E1')
        title = ws['A1']
        title.value = "INCOME DETAILS - ALL SOURCES"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        # Column Headers
        headers = ["S.No.", "Source of Income", "Category", "Amount (₹)", "Remarks"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.border
        
        # Sample data structure
        income_categories = [
            ["1", "Salary", "Salary & Wages", 0, ""],
            ["2", "Business Income", "Business/Profession", 0, ""],
            ["3", "Dividend Income", "Investment Income", 0, ""],
            ["4", "Interest Income", "Investment Income", 0, ""],
            ["5", "Capital Gains (Short-term)", "Capital Gains", 0, ""],
            ["6", "Capital Gains (Long-term)", "Capital Gains", 0, ""],
            ["7", "Rental Income", "Property", 0, ""],
            ["8", "Other Income", "Miscellaneous", 0, ""],
        ]
        
        for idx, row_data in enumerate(income_categories, 4):
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                cell.alignment = self.center_alignment if col == 1 else self.left_alignment if col != 4 else self.right_alignment
        
        # Total row
        total_row = 4 + len(income_categories)
        ws.merge_cells(f'A{total_row}:C{total_row}')
        total_cell = ws[f'A{total_row}']
        total_cell.value = "TOTAL INCOME"
        total_cell.font = self.total_font
        total_cell.fill = self.total_fill
        total_cell.border = self.border
        
        sum_cell = ws.cell(row=total_row, column=4)
        sum_cell.value = f"=SUM(D4:D{total_row-1})"
        sum_cell.font = self.total_font
        sum_cell.fill = self.total_fill
        sum_cell.border = self.border
        
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 18
        
        return ws
    
    def create_business_income_sheet(self):
        """Create Business Income Details Sheet"""
        ws = self.workbook.create_sheet("Business Income")
        
        # Header
        ws.merge_cells('A1:F1')
        title = ws['A1']
        title.value = "BUSINESS/PROFESSIONAL INCOME DETAILS"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        # Business Details
        ws['A3'] = "Nature of Business:"
        ws['B3'] = "[Description]"
        ws['A4'] = "Business Registration:"
        ws['B4'] = "[Registration No.]"
        ws['A5'] = "Financial Year:"
        ws['B5'] = "01-Apr-2023 to 31-Mar-2024"
        
        # Income & Expense Details
        headers = ["S.No.", "Particulars", "Amount (₹)", "Description", "Supporting Docs", "Remarks"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=7, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.border
        
        business_items = [
            ["1", "Total Revenue/Sales", 0, "Gross sales from business"],
            ["2", "Cost of Goods Sold", 0, "Direct material cost"],
            ["3", "Gross Profit", "=C8-C9", "Revenue - COGS"],
            ["4", "Employee Salaries", 0, "Staff cost"],
            ["5", "Rent & Utilities", 0, "Office rent, electricity, water"],
            ["6", "Professional Fees", 0, "Consultant, CA, legal fees"],
            ["7", "Depreciation", 0, "Machinery, equipment depreciation"],
            ["8", "Other Expenses", 0, "Miscellaneous expenses"],
            ["9", "Total Expenses", "=SUM(C11:C17)", "Sum of all expenses"],
            ["10", "Business Net Income", "=C10-C18", "Gross Profit - Total Expenses"],
        ]
        
        for idx, row_data in enumerate(business_items, 8):
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                if col == 1:
                    cell.alignment = self.center_alignment
                elif col == 3:
                    cell.alignment = self.right_alignment
                else:
                    cell.alignment = self.left_alignment
        
        for col in range(1, 7):
            ws.column_dimensions[get_column_letter(col)].width = 18
        
        return ws
    
    def create_investment_income_sheet(self):
        """Create Investment Income Sheet"""
        ws = self.workbook.create_sheet("Investment Income")
        
        # Header
        ws.merge_cells('A1:E1')
        title = ws['A1']
        title.value = "INVESTMENT INCOME DETAILS"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        headers = ["S.No.", "Investment Type", "Amount (₹)", "Tax Status", "Remarks"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.border
        
        investment_types = [
            ["1", "Dividend Income", 0, "Taxable", ""],
            ["2", "Interest from Bank Deposits", 0, "Taxable", ""],
            ["3", "Interest from Post Office", 0, "Taxable", ""],
            ["4", "Mutual Fund Dividends", 0, "Taxable", ""],
            ["5", "Bond Interest", 0, "Taxable", ""],
            ["6", "Short-term Capital Gains", 0, "Taxable", ""],
            ["7", "Long-term Capital Gains", 0, "Concessional", ""],
            ["8", "Capital Gains Exemption", 0, "Exempt", "u/s 54, 54F, etc."],
        ]
        
        for idx, row_data in enumerate(investment_types, 4):
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                if col == 1:
                    cell.alignment = self.center_alignment
                elif col == 3:
                    cell.alignment = self.right_alignment
                else:
                    cell.alignment = self.left_alignment
        
        # Total
        total_row = 4 + len(investment_types)
        ws.merge_cells(f'A{total_row}:B{total_row}')
        ws[f'A{total_row}'].value = "TOTAL INVESTMENT INCOME"
        ws[f'A{total_row}'].font = self.total_font
        ws[f'A{total_row}'].fill = self.total_fill
        ws[f'A{total_row}'].border = self.border
        
        sum_cell = ws.cell(row=total_row, column=3)
        sum_cell.value = f"=SUM(C4:C{total_row-1})"
        sum_cell.font = self.total_font
        sum_cell.fill = self.total_fill
        sum_cell.border = self.border
        
        for col in range(1, 6):
            ws.column_dimensions[get_column_letter(col)].width = 18
        
        return ws
    
    def create_deductions_sheet(self):
        """Create Deductions under Section 80 Sheet"""
        ws = self.workbook.create_sheet("Deductions")
        
        # Header
        ws.merge_cells('A1:D1')
        title = ws['A1']
        title.value = "DEDUCTIONS UNDER SECTION 80"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        headers = ["Section", "Deduction Description", "Amount (₹)", "Limit (₹)"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.border
        
        deductions = [
            ["80C", "Life Insurance Premium", 0, "1,50,000"],
            ["80C", "Provident Fund Contribution", 0, "1,50,000"],
            ["80C", "Public Provident Fund (PPF)", 0, "1,50,000"],
            ["80C", "ELSS Mutual Fund Investment", 0, "1,50,000"],
            ["80C", "NSC (National Savings Certificate)", 0, "1,50,000"],
            ["80D", "Health Insurance Premium", 0, "25,000"],
            ["80D", "Health Insurance (Senior Citizen)", 0, "50,000"],
            ["80E", "Education Loan Interest", 0, "No Limit"],
            ["80G", "Donations to Charity", 0, "No Limit"],
            ["80TTA", "Savings Account Interest", 0, "10,000"],
            ["80U", "Disability/Medical Treatment", 0, "1,00,000"],
        ]
        
        for idx, row_data in enumerate(deductions, 4):
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                if col in [1, 3, 4]:
                    cell.alignment = self.center_alignment
                else:
                    cell.alignment = self.left_alignment
        
        # Total
        total_row = 4 + len(deductions)
        ws.merge_cells(f'A{total_row}:B{total_row}')
        ws[f'A{total_row}'].value = "TOTAL DEDUCTIONS"
        ws[f'A{total_row}'].font = self.total_font
        ws[f'A{total_row}'].fill = self.total_fill
        ws[f'A{total_row}'].border = self.border
        
        sum_cell = ws.cell(row=total_row, column=3)
        sum_cell.value = f"=SUM(C4:C{total_row-1})"
        sum_cell.font = self.total_font
        sum_cell.fill = self.total_fill
        sum_cell.border = self.border
        
        for col in range(1, 5):
            ws.column_dimensions[get_column_letter(col)].width = 20
        
        return ws
    
    def create_tax_calculation_sheet(self):
        """Create Tax Calculation Sheet"""
        ws = self.workbook.create_sheet("Tax Calculation")
        
        # Header
        ws.merge_cells('A1:C1')
        title = ws['A1']
        title.value = "INCOME TAX CALCULATION FY 2024-25"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        # Calculation Details
        calc_rows = [
            ["A", "Gross Income (as per Income Details)", "=SUM('Income Details'!D:D)"],
            ["B", "Deductions (as per Deductions sheet)", "=SUM('Deductions'!C:C)"],
            ["C", "Taxable Income (A - B)", "=A4-A5"],
            ["", "", ""],
            ["D", "Tax Rate (as per slab)", ""],
            ["E", "Tax on Taxable Income", ""],
            ["F", "Rebate u/s 87A (if applicable)", ""],
            ["G", "Tax after Rebate (E - F)", ""],
            ["H", "Health and Education Cess (4%)", "=A10*0.04"],
            ["I", "TOTAL TAX LIABILITY (G + H)", ""],
        ]
        
        for idx, row_data in enumerate(calc_rows, 3):
            if row_data[0] == "":
                continue
            ws.cell(row=idx, column=1).value = row_data[0]
            ws.cell(row=idx, column=2).value = row_data[1]
            ws.cell(row=idx, column=3).value = row_data[2]
            
            for col in range(1, 4):
                cell = ws.cell(row=idx, column=col)
                cell.border = self.border
                if col == 1:
                    cell.font = Font(bold=True)
                    cell.alignment = self.center_alignment
                elif col == 3:
                    cell.alignment = self.right_alignment
                else:
                    cell.alignment = self.left_alignment
        
        # Tax Slab Reference
        ws['A15'] = "TAX SLAB REFERENCE (FY 2024-25)"
        ws['A15'].font = self.section_font
        ws['A15'].fill = self.section_fill
        
        slab_headers = ["Income Range", "Tax Rate", "Notes"]
        for col, header in enumerate(slab_headers, 1):
            cell = ws.cell(row=16, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.border = self.border
        
        slabs = [
            ["Up to ₹2,50,000", "NIL", "No tax"],
            ["₹2,50,001 to ₹5,00,000", "5%", ""],
            ["₹5,00,001 to ₹10,00,000", "20%", ""],
            ["Above ₹10,00,000", "30%", ""],
        ]
        
        for idx, slab_data in enumerate(slabs, 17):
            for col, value in enumerate(slab_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                cell.alignment = self.center_alignment if col != 3 else self.left_alignment
        
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 25
        
        return ws
    
    def create_other_income_sheet(self):
        """Create Other Income Sheet"""
        ws = self.workbook.create_sheet("Other Income")
        
        # Header
        ws.merge_cells('A1:D1')
        title = ws['A1']
        title.value = "INCOME FROM OTHER SOURCES"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        headers = ["S.No.", "Source Description", "Amount (₹)", "Remarks"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.border
        
        other_sources = [
            ["1", "Rental Income (after expenses)", 0, ""],
            ["2", "Agricultural Income", 0, ""],
            ["3", "Casual Income", 0, ""],
            ["4", "Winning from Lottery/Games", 0, ""],
            ["5", "Gift/Inheritance (if taxable)", 0, ""],
            ["6", "Income from Royalty", 0, ""],
        ]
        
        for idx, row_data in enumerate(other_sources, 4):
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                if col in [1, 3]:
                    cell.alignment = self.center_alignment
                else:
                    cell.alignment = self.left_alignment
        
        # Total
        total_row = 4 + len(other_sources)
        ws.merge_cells(f'A{total_row}:B{total_row}')
        ws[f'A{total_row}'].value = "TOTAL"
        ws[f'A{total_row}'].font = self.total_font
        ws[f'A{total_row}'].fill = self.total_fill
        ws[f'A{total_row}'].border = self.border
        
        sum_cell = ws.cell(row=total_row, column=3)
        sum_cell.value = f"=SUM(C4:C{total_row-1})"
        sum_cell.font = self.total_font
        sum_cell.fill = self.total_fill
        sum_cell.border = self.border
        
        for col in range(1, 5):
            ws.column_dimensions[get_column_letter(col)].width = 20
        
        return ws
    
    def create_exemptions_sheet(self):
        """Create Exemptions & Special Cases Sheet"""
        ws = self.workbook.create_sheet("Exemptions")
        
        # Header
        ws.merge_cells('A1:C1')
        title = ws['A1']
        title.value = "EXEMPTIONS & SPECIAL CASES"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        # Exemptions List
        exemptions_data = [
            ["Exemption Type", "Section", "Amount (₹)", "Status"],
            ["Agricultural Income Exemption", "10(1)", 0, "Applicable/Not"],
            ["Gratuity Exemption", "10(10)(iii)", 0, "Applicable/Not"],
            ["Leave Encashment Exemption", "10(10)(ii)", 0, "Applicable/Not"],
            ["Interest on Savings Account", "10(15)", 0, "Applicable/Not"],
            ["Dividend Income Exemption", "10(35)", 0, "Applicable/Not"],
            ["Capital Gains Exemption u/s 54", "54", 0, "Applicable/Not"],
            ["Capital Gains Exemption u/s 54F", "54F", 0, "Applicable/Not"],
        ]
        
        for idx, row_data in enumerate(exemptions_data, 3):
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                
                if idx == 3:  # Header row
                    cell.font = self.subheader_font
                    cell.fill = self.subheader_fill
                    cell.alignment = self.center_alignment
                else:
                    if col == 1:
                        cell.alignment = self.left_alignment
                    else:
                        cell.alignment = self.center_alignment
        
        # Notes
        ws.merge_cells('A12:D12')
        notes = ws['A12']
        notes.value = "NOTES: Refer to Income Tax Act, 1961 for detailed conditions and limits"
        notes.font = Font(italic=True, size=9)
        
        for col in range(1, 5):
            ws.column_dimensions[get_column_letter(col)].width = 22
        
        return ws
    
    def create_checklist_sheet(self):
        """Create Compliance Checklist Sheet"""
        ws = self.workbook.create_sheet("Checklist")
        
        # Header
        ws.merge_cells('A1:C1')
        title = ws['A1']
        title.value = "ITR FILING COMPLIANCE CHECKLIST"
        title.font = self.header_font
        title.fill = self.header_fill
        title.alignment = self.center_alignment
        ws.row_dimensions[1].height = 25
        
        headers = ["S.No.", "Compliance Requirement", "Status"]
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.font = self.subheader_font
            cell.fill = self.subheader_fill
            cell.alignment = self.center_alignment
            cell.border = self.border
        
        checklist_items = [
            ["1", "PAN (Permanent Account Number) Available", "□"],
            ["2", "Bank Account Details Verified", "□"],
            ["3", "All Income Documents Collected", "□"],
            ["4", "Expense Receipts & Invoices Organized", "□"],
            ["5", "Investment Proofs Gathered (80C, 80D, etc.)", "□"],
            ["6", "Business Books of Accounts Audited", "□"],
            ["7", "Capital Asset Details Documented", "□"],
            ["8", "Previous Year ITR Review Completed", "□"],
            ["9", "Tax Planning Discussed with CA", "□"],
            ["10", "Form 16 / 16A Reconciled", "□"],
            ["11", "TDS Credit Claimed", "□"],
            ["12", "Advance Tax Paid", "□"],
            ["13", "Form ITR-1/2/3/4 Selected Appropriately", "□"],
            ["14", "Digital Signature / PIN Ready", "□"],
            ["15", "ITR Submitted within Due Date", "□"],
        ]
        
        for idx, row_data in enumerate(checklist_items, 4):
            for col, value in enumerate(row_data, 1):
                cell = ws.cell(row=idx, column=col)
                cell.value = value
                cell.border = self.border
                if col == 1:
                    cell.alignment = self.center_alignment
                elif col == 3:
                    cell.alignment = self.center_alignment
                    cell.font = Font(size=14)
                else:
                    cell.alignment = self.left_alignment
        
        for col in range(1, 4):
            ws.column_dimensions[get_column_letter(col)].width = 22
        
        return ws
    
    def save(self):
        """Save the Excel workbook"""
        self.workbook.save(self.filename)
        print(f"✓ ITR Composition Excel created successfully: {self.filename}")
        return self.filename


def main():
    """Main function to generate ITR Composition Excel"""
    builder = ITRCompositionExcelBuilder("ITR_Composition_FY2024-25.xlsx")
    
    # Create all sheets
    builder.create_summary_sheet()
    builder.create_income_details_sheet()
    builder.create_business_income_sheet()
    builder.create_investment_income_sheet()
    builder.create_deductions_sheet()
    builder.create_tax_calculation_sheet()
    builder.create_other_income_sheet()
    builder.create_exemptions_sheet()
    builder.create_checklist_sheet()
    
    # Save the workbook
    builder.save()
    
    print("\nSheets created:")
    print("1. ITR Summary - Overall income tax return summary")
    print("2. Income Details - All sources of income")
    print("3. Business Income - Business/professional income details")
    print("4. Investment Income - Dividend, interest, and capital gains")
    print("5. Deductions - Section 80 deductions")
    print("6. Tax Calculation - Income tax computation")
    print("7. Other Income - Miscellaneous income sources")
    print("8. Exemptions - Special exemptions and reliefs")
    print("9. Checklist - ITR filing compliance checklist")


if __name__ == "__main__":
    main()
