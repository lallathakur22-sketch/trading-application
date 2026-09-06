"""
Salary Entry and Complete Accounting Module
Complete salary, EPF, ESI, TDS entry with accounting integration
Based on accounting standards for Indian companies
"""

class SalaryEntry:
    """
    Complete Salary Entry Module
    Handles: Gross Salary, EPF, ESI, TDS, and Net Salary calculations
    """
    
    def __init__(self, employee_id, employee_name, gross_salary):
        self.employee_id = employee_id
        self.employee_name = employee_name
        self.gross_salary = gross_salary
        self.employee_epf = 0
        self.employee_esi = 0
        self.tds = 0
        self.employer_epf = 0
        self.employer_esi = 0
        self.net_salary = 0
    
    def calculate_employee_epf(self, epf_rate=0.12):
        """Calculate Employee EPF (12% of gross salary)"""
        self.employee_epf = self.gross_salary * epf_rate
        return self.employee_epf
    
    def calculate_employer_epf(self, epf_rate=0.12):
        """Calculate Employer EPF (12% of gross salary)"""
        self.employer_epf = self.gross_salary * epf_rate
        return self.employer_epf
    
    def calculate_employee_esi(self, esi_rate=0.0075):
        """Calculate Employee ESI (0.75% of gross salary - if applicable)"""
        # ESI is applicable on wages up to ₹21,000 per month
        if self.gross_salary <= 21000:
            self.employee_esi = self.gross_salary * esi_rate
        return self.employee_esi
    
    def calculate_employer_esi(self, esi_rate=0.0325):
        """Calculate Employer ESI (3.25% of gross salary - if applicable)"""
        # ESI is applicable on wages up to ₹21,000 per month
        if self.gross_salary <= 21000:
            self.employer_esi = self.gross_salary * esi_rate
        return self.employer_esi
    
    def calculate_tds(self, tds_rate=0.05):
        """Calculate TDS (Tax Deducted at Source - 5% as per example)"""
        self.tds = self.gross_salary * tds_rate
        return self.tds
    
    def calculate_net_salary(self):
        """Calculate Net Salary = Gross - Employee EPF - Employee ESI - TDS"""
        self.net_salary = (
            self.gross_salary 
            - self.employee_epf 
            - self.employee_esi 
            - self.tds
        )
        return self.net_salary
    
    def get_salary_summary(self):
        """Get complete salary summary"""
        return {
            'employee_id': self.employee_id,
            'employee_name': self.employee_name,
            'gross_salary': self.gross_salary,
            'employee_epf': self.employee_epf,
            'employee_esi': self.employee_esi,
            'tds': self.tds,
            'net_salary': self.net_salary,
            'employer_epf': self.employer_epf,
            'employer_esi': self.employer_esi,
            'total_employer_contribution': self.employer_epf + self.employer_esi
        }
    
    def print_salary_slip(self):
        """Print salary slip format"""
        print("\n" + "="*60)
        print("SALARY SLIP")
        print("="*60)
        print(f"Employee ID: {self.employee_id}")
        print(f"Employee Name: {self.employee_name}")
        print("-"*60)
        print(f"Gross Salary:           ₹{self.gross_salary:>12,.2f}")
        print("-"*60)
        print("DEDUCTIONS:")
        print(f"  Employee EPF:         ₹{self.employee_epf:>12,.2f}")
        print(f"  Employee ESI:         ₹{self.employee_esi:>12,.2f}")
        print(f"  TDS:                  ₹{self.tds:>12,.2f}")
        print("-"*60)
        print(f"Net Salary:             ₹{self.net_salary:>12,.2f}")
        print("="*60)
        print("EMPLOYER CONTRIBUTION (Additional Expense):")
        print(f"  Employer EPF:         ₹{self.employer_epf:>12,.2f}")
        print(f"  Employer ESI:         ₹{self.employer_esi:>12,.2f}")
        print(f"  Total:                ₹{self.employer_epf + self.employer_esi:>12,.2f}")
        print("="*60 + "\n")


class AccountingEntry:
    """
    Handles complete accounting entries for salary processing
    Including journal entries for salary provision, EPF, ESI, and TDS
    """
    
    def __init__(self, company_name, financial_month):
        self.company_name = company_name
        self.financial_month = financial_month
        self.journal_entries = []
    
    def add_salary_entry(self, salary_obj):
        """
        Entry 1: Salary Provision Entry
        Dr. Salary Expense A/c         - Gross Salary
        Dr. Employer EPF Contribution A/c - Employer EPF
        Dr. Employer ESI Contribution A/c - Employer ESI
            To Employee EPF Payable A/c           - Employee EPF
            To Employee ESI Payable A/c           - Employee ESI
            To TDS Payable A/c                    - TDS
            To Salary Payable A/c                 - Net Salary
        """
        entry = {
            'entry_type': 'Salary Provision Entry',
            'description': f'Salary for {salary_obj.employee_name} - {self.financial_month}',
            'debits': [
                {'account': 'Salary Expense A/c', 'amount': salary_obj.gross_salary},
                {'account': 'Employer EPF Contribution A/c', 'amount': salary_obj.employer_epf},
                {'account': 'Employer ESI Contribution A/c', 'amount': salary_obj.employer_esi}
            ],
            'credits': [
                {'account': 'Employee EPF Payable A/c', 'amount': salary_obj.employee_epf},
                {'account': 'Employee ESI Payable A/c', 'amount': salary_obj.employee_esi},
                {'account': 'TDS Payable A/c', 'amount': salary_obj.tds},
                {'account': 'Salary Payable A/c', 'amount': salary_obj.net_salary}
            ],
            'total_amount': salary_obj.gross_salary + salary_obj.employer_epf + salary_obj.employer_esi
        }
        self.journal_entries.append(entry)
        return entry
    
    def add_net_salary_payment_entry(self, salary_obj):
        """
        Entry 2: Payment of Net Salary to Employee
        Dr. Salary Payable A/c         - Net Salary
            To Bank A/c                 - Net Salary
        """
        entry = {
            'entry_type': 'Net Salary Payment',
            'description': f'Payment to {salary_obj.employee_name}',
            'debits': [
                {'account': 'Salary Payable A/c', 'amount': salary_obj.net_salary}
            ],
            'credits': [
                {'account': 'Bank A/c', 'amount': salary_obj.net_salary}
            ],
            'total_amount': salary_obj.net_salary
        }
        self.journal_entries.append(entry)
        return entry
    
    def add_epf_deposit_entry(self, salary_obj):
        """
        Entry 3: EPF Deposit
        Dr. Employee EPF Payable A/c   - Employee EPF
        Dr. Employer EPF Payable A/c   - Employer EPF
            To Bank A/c                 - Total EPF (Employee + Employer)
        """
        total_epf = salary_obj.employee_epf + salary_obj.employer_epf
        entry = {
            'entry_type': 'EPF Deposit',
            'description': f'EPF deposit for {salary_obj.employee_name}',
            'debits': [
                {'account': 'Employee EPF Payable A/c', 'amount': salary_obj.employee_epf},
                {'account': 'Employer EPF Payable A/c', 'amount': salary_obj.employer_epf}
            ],
            'credits': [
                {'account': 'Bank A/c', 'amount': total_epf}
            ],
            'total_amount': total_epf
        }
        self.journal_entries.append(entry)
        return entry
    
    def add_esi_deposit_entry(self, salary_obj):
        """
        Entry 4: ESI Deposit
        Dr. Employee ESI Payable A/c   - Employee ESI
        Dr. Employer ESI Payable A/c   - Employer ESI
            To Bank A/c                 - Total ESI (Employee + Employer)
        """
        total_esi = salary_obj.employee_esi + salary_obj.employer_esi
        entry = {
            'entry_type': 'ESI Deposit',
            'description': f'ESI deposit for {salary_obj.employee_name}',
            'debits': [
                {'account': 'Employee ESI Payable A/c', 'amount': salary_obj.employee_esi},
                {'account': 'Employer ESI Payable A/c', 'amount': salary_obj.employer_esi}
            ],
            'credits': [
                {'account': 'Bank A/c', 'amount': total_esi}
            ],
            'total_amount': total_esi
        }
        self.journal_entries.append(entry)
        return entry
    
    def add_tds_deposit_entry(self, salary_obj):
        """
        Entry 5: TDS Deposit to Government
        Dr. TDS Payable A/c            - TDS Amount
            To Bank A/c                 - TDS Amount
        """
        entry = {
            'entry_type': 'TDS Deposit',
            'description': f'TDS deposit for {salary_obj.employee_name}',
            'debits': [
                {'account': 'TDS Payable A/c', 'amount': salary_obj.tds}
            ],
            'credits': [
                {'account': 'Bank A/c', 'amount': salary_obj.tds}
            ],
            'total_amount': salary_obj.tds
        }
        self.journal_entries.append(entry)
        return entry
    
    def print_journal_entries(self):
        """Print all journal entries"""
        print("\n" + "="*80)
        print(f"JOURNAL ENTRIES - {self.company_name}")
        print(f"Financial Month: {self.financial_month}")
        print("="*80)
        
        for idx, entry in enumerate(self.journal_entries, 1):
            print(f"\n[Entry {idx}] {entry['entry_type']}")
            print("-"*80)
            print(f"Description: {entry['description']}")
            print("-"*80)
            print("Debit Entries:")
            for debit in entry['debits']:
                print(f"  {debit['account']:<40} ₹{debit['amount']:>12,.2f}")
            print("\nCredit Entries:")
            for credit in entry['credits']:
                print(f"  {credit['account']:<40} ₹{credit['amount']:>12,.2f}")
            print("-"*80)
            print(f"Total Amount: ₹{entry['total_amount']:,.2f}")
            print("="*80)
    
    def get_ledger_heads(self):
        """Get all affected ledger heads"""
        ledger_heads = {
            'Assets': [
                'Bank A/c'
            ],
            'Liabilities': [
                'Employee EPF Payable A/c',
                'Employee ESI Payable A/c',
                'TDS Payable A/c',
                'Salary Payable A/c'
            ],
            'Expenses': [
                'Salary Expense A/c',
                'Employer EPF Contribution A/c',
                'Employer ESI Contribution A/c'
            ],
            'Other Payables': [
                'Employer EPF Payable A/c',
                'Employer ESI Payable A/c'
            ]
        }
        return ledger_heads
    
    def print_ledger_classification(self):
        """Print ledger head classification"""
        print("\n" + "="*80)
        print("LEDGER HEAD CLASSIFICATION")
        print("="*80)
        
        ledger_heads = self.get_ledger_heads()
        
        for category, accounts in ledger_heads.items():
            print(f"\n{category}:")
            for account in accounts:
                print(f"  • {account}")
        
        print("\n" + "="*80)


class SalaryProcessing:
    """
    Complete Salary Processing Module
    Manages multiple employees and generates complete accounting entries
    """
    
    def __init__(self, company_name, financial_month):
        self.company_name = company_name
        self.financial_month = financial_month
        self.employees = []
        self.accounting = AccountingEntry(company_name, financial_month)
    
    def add_employee(self, employee_id, employee_name, gross_salary):
        """Add employee and calculate salary components"""
        salary = SalaryEntry(employee_id, employee_name, gross_salary)
        
        # Calculate all components
        salary.calculate_employee_epf()
        salary.calculate_employer_epf()
        salary.calculate_employee_esi()
        salary.calculate_employer_esi()
        salary.calculate_tds()
        salary.calculate_net_salary()
        
        self.employees.append(salary)
        
        # Add accounting entries
        self.accounting.add_salary_entry(salary)
        self.accounting.add_net_salary_payment_entry(salary)
        self.accounting.add_epf_deposit_entry(salary)
        self.accounting.add_esi_deposit_entry(salary)
        self.accounting.add_tds_deposit_entry(salary)
        
        return salary
    
    def process_all_employees(self):
        """Process all employees and generate summaries"""
        total_gross = 0
        total_employee_epf = 0
        total_employer_epf = 0
        total_employee_esi = 0
        total_employer_esi = 0
        total_tds = 0
        total_net = 0
        
        for employee in self.employees:
            summary = employee.get_salary_summary()
            total_gross += summary['gross_salary']
            total_employee_epf += summary['employee_epf']
            total_employer_epf += summary['employer_epf']
            total_employee_esi += summary['employee_esi']
            total_employer_esi += summary['employer_esi']
            total_tds += summary['tds']
            total_net += summary['net_salary']
        
        return {
            'total_gross': total_gross,
            'total_employee_epf': total_employee_epf,
            'total_employer_epf': total_employer_epf,
            'total_employee_esi': total_employee_esi,
            'total_employer_esi': total_employer_esi,
            'total_tds': total_tds,
            'total_net': total_net,
            'total_employer_contribution': total_employer_epf + total_employer_esi,
            'employee_count': len(self.employees)
        }
    
    def print_payroll_summary(self):
        """Print complete payroll summary"""
        summary = self.process_all_employees()
        
        print("\n" + "="*80)
        print(f"PAYROLL SUMMARY - {self.company_name}")
        print(f"For the Month: {self.financial_month}")
        print("="*80)
        
        print(f"\nTotal Employees: {summary['employee_count']}")
        print("-"*80)
        print(f"Total Gross Salary:                    ₹{summary['total_gross']:>12,.2f}")
        print("-"*80)
        print("EMPLOYEE DEDUCTIONS:")
        print(f"  Employee EPF:                        ₹{summary['total_employee_epf']:>12,.2f}")
        print(f"  Employee ESI:                        ₹{summary['total_employee_esi']:>12,.2f}")
        print(f"  TDS:                                 ₹{summary['total_tds']:>12,.2f}")
        print("-"*80)
        print(f"Total Net Salary Payable:              ₹{summary['total_net']:>12,.2f}")
        print("="*80)
        print("EMPLOYER CONTRIBUTION (Additional Expense):")
        print(f"  Employer EPF:                        ₹{summary['total_employer_epf']:>12,.2f}")
        print(f"  Employer ESI:                        ₹{summary['total_employer_esi']:>12,.2f}")
        print("-"*80)
        print(f"Total Employer Contribution:           ₹{summary['total_employer_contribution']:>12,.2f}")
        print("="*80 + "\n")


def main():
    """Main function demonstrating salary entry and accounting"""
    
    # Initialize salary processing
    processor = SalaryProcessing("ABC Trading Company", "September 2024")
    
    # Add employees (example data)
    print("Processing Employee Salaries...\n")
    
    emp1 = processor.add_employee("EMP001", "Rajesh Kumar", 30000)
    emp1.print_salary_slip()
    
    emp2 = processor.add_employee("EMP002", "Priya Singh", 25000)
    emp2.print_salary_slip()
    
    emp3 = processor.add_employee("EMP003", "Amit Patel", 18000)
    emp3.print_salary_slip()
    
    # Print payroll summary
    processor.print_payroll_summary()
    
    # Print journal entries
    processor.accounting.print_journal_entries()
    
    # Print ledger classification
    processor.accounting.print_ledger_classification()


if __name__ == "__main__":
    main()
