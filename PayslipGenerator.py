import pandas as pd
from fpdf import FPDF

# Load the employee data from the Excel file
file_path = 'C:/Users/91879/OneDrive/Pictures/employee_data.xlsx'
employee_data_df = pd.read_excel(file_path)

# Define a function to create the payslip PDF
class PDF(FPDF):
    def header(self):
        self.image('C:/Users/91879/OneDrive/Pictures/LOGO.png', 10, 8, 33)  # Insert the logo on the top left corner
        self.set_font("Arial", 'B', 12)
        self.cell(0, 5, "SYMBIOSYS TECHNOLOGIES", ln=True, align='C')
        self.cell(0, 5, "Plot No 1&2, Hill no-2, IT Park,", ln=True, align='C')
        self.cell(0, 5, "Rushikonda, Visakhapatnam-45", ln=True, align='C')
        self.cell(0, 5, "Ph: 2550369, 2595657", ln=True, align='C')
        self.ln(20)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def create_payslip(employee):
    pdf = PDF()
    pdf.add_page()
    
    # Add the payslip title
    pdf.set_font("Arial", 'B', 12)
    pdf.cell(0, 10, "SALARY STATEMENT FOR THE MONTH OF JANUARY 2024", ln=True, align='C')
    pdf.ln(5)
    
    # Table 1: Employee Code, Name, Designation
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(65, 8, "Employee Code", border=1)
    pdf.cell(65, 8, "Employee Name", border=1)
    pdf.cell(65, 8, "Designation", border=1)
    pdf.ln()
    pdf.set_font("Arial", size=10)
    pdf.cell(65, 8, f"{employee['Employee Code']}", border=1)
    pdf.cell(65, 8, f"{employee['Employee Name']}", border=1)
    pdf.cell(65, 8, f"{employee['Designation']}", border=1)
    pdf.ln(10)
    
    # Table 2: Date of Joining, Employment Status, Statement for the month
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(65, 8, "Date of Joining", border=1)
    pdf.cell(65, 8, "Employment Status", border=1)
    pdf.cell(65, 8, "Statement for the month", border=1)
    pdf.ln()
    pdf.set_font("Arial", size=10)
    date_of_joining = pd.to_datetime(employee['Date of Joining']).strftime('%d-%m-%Y')
    pdf.cell(65, 8, f"{date_of_joining}", border=1)
    pdf.cell(65, 8, f"{employee['Employment Status']}", border=1)
    pdf.cell(65, 8, "", border=1)
    pdf.ln(10)
    
    # Table 3 and 4: Classified Income and Deductions side by side
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(70, 8, "Classified Income", border=1, align='C')
    pdf.cell(30, 8, "Amount (Rs.)", border=1, align='C')
    pdf.cell(60, 8, "Deductions", border=1, align='C')
    pdf.cell(30, 8, "Amount (Rs.)", border=1, align='C')
    pdf.ln()
    pdf.set_font("Arial", size=10)
    
    income_items = [
        "Basic Pay (Rs.)", "House Rent Allowance (Rs.)", 
        "City Compensatory Allowance (Rs.)", "Travel Allowance (Rs.)", 
        "Food Allowance (Rs.)", "Performance Incentives (Rs.)"
    ]
    deduction_items = [
        "Professional Tax (Rs.)", "Income Tax (Rs.)", 
        "Provident Fund (Rs.)", "ESI (Rs.)", 
        "Leaves-Loss of Pay (Rs.)", "Others (Rs.)"
    ]
    
    for income, deduction in zip(income_items, deduction_items):
        pdf.cell(70, 8, f"{income.replace('(Rs.)', '').strip()}", border=1)
        pdf.cell(30, 8, f"Rs. {employee[income]:.2f}", border=1, align='R')
        pdf.cell(60, 8, f"{deduction.replace('(Rs.)', '').strip()}", border=1)
        pdf.cell(30, 8, f"Rs. {employee[deduction]:.2f}", border=1, align='R')
        pdf.ln()
    
    # Add spacing before the Totals section
    pdf.ln(10)
    
    # Totals section
    pdf.set_font("Arial", 'B', 10)
    pdf.cell(70, 8, "GROSS PAY", border=1)
    pdf.cell(30, 8, f"Rs. {employee['Gross Pay (Rs.)']:.2f}", border=1, align='R')
    pdf.cell(60, 8, "DEDUCTIONS", border=1)
    pdf.cell(30, 8, f"Rs. {employee['Deductions (Rs.)']:.2f}", border=1, align='R')
    pdf.ln()
    pdf.cell(100, 8, "NET PAY", border=1)
    pdf.cell(80, 8, f"Rs. {employee['Net Pay (Rs.)']:.2f}", border=1, align='R')
    pdf.ln(20)
    
    # Footer section with added spacing
    pdf.cell(0, 8, "AUTHORISED SIGNATORY", ln=True)
    pdf.ln(20)  # Added more spacing here
    pdf.cell(0, 8, "Durgaaprasadh,", ln=True)
    pdf.cell(0, 8, "H.R Executive", ln=True)
    pdf.ln(10)
    pdf.set_font('Arial', 'I', 8)
    pdf.cell(0, 8, "We request you to verify employment details with our office on email: hr@symbiosystech.com. (+91-0891-2550369)", ln=True)

    return pdf

# Example employee data to match the format in the image
employee_id = int(input("Enter the Employee ID for which the payslip should be generated: "))
employee = employee_data_df.loc[employee_data_df['Employee Code'] == employee_id].squeeze()

if not employee.empty:
    payslip_pdf = create_payslip(employee)
    employee_code = employee['Employee Code']
    output_path = f"C:/Users/91879/OneDrive/Pictures/internship pdf/{employee_code}.pdf"
    payslip_pdf.output(output_path)
    print(f"C:/Users/91879/OneDrive/Pictures/internship pdf/{employee_code}.pdf")
else:
    print("Employee ID not found.")
