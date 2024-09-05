import os
import pandas as pd
from fpdf import FPDF
import tkinter as tk
from tkinter import filedialog, messagebox


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
    pdf.cell(0, 8,
             "We request you to verify employment details with our office on email: hr@symbiosystech.com. (+91-0891-2550369)",
             ln=True)

    return pdf


def generate_payslip():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return

    try:
        employee_data_df = pd.read_excel(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        return

    required_columns = [
        'Employee Code', 'Employee Name', 'Designation', 'Date of Joining',
        'Employment Status', 'Basic Pay (Rs.)', 'House Rent Allowance (Rs.)',
        'City Compensatory Allowance (Rs.)', 'Travel Allowance (Rs.)',
        'Food Allowance (Rs.)', 'Performance Incentives (Rs.)',
        'Professional Tax (Rs.)', 'Income Tax (Rs.)',
        'Provident Fund (Rs.)', 'ESI (Rs.)',
        'Leaves-Loss of Pay (Rs.)', 'Others (Rs.)',
        'Gross Pay (Rs.)', 'Deductions (Rs.)', 'Net Pay (Rs.)'
    ]

    if not all(col in employee_data_df.columns for col in required_columns):
        messagebox.showerror("Error", "Excel file is missing required columns.")
        return

    employee_id = employee_id_entry.get()
    if not employee_id.isdigit():
        messagebox.showerror("Error", "Employee ID must be a number.")
        return

    employee_id = int(employee_id)
    employee = employee_data_df.loc[employee_data_df['Employee Code'] == employee_id].squeeze()

    if employee.empty:
        messagebox.showerror("Error", "Employee ID not found.")
        return

    try:
        payslip_pdf = create_payslip(employee)
        output_dir = r'C:\Users\91879\OneDrive\Pictures\internship pdf'
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        output_path = os.path.join(output_dir, f"{employee['Employee Code']}.pdf")
        payslip_pdf.output(output_path)
        messagebox.showinfo("Success", f"Payslip generated successfully: {output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate payslip: {e}")


# Setting up the GUI
root = tk.Tk()
root.title("Payslip Generator")

tk.Label(root, text="Enter Employee ID:").pack(pady=5)
employee_id_entry = tk.Entry(root)
employee_id_entry.pack(pady=5)

tk.Button(root, text="Generate Payslip", command=generate_payslip).pack(pady=20)

root.mainloop()
