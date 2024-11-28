import pandas as pd
from docx import Document
import pypandoc
from tkinter import Tk, Label, Button, filedialog, messagebox
import os
import comtypes.client  # For Word to PDF conversion


# Ghana PAYE tax calculation function
def calculate_paye(taxable):
    tax = 0
    bands = [
        (490, 0.0),  # First GHS 528 at 0%
        (110, 0.05),  # Next GHS 100 at 5%
        (130, 0.10),  # Next GHS 160 at 10%
        (3166.67, 0.175),  # Next GHS 3020 at 17.5%
        (16302, 0.25),  # Next GHS 16,302 at 25%
        (float('inf'), 0.30),  # Amount above GHS 20,110 at 30%
    ]

    remaining_income = taxable
    for band_limit, rate in bands:
        if remaining_income > band_limit:
            tax += band_limit * rate
            remaining_income -= band_limit
        else:
            tax += remaining_income * rate
            break
    return tax

# Function to perform gross pay, PAYE tax, and net pay calculations
def calculate_net_pay(row):
    gross_pay = row['Basic Salary'] + row['Allowances']
    ssnit= gross_pay*0.055
    taxable_income=gross_pay-ssnit
    paye = calculate_paye(taxable_income)  # Calculate PAYE tax
    net_pay = taxable_income - paye - row['Deductions']
    return gross_pay, paye, net_pay,ssnit,taxable_income

# Function to generate payslip from a Word template
def generate_payslip_from_template(employee_data, template_path):
    # Load the Word template
    doc = Document(template_path)

    # Replace placeholders with employee data
    for paragraph in doc.paragraphs:
        if '{Name}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Name}', employee_data['Name'])
        if '{Basic Salary}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Basic Salary}', "{:.2f}".format(employee_data['Basic Salary']))
        if '{Allowances}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Allowances}', "{:.2f}".format(employee_data['Allowances']))
        if '{Gross Pay}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Gross Pay}', "{:.2f}".format(employee_data['Gross Pay']))
        if '{PAYE}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{PAYE}', "{:.2f}".format(employee_data['PAYE']))
        if '{Deductions}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Deductions}', "{:.2f}".format(employee_data['Deductions']))
        if '{Net Pay}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Net Pay}', "{:.2f}".format(employee_data['Net Pay']))
        if '{SSNIT}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{SSNIT}', "{:.2f}".format(employee_data['SSNIT']))
        if '{Taxable}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Taxable}', "{:.2f}".format(employee_data['Taxable']))
        if '{Year}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Year}', str(employee_data['Year']))
        if '{Role}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Role}', str(employee_data['Role']))
        if '{Month}' in paragraph.text:
            paragraph.text = paragraph.text.replace('{Month}', str(employee_data['Month']))

    # Save the generated payslip as a new Word document
    output_path = f"payslip_{employee_data['Name']}.docx"
    doc.save(output_path)
    print(f"Pay slip generated: {output_path}")

# Function to load the Excel file and generate pay slips
def load_file():
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", ".xlsx;.xls")])
    if file_path:
        # Load the Excel data
        df = pd.read_excel(file_path)
        
        # Calculate Gross Pay, PAYE, and Net Pay
        df['Gross Pay'], df['PAYE'], df['Net Pay'],df['SSNIT'],df['Taxable'] = zip(*df.apply(calculate_net_pay, axis=1))
        
        # Select the Word template for generating payslips
        template_path = filedialog.askopenfilename(filetypes=[("Word files", ".docx")])
        if not template_path:
            messagebox.showwarning("Warning", "No template selected!")
            return
        
        # Generate a payslip for each employee
        for index, row in df.iterrows():
            employee_data = row.to_dict()
            generate_payslip_from_template(employee_data, template_path)
        
        messagebox.showinfo("Success", "Pay slips generated successfully!")
    else:
        messagebox.showwarning("Warning", "No file selected!")

# GUI Setup
root = Tk()
root.title("Pay Slip Generator")

label = Label(root, text="Select Excel File to Generate Pay Slips")
label.pack(pady=20)

button = Button(root, text="Load Excel File", command=load_file)
button.pack(pady=10)

root.mainloop()
