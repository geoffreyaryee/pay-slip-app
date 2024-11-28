import pandas as pd
import matplotlib.pyplot as plt
from docx import Document

doc=Document("payslip template.docx")


# Function to calculate PAYE
def calculate_paye(taxable_income):
    tax_free_threshold = 365
    band_1_limit = 110
    band_2_limit = 130
    band_3_limit = 3000
    band_4_threshold = 3605

    rate_1 = 0.05
    rate_2 = 0.10
    rate_3 = 0.175
    rate_4 = 0.25

    tax = 0

    if taxable_income <= tax_free_threshold:
        return 0

    if taxable_income > tax_free_threshold:
        band_1_taxable = min(band_1_limit, taxable_income - tax_free_threshold)
        tax += band_1_taxable * rate_1

    if taxable_income > tax_free_threshold + band_1_limit:
        band_2_taxable = min(band_2_limit, taxable_income - tax_free_threshold - band_1_limit)
        tax += band_2_taxable * rate_2

    if taxable_income > band_4_threshold:
        band_3_taxable = min(band_3_limit, taxable_income - tax_free_threshold - band_1_limit - band_2_limit)
        tax += band_3_taxable * rate_3

    if taxable_income > band_4_threshold:
        band_4_taxable = taxable_income - band_4_threshold
        tax += band_4_taxable * rate_4

    return round(tax, 2)

# Function to process the employee Excel file and calculate PAYE
def process_employee_data(file_path):
    # Read employee data from Excel
    df = pd.read_excel(file_path)

    # Calculate gross salary (basic + allowance)
    df['Gross Salary'] = df['basic'] + df['allowance']

    # Calculate SSNIT contribution (5.5% of Gross Salary)
    df['SSNIT'] = df['Gross Salary'] * 0.055

    # Calculate taxable income (Gross Salary - SSNIT)
    df['Taxable Income'] = df['Gross Salary'] - df['SSNIT']

    # Calculate PAYE on the taxable income
    df['PAYE'] = df.apply(lambda row: calculate_paye(row['Taxable Income']), axis=1)

    # Subtract PAYE and other deductions to get net salary
    df['Net Salary'] = df['Gross Salary'] - df['PAYE'] - df['deduction']

    # Generate and save individual PDFs for each employee
    for index, row in df.iterrows():
        employee_data = row.to_frame().T  # Convert the row to a dataframe
        output_file_pdf = f"payslip_{row['name']}_{row['month']}_{row['year']}.pdf"
        generate_employee_pdf(employee_data, output_file_pdf)

    # Save the updated data with PAYE calculation (optional)
    output_file_excel = 'employee_payslip_with_paye.xlsx'
    df.to_excel(output_file_excel, index=False)

    return output_file_excel,df

# Function to generate PDF for each employee
def generate_employee_pdf(employee_data, output_file_pdf):
    fig, ax = plt.subplots(figsize=(12, 2))  # Set size for individual employee's PDF

    # Hide axes
    ax.xaxis.set_visible(False)
    ax.yaxis.set_visible(False)
    ax.set_frame_on(False)

    # Create the table for the employee
    table = ax.table(cellText=employee_data.values, colLabels=employee_data.columns, cellLoc='center', loc='center')

    # Auto-scale the table to fit the PDF
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1.2, 1.2)  # Scale the table to fit

    # Save the table as a PDF
    plt.savefig(output_file_pdf, bbox_inches='tight', dpi=300)

    # Close the plot to free up memory
    plt.close()

# Example: Path to the Excel file (uploaded by user)
file_path = 'emplyee data.xlsx'
process_employee_data(file_path)