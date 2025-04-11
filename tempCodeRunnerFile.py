import pandas as pd
from fpdf import FPDF
import yagmail
import os

# Load environment variables
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
print(f"Email: {SENDER_EMAIL}")
print(f"Password: {EMAIL_PASSWORD}")


# Create payslip directory if it doesn't exist
os.makedirs("payslips", exist_ok=True)

# Read employee data
df = pd.read_excel("employees.xlsx")


# Generate payslip PDFs
for index, row in df.iterrows():
    emp_id = row['Employee ID']
    name = row['Name']
    email = row['Email']
    basic = row['Basic Salary']
    allow = row['Allowances']
    deduct = row['Deductions']
    net = basic + allow - deduct

    # Create PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt=f"Payslip for {name} (ID: {emp_id})", ln=True, align='C')
    pdf.ln(10)
    pdf.cell(200, 10, txt=f"Basic Salary: ${basic:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Allowances: ${allow:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Deductions: ${deduct:.2f}", ln=True)
    pdf.cell(200, 10, txt=f"Net Salary: ${net:.2f}", ln=True)

    pdf_path = f"payslips/{emp_id}.pdf"
    pdf.output(pdf_path)

    # Send Email
    try:
        yag = yagmail.SMTP(SENDER_EMAIL, EMAIL_PASSWORD)
        yag.send(
            to=email,
            subject="Your Payslip for This Month",
            contents="Dear employee,\n\nPlease find your attached payslip for this month.\n\nBest regards,\nHR Department",
            attachments=pdf_path,
        )
        print(f"Email sent to {name} ({email})")
    except Exception as e:
        print(f"Failed to send email to {name}: {e}")