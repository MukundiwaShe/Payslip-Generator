import pandas as pd
import yagmail
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
import getpass


def generate_payslip_pdf(row):
    name = row['Name']
    safe_name = name.replace(" ", "_")
    #position = row['Position']
    salary = row['Basic Salary']
    allowances = row['Allowances']
    deductions = row['Deductions']
    net_pay = salary + allowances - deductions

    filename = f"payslips/{safe_name}_payslip.pdf"
    os.makedirs("payslips", exist_ok=True)

    c = canvas.Canvas(filename, pagesize=A4)
    c.setFont("Helvetica-Bold", 14)
    c.drawString(100, 800, "📄 Monthly Payslip")
    
    c.setFont("Helvetica", 12)
    c.drawString(100, 760, f"Employee Name: {name}")
   # c.drawString(100, 740, f"Position: {position}")
    c.drawString(100, 720, f"Basic Salary: ${salary:.2f}")
    c.drawString(100, 700, f"Allowances: ${allowances:.2f}")
    c.drawString(100, 680, f"Deductions: ${deductions:.2f}")
    c.drawString(100, 660, f"Net Pay: ${net_pay:.2f}")
    c.drawString(100, 640, f"Thank you for your work!")

    c.save()
    return os.path.abspath(filename)


df = pd.read_excel("employees.xlsx")


sender_email = input("Enter your Gmail address: ")
app_password = getpass.getpass("Enter your Gmail app password: ")
yag = yagmail.SMTP(user=sender_email, password=app_password)


for index, row in df.iterrows():
    name = row['Name']
    email = row['Email']


    pdf_path = generate_payslip_pdf(row)

    
    subject = "Your Monthly Payslip"
    body = f"Hi {name},\n\nPlease find attached your payslip for this month.\n\nBest regards,\nHR Team"

    yag.send(
        to=email,
        subject=subject,
        contents=body,
        attachments=pdf_path
    )

    print(f"✅ Payslip sent to {name} ({email})")



    


