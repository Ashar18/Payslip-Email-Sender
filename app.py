import os
import re
import pdfplumber
import pandas as pd
import numpy as np
from PyPDF2 import PdfReader, PdfWriter

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

import streamlit as st
# import base64
import shutil


def send(subject, message, sender_email, receiver_email, attachment_filename, smtp_server, smtp_port, smtp_username, smtp_password):
    # Create a multipart message
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = subject

    # Attach the message body
    msg.attach(MIMEText(message, 'plain'))

    # Attach the PDF file
    with open(attachment_filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {attachment_filename}")
    msg.attach(part)

    # Connect to the SMTP server
    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
        try:
        #   print(smtp_username,smtp_password)
          server.login(smtp_username, smtp_password)
          server.sendmail(sender_email, receiver_email, msg.as_string())
          st.success(f'Email sent to {receiver_email}')


        except:
          st.error("Wrong username and password !!!")
          return False

def send_mail(email_id,email_pass,payslip_pdf,emp_sheet,email_template,subject,month,year):
    data = pd.read_excel(emp_sheet)
    
    pdf_file_path = payslip_pdf
    month = month
    year = year

    out_folder = f'{month}-{year}'
    if os.path.exists(out_folder):
        shutil.rmtree(out_folder)

    # with open(pdf_file_path, "rb") as pdf_file:
    # pdf_reader = PdfReader(pdf_file)
    pdf_reader = PdfReader(pdf_file_path)
    for page_num, page in enumerate(pdf_reader.pages):
        os.makedirs(out_folder, exist_ok=True)
        # if os.path.exists(out_folder):
            # shutil.rmtree(out_folder)
        
        # Create the directory
        # os.makedirs(out_folder)
        # print(f"Created directory: {directory}")

        output_file_name = f"{out_folder}/{page_num}.pdf"
        pdf_writer = PdfWriter()
        pdf_writer.add_page(page)
        with open(output_file_name, 'wb') as output_file:
            pdf_writer.write(output_file)
                
                
    ### pdf to jpg


    subject = subject

    emailfile = email_template  # Path to your text file



    sender_email = email_id
    receiver_email = ''
    output_file_name = ''

    attachment_filename = '' # Change to the name of your PDF file
    smtp_server = 'smtp.gmail.com'
    smtp_port = 465
    smtp_username = email_id
    smtp_password = email_pass
    count = 0
    dic = {}
    


    for i in os.listdir(out_folder):
        with pdfplumber.open(f'{out_folder}/{i}') as pdf:
            text = ''
            # Extract text from each page
            for page in pdf.pages:
                text = page.extract_text()
    #             print(text)

            try:
                emp_code = re.search(r'EmployeeCode(\d+)', text.replace(':','').replace(' ','')).group(1)
                emp_salary = re.search(r'NETAMOUNTPAYABLE([\d,.]+)', text.upper().replace(':','').replace(' ','')).group(1)
            except:
                emp_code = ''
                emp_salary = ''

            if emp_code!='' and emp_salary !='':
                dic[emp_code] = emp_salary
                try:
                    emp_name = data[data['Emp Code.'].astype(str)==emp_code]['Employee Name'].values[0]
                except:
                    emp_name = ''
                try:
                    emp_email = data[data['Emp Code.'].astype(str)==emp_code]['Email Address'].values[0]
                except:
                    emp_email = ''


                # with open(emailfile, "r") as file:
                #     message = file.read()
                message = emailfile.getvalue().decode("utf-8")
                # message = emailfile.read()
                message = message.format(Employee_Name=emp_name, Net_Salary_Amount=emp_salary,Month=month)
                receiver_email = emp_email
                output_file_name = f'{out_folder}/{emp_code}-{emp_name} Payslip.pdf'

        print(emp_name,emp_code)
        os.rename(f'{out_folder}/{i}',output_file_name)
        attachment_filename = output_file_name

                  # Send email
        if receiver_email!='':
            a= send(subject, message, sender_email, receiver_email, attachment_filename, smtp_server, smtp_port, smtp_username, smtp_password)
            if a == False:
              break
            
            count += 1
    shutil.rmtree(out_folder)        
    return count




# from send_mail_function import send_mail  # Assuming your function is in a file named send_mail_function.py

def main():
    st.title("PaySlip Email Sender")
    
    # Input fields
    email_id = st.text_input("Enter your email address:")
    email_pass = st.text_input("Enter your email password:", type="password")
    payslip_pdf = st.file_uploader("Upload the payslip PDF file:", type=["pdf"])
    emp_sheet = st.file_uploader("Upload the employee sheet (Excel file):", type=["xlsx","xls"])
    email_template = st.file_uploader("Upload the email template (text file):", type=["txt"])
    subject = st.text_input("Enter email subject:")
    month = st.text_input("Enter month:")
    year = st.text_input("Enter year:")
    
    if st.button("Send Email"):
        # try:
            if email_id and email_pass and payslip_pdf and emp_sheet and email_template and subject and month and year:
                count = send_mail(email_id, email_pass, payslip_pdf, emp_sheet, email_template, subject, month, year)
                st.success(f"Done! {count} emails sent successfully!")
            else:
                st.error("Please fill in all the required fields.")
        # except:
            # st.error("Something Went Wrong.")

if __name__ == "__main__":
    main()

