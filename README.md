# Payslip-Email-Sender

## Overview
This repository contains a Streamlit application that automates the process of sending payslips via email. The application reads a PDF file containing multiple payslips, extracts the necessary details, and sends personalized emails with individual payslips as attachments to employees.

## Features
* Upload a PDF file containing multiple payslips.
* Upload an Excel sheet with employee details.
* Upload an email template in a text file.
* Extract and parse payslip information.
* Send personalized emails with payslip attachments to employees.
* Provides success and error messages within the Streamlit interface.

## Installation
### 1. Clone the repository
![image](https://github.com/Ashar18/Payslip-Email-Sender/assets/64865488/6f6fff02-4a42-43ad-b9c0-6491ffe709a5)

### 2. Create and activate a virtual environment
![image](https://github.com/Ashar18/Payslip-Email-Sender/assets/64865488/e6d55d0b-bdab-4400-85f7-f71e81728c44)

### 3. Install the required packages
![image](https://github.com/Ashar18/Payslip-Email-Sender/assets/64865488/d67e4dc8-014a-43e7-a3c2-164efa75241c)

## Usage
### 1. Run the Streamlit app
![image](https://github.com/Ashar18/Payslip-Email-Sender/assets/64865488/8d020e48-dff8-4bda-a3c1-1466edc58695)

### 2. Input the necessary details in the Streamlit app interface:
* Email address and password.
* Upload the payslip PDF file.
* Upload the employee details Excel file.
* Upload the email template text file.
* Enter the email subject, month, and year.
* Send Emails

Click the "Send Email" button to start the process. The app will read the payslip PDF, extract the relevant details, and send personalized emails to each employee.

## File Descriptions
"**app.py**"
Main script to run the Streamlit application. It contains the following key functions:
**send**: Sends an email with a PDF attachment.
**send_mail**: Main function to handle the extraction and email sending process.
**main**: Streamlit interface for user input and triggering the email sending process.

## Dependencies
**os**

**re**
**pdfplumber**
**pandas**
**numpy**
**PyPDF2**
**smtplib**
**email.mime.multipart**
**email.mime.base**
**email.mime.text**
**streamlit**
**shutil**

## Note
* Ensure your email provider allows less secure apps or generate an app-specific password for SMTP.
* The provided PDF file must contain payslips in a format that the script can parse.
* Customize the email template text file as needed using placeholders such as {Employee_Name}, {Net_Salary_Amount}, and {Month}.

## Contributing
Feel free to submit issues and enhancement requests.



Developed by **Ashar Nadeem**

