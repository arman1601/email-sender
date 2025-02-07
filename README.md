# email-sender
Email Automation Script - User Guide

Overview
This Python script automates sending personalized emails by extracting recipient details from an Excel file and merging them into a Word document template. The generated document is then sent as an email attachment to each recipient.

How It Works

Reads an Excel file (data1.xlsx) containing recipient details.
Loads a Word template (template.docx) with placeholders (Data1, Data2, etc.).
Replaces placeholders in the Word file with actual values from Excel.
Saves a personalized Word document for each recipient.
Sends an email with the generated document as an attachment.
Deletes the temporary Word file after sending the email.

Prerequisites
Install required dependencies:
sh

pip install smtplib openpyxl python-docx

Configure SMTP settings (e.g., Yandex, Gmail) in your email provider settings.

Ensure Excel and Word template files are correctly formatted.

Prerequisites 

The Email column in the Excel file is mandatory. The script will not send emails if this column is missing.
The Excel file can contain multiple variables (e.g., Data1, Data2, Data3, etc.).
For optimal performance, we recommend not exceeding 15 variables in the Excel file.
All variable names must start with "Data" (e.g., Data1, Data2, Data3), ensuring correct replacement in the Word template.
