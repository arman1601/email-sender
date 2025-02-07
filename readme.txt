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