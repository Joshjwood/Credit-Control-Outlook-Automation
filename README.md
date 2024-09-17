# Credit-Control-Outlook-Automation

Introduction

This Python script automates the process of sending credit control emails to customers with overdue payments. Using Microsoft Outlook, the script reads customer data from an Excel spreadsheet, generates personalized emails based on predefined templates, attaches any relevant overdue statements, and sends these emails to the respective customers.

The script is designed to handle multiple customers at once, selecting one of several email templates at random to avoid sending repetitive messages to the same customer.
How It Works

    Reads Input Data: The script reads customer information from an Excel spreadsheet (overdue_customers.xlsx). This file must contain specific columns with details about each customer's overdue payment.
    Generates Personalized Emails: Using the data from the spreadsheet, the script generates a personalized email for each customer. The email includes the amount overdue, the due date (formatted in a friendly way), and the number of days the payment has been overdue.
    Finds Attachments: For each customer, the script searches for relevant overdue statements in a specified folder (attachments) and attaches them to the email if found.
    Sends Emails via Outlook: The script uses the Microsoft Outlook application to send the emails to the customers.
    Console Output: Throughout the process, the script provides detailed output in the console to help you monitor which customers are being processed and the status of each email sent.

Input Requirements
1. Excel Spreadsheet: overdue_customers.xlsx

    The script requires an Excel file named overdue_customers.xlsx located in the same directory as the script.
    The spreadsheet must contain four columns in the following order:
        Column A: Overdue Date - The date when the payment was originally due. The date should be in a recognizable date format (e.g., 2023-05-19, 07/15/2024, etc.). The script will reformat this date for the emails.
        Column B: Overdue Amount - The amount of money that is overdue. This can be a numeric value representing the currency (e.g., 100.50, 2500).
        Column C: Customer Name - The name of the customer. This should be a string and will be used to personalize the email (e.g., John Doe).
        Column D: Customer Email - The email address of the customer. This must be a valid email address (e.g., johndoe@example.com).

Example Excel Format:
Overdue Date	Overdue Amount	Customer Name	Customer Email
2024-07-15	500.00	John Doe	johndoe@example.com
2024-06-20	250.50	Jane Smith	janesmith@example.com
2. Attachments Folder: attachments

    The script looks for a folder named attachments in the same directory where the script is located.
    This folder should contain files relevant to the customers listed in the spreadsheet. The script will attach any files found in this folder that include the customer's name in the filename.
    If no attachments are found for a customer, the script will still send the email without attachments.

Example File Structure:

markdown

- project_directory/
  - overdue_customers.xlsx
  - attachments/
    - John Doe_statement_July2024.pdf
    - Jane Smith_invoice_062024.pdf
  - credit_control_outlook_personalized.py

How to Use

    Prepare the Spreadsheet: Ensure that the Excel file overdue_customers.xlsx is formatted correctly and placed in the same directory as the script.
    Prepare the Attachments Folder: Place all relevant attachment files in the attachments folder, ensuring that filenames contain the customer's name for the script to recognize and attach them correctly.
    Run the Script: Execute the script in your preferred Python environment (e.g., PyCharm). The script will process each customer in the spreadsheet, generate a personalized email, attach any relevant files, and send the email through Outlook.
    Monitor Console Output: The script will provide detailed output in the console, showing the progress and status of each email sent.

Prerequisites

    Python 3.x: Make sure Python is installed on your system.
    Microsoft Outlook: The script uses the Outlook application to send emails, so Outlook must be installed and configured on your system.
    Python Packages: The script requires pandas and pywin32 for handling Excel files and interacting with Outlook, respectively. Install these packages using:

    bash

    pip install pandas pywin32 openpyxl

Notes

    Ensure that Microsoft Outlook is set up with the correct account from which you want to send the emails.
    Test the script with a small batch of customer data before using it for a larger mailing.
    Always review the email templates and adjust them to fit your company's tone and communication style.
