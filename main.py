import pandas as pd
import os
import random
import win32com.client
from datetime import datetime

# Email templates
EMAIL_TEMPLATES = [
    "Dear {customer_name},\n\nOur records indicate that your payment of {overdue_amount} was due on {overdue_date}, "
    "which means it has now been overdue for {days_overdue} days. We understand that oversights happen, "
    "and we kindly request that you settle this balance as soon as possible to avoid any service interruptions.\n\n"
    "We would appreciate your prompt attention to this matter. If you have any questions or concerns about this overdue payment, "
    "please do not hesitate to reach out.\n\nThank you for your cooperation.\n\nSincerely,\nYour Company",

    "Hello {customer_name},\n\nThis is a friendly reminder that your payment of {overdue_amount} was due on {overdue_date}, "
    "and it has now been overdue for {days_overdue} days. Timely payments help us continue to provide you with the best service possible. "
    "We kindly ask that you clear the outstanding amount at your earliest convenience.\n\n"
    "Please contact us if there are any issues preventing you from making this payment. We are here to assist you.\n\n"
    "Best regards,\nYour Company",

    "Hi {customer_name},\n\nWe wanted to bring to your attention that your account shows an overdue balance of {overdue_amount} since {overdue_date}. "
    "As of today, this amount has been overdue for {days_overdue} days. We request you to make the payment promptly to avoid further reminders.\n\n"
    "We value you as a customer and would like to resolve this matter as soon as possible. If you need to discuss this further, please get in touch.\n\n"
    "Thank you for your understanding.\n\nKind regards,\nYour Company",

    "Dear {customer_name},\n\nOur records show that your payment of {overdue_amount} was due on {overdue_date}, "
    "and it has now been outstanding for {days_overdue} days. To maintain a smooth business relationship, we kindly ask you to settle this overdue amount at your earliest convenience.\n\n"
    "If you have already made the payment, please disregard this message. Otherwise, we appreciate your prompt attention to this matter.\n\n"
    "Sincerely,\nYour Company"
]


def format_overdue_date(date):
    # Format date to 'DayOfWeek 7th Month Year'
    formatted_date = date.strftime('%a %d %B %Y')  # 'Mon 07 July 2024'
    # Add the appropriate suffix to the day
    day = int(date.strftime('%d'))
    suffix = 'th' if 11 <= day <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(day % 10, 'th')
    formatted_date = date.strftime(f'%a {day}{suffix} %B %Y')
    return formatted_date


def find_attachments(attachments_folder, customer_name):
    print(f"Looking for attachments for customer: {customer_name}")
    print("-" * 50)  # Separator line for clarity
    attachments = []
    for file_name in os.listdir(attachments_folder):
        # Check if the file name contains the customer's name
        if customer_name in file_name:
            attachment_path = os.path.join(attachments_folder, file_name)
            attachments.append(attachment_path)
            print(f"Found attachment: {attachment_path}")
    if not attachments:
        print(f"No attachments found for {customer_name}")
    return attachments


def send_email_via_outlook(to_email, subject, message, attachments):
    try:
        print(f"\nAttempting to create an Outlook email for {to_email}")
        # Create an Outlook application instance
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0: olMailItem

        # Set email parameters
        mail.To = to_email
        mail.Subject = subject
        mail.Body = message
        print(f"Email created for {to_email} with subject: {subject}")

        # Attach files if there are any
        for attachment in attachments:
            print(f"Attaching file: {attachment}")
            mail.Attachments.Add(attachment)

        # Send the email
        mail.Send()
        print(f"Email successfully sent to {to_email}")
    except Exception as e:
        print(f"Failed to send email to {to_email}: {e}")


def process_overdue_customers(spreadsheet_path, attachments_folder):
    print(f"\nReading spreadsheet from {spreadsheet_path}\n")
    # Read the spreadsheet
    try:
        data = pd.read_excel(spreadsheet_path)
        print("Spreadsheet successfully read.")
        print(f"Data read: {data.head()}\n")  # Print the first few rows to verify the content
    except Exception as e:
        print(f"Failed to read spreadsheet: {e}\n")
        return

    if data.empty:
        print("The spreadsheet is empty.\n")
        return

    for index, row in data.iterrows():
        print("\n" + "=" * 50)  # Separator line for clarity
        # Print the row data to see what is being processed
        print(f"Processing Row {index + 1}: {row}\n")

        # Verify column access, update keys if needed
        try:
            overdue_date = row['A']
            overdue_amount = row['B']
            customer_name = row['C']
            customer_email = row['D']
        except KeyError as e:
            print(f"KeyError: {e}. Check if the column names 'A', 'B', 'C', 'D' exist in the spreadsheet.\n")
            continue

        # Format the overdue date and calculate days overdue
        overdue_date = pd.to_datetime(overdue_date)
        formatted_overdue_date = format_overdue_date(overdue_date)
        days_overdue = (datetime.now() - overdue_date).days

        print(f"Customer: {customer_name}")
        print(f"Overdue Amount: {overdue_amount}")
        print(f"Overdue Date: {formatted_overdue_date}")
        print(f"Days Overdue: {days_overdue}\n")

        # Choose a random email template
        email_body = random.choice(EMAIL_TEMPLATES).format(
            customer_name=customer_name,
            overdue_amount=overdue_amount,
            overdue_date=formatted_overdue_date,
            days_overdue=days_overdue
        )
        print(f"Email body created for {customer_name}:\n{email_body}\n")

        # Find all relevant attachments for the customer
        attachments = find_attachments(attachments_folder, customer_name)

        # Send the email with the attachments (if any)
        send_email_via_outlook(customer_email, "Overdue Payment Notification", email_body, attachments)


if __name__ == "__main__":
    # Path to the spreadsheet and attachments folder
    spreadsheet_path = 'overdue_customers.xlsx'  # Replace with the path to your spreadsheet
    attachments_folder = 'attachments'  # Replace with the folder where overdue statements are stored

    # Ensure the folder exists
    if not os.path.isdir(attachments_folder):
        print(f"Attachments folder '{attachments_folder}' does not exist.\n")
    else:
        print(f"Attachments folder '{attachments_folder}' found.\n")
        process_overdue_customers(spreadsheet_path, attachments_folder)
