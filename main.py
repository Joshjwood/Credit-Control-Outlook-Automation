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
    attachments = []
    for file_name in os.listdir(attachments_folder):
        # Check if the file name contains the customer's name
        if customer_name in file_name:
            attachment_path = os.path.join(attachments_folder, file_name)
            attachments.append(attachment_path)
    return attachments


def send_email_via_outlook(to_email, subject, message, attachments, from_email=None):
    try:
        print(f"\nAttempting to create an Outlook email for {to_email}")
        # Create an Outlook application instance
        outlook = win32com.client.Dispatch('Outlook.Application')
        mail = outlook.CreateItem(0)  # 0: olMailItem

        # Set email parameters
        mail.To = to_email
        mail.Subject = subject
        mail.Body = message

        # Set the 'From' email address if specified
        if from_email:
            mail.SentOnBehalfOfName = from_email

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
        print("Spreadsheet successfully read.\n")
    except Exception as e:
        print(f"Failed to read spreadsheet: {e}\n")
        return

    if data.empty:
        print("The spreadsheet is empty.\n")
        return

    # List to store customer email details for sending later
    email_summaries = []

    # Assessment Phase
    for index, row in data.iterrows():
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

        # Find all relevant attachments for the customer
        attachments = find_attachments(attachments_folder, customer_name)
        num_attachments = len(attachments)

        # Create summary message for this customer
        summary_message = f"{customer_name} will be sent an email with {num_attachments} attachment(s)"
        print(summary_message)

        # Store details for sending emails later
        email_summaries.append({
            'to_email': customer_email,
            'subject': "Overdue Payment Notification",
            'message': random.choice(EMAIL_TEMPLATES).format(
                customer_name=customer_name,
                overdue_amount=overdue_amount,
                overdue_date=formatted_overdue_date,
                days_overdue=days_overdue
            ),
            'attachments': attachments,
            'summary': summary_message
        })

    # Confirmation
    print("\nIs this correct? Press Y to continue (send emails), or any other key to abort:")
    user_input = input().strip().upper()

    if user_input == 'Y':
        # Send Emails Phase
        for email_detail in email_summaries:
            print("\n" + "=" * 50)  # Separator line for clarity
            print(f"Sending to: {email_detail['summary']}\n")
            send_email_via_outlook(
                email_detail['to_email'],
                email_detail['subject'],
                email_detail['message'],
                email_detail['attachments']
            )
    else:
        print("\nAborted. No emails were sent.")


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
