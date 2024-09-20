Google Sheets Email Automation Project
<br>
Author: Shiv Kumar
<br>
Overview
<br>
This project is designed to automate the process of sending emails directly from a Google Spreadsheet using Google Apps Script. The script allows for dynamic email content, including personalized data for each recipient and an HTML table format that displays important information.
<br>
The project covers two main functions:
<br>
Main Function: Sends personalized emails to recipients with data fetched from Sheet1.
Additional Data Emails Function: Sends a separate email to a different set of recipients using data from Sheet2, presented in an HTML table format.
<br>
Features
<br>
Email Personalization: The script dynamically generates email content based on the data in the spreadsheet, ensuring each recipient receives information tailored to them.
HTML Email Body: Both functions send emails with an HTML table, providing a clear and structured display of the recipient’s data.
Error Handling: The script includes logging for cases where an email address is missing or invalid, allowing for easy troubleshooting.
Dual Data Source: The script handles data from two different sheets (Sheet1 and Sheet2), allowing flexibility in email content management.
Sheets Structure
Sheet1
Contains the primary data, which includes:

Column B: Name
Column D: Marks
Column E: Email Address
Sheet2
Contains the additional data, which includes:

Column A: Name
Column B: Marks
Column C: Email Address (for sending emails)
Script Functions
Main Function (main)

This function retrieves data from Sheet1 and sends emails to the corresponding recipients.
The email body includes the recipient's name and marks, formatted in a user-friendly way.
The function also incorporates additional data fetched from Sheet2, included in the email body.
Additional Data Function (sendAdditionalDataEmails)

This function sends emails based on data retrieved from Sheet2.
It generates an HTML table containing the names and marks for each recipient and sends it to the email address provided in Sheet2.
How to Use
Setup Google Sheets: Create a Google Spreadsheet with two sheets:

Sheet1 for the main email content.
Sheet2 for the additional email content.
Script Setup:

Open the Google Apps Script editor from the spreadsheet (Extensions > Apps Script).
Copy the provided script into the editor.
Authorize: The script requires authorization to access your Gmail and Google Sheets. When running the script for the first time, click “Authorize” when prompted.

Run the Script:

Manually run the main function to trigger the email sending process.
Both sets of emails (from Sheet1 and Sheet2) will be sent in separate processes.
Error Handling
No Email Address: If an email address is missing in either sheet, the script will skip that entry and log the error in the Apps Script logs.
Email Sending Error: Any errors encountered while sending an email will be logged, along with the specific error message for troubleshooting.
Future Improvements
Option to add attachments to emails.
More advanced error handling, such as retry mechanisms.
Scheduling emails at specific times using Google Apps Script triggers.
Author
Shiv Kumar
