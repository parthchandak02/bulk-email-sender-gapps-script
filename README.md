# Email Management and Response Tracking System

A Google Apps Script solution for managing email campaigns and tracking responses automatically. This system helps you send personalized emails and monitor responses through a Google Spreadsheet interface.

## Features

- Automated email sending with personalized templates
- Response tracking and monitoring
- Email link storage and organization
- Progress tracking through a user-friendly interface
- Batch processing capabilities

## Prerequisites

1. Google Account with access to:
   - Google Sheets
   - Google Apps Script
   - Gmail

## Setup Instructions

1. Create a new Google Spreadsheet
2. Import the `sample_data.csv` file into your spreadsheet
3. Open Script Editor (Extensions > Apps Script)
4. Copy the contents of `email.gs` into the script editor
5. Update the `CONFIG` object with your:
   - Spreadsheet ID (from the URL of your spreadsheet)
   - Sheet name (if different from 'Sheet1')
   - Email subject
   - Sender name
6. Customize the email template in the `getEmailBody` function
7. Save and run the `setup` function to grant necessary permissions

## Usage

1. The spreadsheet should have the following columns:
   - First Name
   - Email
   - Action
   - Sent Date
   - Email Link
   - Response columns (will be populated automatically)

2. Set the "Action" column to "Send Email" for rows you want to process

3. Use the custom menu "Email Manager" to:
   - Send emails
   - Check for responses
   - Monitor progress

## Spreadsheet Structure

The system expects the following column structure:
- Column A: First Name
- Column B: Email Address
- Column C: Action Status
- Column D: Sent Date
- Column E: Email Link
- Column F onwards: Response Data

## Action Statuses

- "Send Email": Row is ready to be processed
- "Awaiting Response": Email has been sent, waiting for reply
- [Empty]: No action needed

## Notes

- The script includes rate limiting and error handling
- Responses are automatically formatted with emojis for better readability
- The system tracks email threads to avoid duplicate responses

## Troubleshooting

If you encounter issues:
1. Check the script execution logs in Apps Script
2. Verify all permissions are granted
3. Ensure the spreadsheet structure matches the expected format
4. Check your Gmail quota and limits

## Contributing

Feel free to submit issues and enhancement requests!
