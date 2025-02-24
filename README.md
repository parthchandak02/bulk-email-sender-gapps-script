# Email Management and Response Tracking System

A Google Apps Script solution for managing email campaigns and tracking responses automatically. This system helps you send personalized emails and monitor responses through a Google Spreadsheet interface.

## Features

- Automated email sending with personalized templates
- Response tracking and monitoring
- Email link storage and organization
- Progress tracking through a user-friendly interface
- Batch processing capabilities
- Emoji-enhanced UI for better readability
- Rate limiting and error handling

## Prerequisites

1. Google Account with access to:
   - Google Sheets
   - Google Apps Script
   - Gmail
2. Browser with JavaScript enabled
3. Necessary Google Apps Script permissions (will be requested during setup)

## Detailed Setup Instructions

### 1. Google Spreadsheet Setup
1. Create a new Google Spreadsheet at [sheets.google.com](https://sheets.google.com)
2. Import `sample_data.csv` into your spreadsheet:
   - File > Import > Upload > Select `sample_data.csv`
   - Select "Replace current sheet" in import options
   - Click "Import data"

### 2. Google Apps Script Setup
1. From your spreadsheet, go to Extensions > Apps Script
2. In the Apps Script editor:
   - Delete any existing code in `Code.gs`
   - Copy the contents of `email.gs` into the editor
   - File > Save (or Ctrl/Cmd + S)

### 3. Configuration
1. In the Apps Script editor, locate the `CONFIG` object at the top of the code:
   ```javascript
   const CONFIG = {
     SPREADSHEET_ID: 'YOUR_SPREADSHEET_ID_HERE',
     SHEET_NAME: 'Sheet1',
     EMAIL_SUBJECT: 'Follow-up Request',
     SENDER_NAME: 'Your Name'
   };
   ```
2. Update the configuration:
   - `SPREADSHEET_ID`: Copy from your spreadsheet's URL (the long string between /d/ and /edit)
   - `SHEET_NAME`: Your sheet name (default is 'Sheet1')
   - `EMAIL_SUBJECT`: Your default email subject
   - `SENDER_NAME`: Your name as it should appear in emails

### 4. Email Template Customization
1. Find the `getEmailBody` function in the script
2. Modify the HTML template to match your needs:
   - Update the greeting
   - Customize the main message
   - Add your own sections
   - Modify styling if needed

### 5. Authorization Setup
1. Click the "Run" button (▶️) in the Apps Script editor
2. Select the `setup` function
3. Click "Review permissions" in the authorization popup
4. Choose your Google Account
5. Click "Advanced" > "Go to [Your Project Name] (unsafe)"
6. Click "Allow"

### 6. Testing the Setup
1. Return to your spreadsheet
2. Refresh the page
3. Look for the new "📧 Email Manager" menu
4. Test the setup:
   - Click "📧 Email Manager" > "Send Emails..."
   - Verify the dialog appears
   - Check if the test row is visible

## Usage Instructions

### Sending Emails
1. Fill in the spreadsheet with recipient information:
   - First Name
   - Email Address
   - Set Action to "Send Email"
2. Click "📧 Email Manager" > "Send Emails..."
3. Choose to process single or all rows
4. Monitor the progress in the dialog

### Checking Responses
1. Click "📧 Email Manager" > "Check Responses..."
2. The system will automatically:
   - Scan for new responses
   - Update the spreadsheet
   - Show progress in real-time

### Column Structure
- Column A: First Name
- Column B: Email Address
- Column C: Action Status
- Column D: Sent Date
- Column E: Email Link
- Column F onwards: Response Data

### Action Statuses
- "Send Email": Ready to send
- "Awaiting Response": Email sent, waiting for reply
- [Empty]: No action needed

## Troubleshooting

### Common Issues
1. "Authorization Required"
   - Run the `setup` function again
   - Check Google Apps Script permissions

2. "Spreadsheet ID Error"
   - Verify CONFIG.SPREADSHEET_ID matches your spreadsheet's URL
   - Ensure you have edit access to the spreadsheet

3. "Gmail Quota Exceeded"
   - Wait for quota to reset (usually 24 hours)
   - Reduce batch size

4. "Script Timeout"
   - Process fewer emails at once
   - Check for slow network connections

### Debug Mode
To enable detailed logging:
1. In Apps Script editor, View > Execution Log
2. Run your function
3. Check logs for errors

## Rate Limits & Quotas

- Gmail sending limit: 100 emails/day for consumer accounts
- Script runtime: 6 minutes maximum
- Properties storage: 500KB per script
- Triggers: 20 per script

## Best Practices

1. Regular Backups
   - Export spreadsheet regularly
   - Keep copy of script code

2. Testing
   - Test with small batches first
   - Use test email addresses initially
   - Verify email template rendering

3. Maintenance
   - Check logs periodically
   - Clear old response data
   - Update authorization if needed

## Contributing

Feel free to submit issues and enhancement requests!

## Security Notes

- Never share your `CONFIG.SPREADSHEET_ID`
- Keep your email template professional
- Regularly review authorized access
- Monitor script execution logs
