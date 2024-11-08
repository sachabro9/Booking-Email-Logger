# Booking-Email-Logger
# Booking Email Automation Script

This Google Apps Script automates the processing of booking confirmation emails received from various sources (e.g., Airbnb, WeTravel, and Bókun). It extracts essential booking information from each email, logs it into a Google Sheet, and creates an event in Google Calendar with the details.

## Features

- **Automated Email Processing**: Scans Gmail inbox for new booking confirmation emails labeled with a specified label.
- **Data Extraction**: Extracts key information including booking source, booking date, price, guest name, and contact information.
- **Google Sheets Logging**: Logs booking information in a Google Sheet for easy access and record-keeping.
- **Google Calendar Integration**: Creates a calendar event for each booking, detailing the booking source, guest name, contact information, and price.

## Getting Started

### Prerequisites

- A Google account with access to Gmail, Google Sheets, and Google Calendar.
- Basic understanding of Google Apps Script.

### Setup Instructions

1. **Copy the Script**: Open Google Apps Script and create a new project. Copy the provided code into your Apps Script editor.

2. **Configuration**:
   - Replace `"YOUR_LABEL_NAME"` with the exact name of the Gmail label that contains the booking confirmation emails.
   - Replace `"YOUR_SHEET_NAME"` with the name of the Google Sheet where booking data will be logged.
   - Replace `"YOUR_CALENDAR_ID"` with your Google Calendar ID (found in your calendar settings under "Integrate calendar").

3. **OAuth Scopes**: Ensure the script requests appropriate permissions. These should be:
   - `https://www.googleapis.com/auth/gmail.readonly`
   - `https://www.googleapis.com/auth/spreadsheets`
   - `https://www.googleapis.com/auth/gmail.modify`
   - `https://www.googleapis.com/auth/calendar`

4. **Testing**: Run the `logBookingEmails` function manually to ensure it works as expected. Verify that:
   - Booking data appears in the specified Google Sheet.
   - Events are created in Google Calendar with the correct details.

### Automation

To automate this script, set up a time-driven trigger in Apps Script:
1. Go to **Triggers** in your Apps Script editor.
2. Add a new trigger for the `logBookingEmails` function.
3. Set the trigger to run at your preferred interval (e.g., daily or hourly).

## License

This project is open-source and available under the MIT License.

## Acknowledgments

Special thanks to Fernando Sabater, one of my most important mentors.
