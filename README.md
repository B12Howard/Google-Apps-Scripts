# Google-Apps-Scripts
Snippets of Google Apps Scripts I've used for projects. 

## Examples of a Google Apps Script for building a Query based on user input. 

### GmailWithSheetsCriteria.gs
Iterates through a Sheet's tabs, uses Gmail to send emails depending on values in a Google Sheet. 

### OldQueryBuilder.gs
It also deals with how to query two separate worksheets in Google Sheets of different column sizes
by creating essentially a temporary table of values from the two uneven sheets to 
create outputs with the same number of columns so Query can read the data without erroring.

This script was used until we came up with a better way to organize our data in Sheets.

### Messenger.gs
Example of how to pull information from Google Sheeets and send via Twilio or Gmail
