# Tracker Spreadsheet Organizer

The Tracker Spreadsheet Organizer is a user-friendly tool designed to help organize and clean email campaign engagement data. This application merges multiple CSV files (for opens and clicks) into a consolidated spreadsheet while automatically filtering out unengaged contacts for better efficiency.

---

## Features

- **File Selection**: Upload CSV files for opens and clicks across Email 1, Email 2, and Email 3.
- **Engagement Tracking**:
  - Converts engagement data (opens and clicks) to "Y" based on thresholds:
    - **Opens**: `Y` if the contact opened 2 or more times.
    - **Clicks**: `Y` if the contact clicked 1 or more times.
- **Filtering**: Automatically removes contacts with no engagement (no opens or clicks).
- **Key Section**: Includes a color-coded key for reference.
- **Final Output**: Generates a clean Excel spreadsheet, formatted and ready for further processing.

---

## How to Use

### Launch the Application
1. Run the executable file (**TrackerSpreadsheetOrganizer.exe**) provided with this package.

### Enter Campaign Code
2. Enter the campaign code for the email campaign you are working on.

### Select Files
3. Click the buttons to upload CSV files for:
   - **Email 1 Opens and Clicks**
   - **Email 2 Opens and Clicks**
   - **Email 3 Opens and Clicks**

### Process Files
4. Once all files are selected, click the **Process & Save Tracker Spreadsheet** button.

### Save Output
5. Choose where to save the final Excel file, and you’re done!

---

## Output Structure

The generated Excel file contains:

### Key Section
- A color-coded reference for Jacqui’s use:
  - **Yellow**: Sent follow-up
  - **Red**: No response
  - **Green**: Response

### Consolidated Data
- Engagement metrics (opens and clicks converted to "Y").
- Cleaned and filtered contact information.
- Pre-filled columns for campaign details:
  - **Contact Status**, **Topic**, **Status Reason**, **Campaign Code**, and **Lead Source**.

---

## Requirements

- This application runs as a standalone executable. No installation or additional software is required.
- Compatible with Windows.
- Ensure the CSV files are properly formatted with headers like `email address`, `opens`, `clicks`, etc.

---

## Support

If you encounter any issues, please contact [rossjohnstone@lifelink.org.uk](mailto:rossjohnstone@lifelink.org.uk).
