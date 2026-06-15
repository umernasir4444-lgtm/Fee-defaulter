# Fee Defaulter Report Generator

This local tool lets you upload a fee collection Excel file and automatically creates:

- `fee_defaulter_report.xlsx`
- `reminder_letters.html`
- `reminder_letters.txt`
- `defaulters.csv`
- `fee_defaulter_outputs.zip`

## How to Run

1. Double-click `launch_windows.bat`.
2. Open `http://127.0.0.1:8765` in your browser if it does not open automatically.
3. Upload your Excel file.
4. Download the generated report and reminder letters.

## Expected Excel Columns

The tool automatically detects common column names, including:

- Student: `Student`, `Student Name`, `Child Name`
- Parent: `Parent Name`, `Father Name`, `Guardian Name`
- Class: `Class`, `Grade`, `Section`
- Fee values: `Total Fee`, `Amount Due`, `Paid`, `Amount Paid`, `Pending`, `Balance`, `Outstanding`
- Optional contact fields: `Phone`, `Mobile`, `Email`
- Optional due field: `Due Date`, `Last Date`

If your sheet already has `Pending`, `Balance`, or `Outstanding`, that value is used. Otherwise the pending amount is calculated as `Total Fee - Paid`.
