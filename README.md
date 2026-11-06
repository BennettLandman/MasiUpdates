# MasiUpdates

Automated tool to download project status updates from Google Sheets / Google Forms and generate cleanly formatted Word (`.docx`) reports.

## Features

- **Direct Google Sheet Export**: Downloads the latest responses from the first sheet tab and saves a local Excel copy (`Form_Responses_1.xlsx`).
- **Timestamp Tracking**: Records the exact timestamp (Column A) of the latest submitted entry into `lastupdated.txt`.
- **Name Extraction**: Cleans email addresses in Column B by stripping `@<domain>` to isolate team member usernames.
- **Bullet Formatting**: Normalizes manual/soft line breaks in Column C and formats each update as a styled, hanging bullet point.
- **Alphabetical Sorting**: Automatically sorts all entries by team member name in Column 1.
- **Word Document Generation**: Produces a Microsoft Word report (`YYYY.MM.DD.docx`) with:
  - Document & page headers dated with the current date.
  - Table column headers matching the Google Form questions.
  - Alternating row background colors for easy reading.
- **Zero Dependencies**: Pure Python 3 standard library implementation—works out of the box without requiring `pip install`.

## Usage

### Run Default Workflow
```bash
python3 generate_report.py
```
This will:
1. Fetch latest entries from the Google Sheet.
2. Save raw data to `Form_Responses_1.xlsx`.
3. Save the latest submission timestamp to `lastupdated.txt`.
4. Generate the formatted Word document `YYYY.MM.DD.docx` (e.g., `2026.11.06.docx`).

### Custom Options
```bash
# Specify custom Google Sheet URL or ID
python3 generate_report.py --url "https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit?usp=sharing"

# Generate report from a local Excel file already downloaded
python3 generate_report.py --input "Form_Responses_1.xlsx" --output "Status_Report.docx"

# Specify a custom timestamp file path
python3 generate_report.py --last-updated-file "latest_timestamp.txt"
```
