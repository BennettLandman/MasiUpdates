# MasiUpdates

Automated tool to download project status updates from Google Sheets / Google Forms and generate cleanly formatted Word (`.docx`) reports.

## Features

- **Direct Google Sheet Export**: Downloads the latest responses from the first sheet tab and saves a local Excel copy (`Form_Responses_1.xlsx`).
- **Timestamp Tracking & Incremental Filtering**: Records the timestamp (Column A) of the latest submitted entry into `lastupdated.txt`. If `lastupdated.txt` exists, subsequent runs will automatically only include new entries submitted after that recorded time.
- **Overwrite Protection**: Never overwrites existing reports. If `YYYY.MM.DD.docx` exists, it automatically increments version suffixes (`YYYY.MM.DD_v2.docx`, `YYYY.MM.DD_v3.docx`, etc.).
- **Name Extraction**: Cleans email addresses in Column B by stripping `@<domain>` to isolate team member usernames.
- **Bullet Formatting**: Normalizes manual/soft line breaks in Column C and formats each update as a styled, hanging bullet point.
- **Alphabetical Sorting**: Automatically sorts all entries by team member name in Column 1.
- **Word Document Generation**: Produces a Microsoft Word report (`YYYY.MM.DD.docx`) matching the reference template with:
  - Document & page headers dated with the current date.
  - Dynamic page numbering in the lower-right footer.
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
2. Check `lastupdated.txt` and filter for any entries submitted after the previous run.
3. Save raw data to `Form_Responses_1.xlsx`.
4. Generate the formatted Word document `YYYY.MM.DD.docx` (or `_v2`, `_v3` if the file already exists).
5. Update `lastupdated.txt` with the timestamp of the latest entry processed.

### Custom Options
```bash
# Specify custom Google Sheet URL or ID
python3 generate_report.py --url "https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit?usp=sharing"

# Force inclusion of all entries regardless of lastupdated.txt
python3 generate_report.py --ignore-last-updated

# Generate report from a local Excel file already downloaded
python3 generate_report.py --input "Form_Responses_1.xlsx" --output "Status_Report.docx"

# Specify a custom timestamp file path
python3 generate_report.py --last-updated-file "latest_timestamp.txt"
```

