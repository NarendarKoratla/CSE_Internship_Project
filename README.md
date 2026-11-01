# Email-Enabled Attendance Tracker

A Python desktop application for recording student absences in Excel workbooks and sending email notifications when students reach configured absence thresholds. Built as a CSE internship project.

## Features

- Tkinter desktop interface for selecting academic year, section, subject, and student roll numbers.
- Reads and updates attendance workbooks using `openpyxl`.
- Sends student notifications through Gmail SMTP.
- Sends a staff notification when a student exceeds the configured threshold.
- Loads email credentials and the Excel data directory from environment variables using `python-dotenv`.
- Includes a sample-data generator for testing without real student records.

## Technology

- Python 3
- Tkinter (included with many standard Python installations)
- openpyxl
- python-dotenv
- smtplib (Python standard library)

## Requirements

- Python 3
- Gmail account configured for SMTP with a Google App Password
- Excel workbooks matching the expected column layout

## Run locally

### 1. Clone the repository

```bash
git clone https://github.com/NarendarKoratla/CSE_Internship_Project.git
cd CSE_Internship_Project
```

### 2. Install Python dependencies

```bash
python -m pip install openpyxl python-dotenv
```

### 3. Configure environment variables

Copy `.env.example` to `.env`, then update the values:

```dotenv
SENDER_EMAIL=your_email@gmail.com
SENDER_PASSWORD=your_google_app_password
EXCEL_FILE_DIR=./sample_data
```

Use a Google App Password where required; do not use your regular Google account password. Never commit `.env`, credentials, or real student data. Confirm that `.env` is ignored by Git before adding secrets.

### 4. Generate sample workbooks

```bash
python generate_sample_data.py
```

Set `EXCEL_FILE_DIR` to the generated `sample_data` directory (use an absolute path if your environment does not resolve the relative path as expected).

### 5. Start the application

```bash
python app.py
```

## Expected Excel layout

The app expects each workbook to contain columns in this order:

| Column | Value |
|---|---|
| A | Roll number |
| B | Student name |
| C | Student email |
| D onward | Absence counts for the configured subjects |

Subject-column ordering must match the `SUBJECTS` configuration in `app.py`. Use dummy records when testing.

## Notification logic

The current code sends progressively stronger messages at the first, second, and third absence, and sends an exceeded-limit message after the third absence. It also attempts to notify the configured staff email for that year and subject.

## Notes and limitations

- The application updates Excel files locally; keep backups before testing with important data.
- The number-of-absentees field is collected by the UI, while the current processing logic uses the supplied roll-number list. Ensure the list matches the actual absentees.
- Staff email addresses in the source are placeholders and must be configured before use.
- Email delivery requires valid SMTP credentials and network access.
- This is an educational project. Test with sample data before using any real student information.

## Author

**Narendar Koratla** · [GitHub](https://github.com/NarendarKoratla)
