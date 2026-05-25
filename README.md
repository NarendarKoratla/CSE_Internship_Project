# Email-Enabled Attendance Tracker

A desktop application built with Python and Tkinter that tracks student attendance using Excel files and automatically sends email notifications to students who exceed their leave limit.

## Features

- **GUI Interface**: Clean Tkinter-based two-page interface for selecting year, section, subject, and entering absentee roll numbers.
- **Excel Integration**: Reads and updates student attendance records stored in `.xlsx` files using `openpyxl`.
- **Automated Email Alerts**: Sends progressive warning emails to students based on their absence count:
  - 1st leave — Caution notice
  - 2nd leave — Warning notice
  - 3rd leave — Final alert
  - 4+ leaves — Exceeded limit notice + staff notification
- **Staff Notification**: Automatically notifies the respective faculty when a student crosses the allowed leave threshold.

## Tech Stack

- **Python 3**
- **Tkinter** — GUI framework
- **openpyxl** — Excel file read/write
- **smtplib** — Email delivery via Gmail SMTP
- **python-dotenv** — Secure credential management

## Project Structure

```
Email-Enabled Attendance Tracker/
├── app.py              # Main application
├── .env                # Email credentials (not tracked by git)
├── .gitignore          # Git ignore rules
├── README.md           # This file
└── *.xlsx              # Student data files (not tracked by git)
```

## Setup

1. **Install dependencies:**
   ```bash
   pip install openpyxl python-dotenv
   ```

2. **Create a `.env` file** in the project root:
   ```
   SENDER_EMAIL=your_email@gmail.com
   SENDER_PASSWORD=your_app_password
   ```
   > **Note:** You must use a [Google App Password](https://myaccount.google.com/apppasswords), not your regular Gmail password.

3. **Prepare Excel files** named in the format `{Year}_{Section}.xlsx` (e.g., `1st_CSE-A.xlsx`) with the following column layout:

   | Column A | Column B | Column C | Column D+ |
   |----------|----------|----------|-----------|
   | Roll No  | Name     | Email    | Subject absence counts |

4. **Run the application:**
   ```bash
   python app.py
   ```

## How It Works

1. Select the **Academic Year** and **Section** on the first page.
2. Select the **Subject**, enter the number of absentees, and type their **Roll Numbers** (space-separated) on the second page.
3. The app updates the Excel file and sends appropriate email alerts automatically.

## License

This project is open-source and available under the [MIT License](LICENSE).
