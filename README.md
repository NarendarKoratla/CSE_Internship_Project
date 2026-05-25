# Email-Enabled Attendance Tracker

A desktop application built with Python and Tkinter that tracks student attendance using Excel files and automatically sends email notifications to students who exceed their leave limit.

> Built as a **2nd Year CSE Internship Project**.

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
CSE_Internship_Project/
├── app.py                    # Main application
├── generate_sample_data.py   # Script to generate sample Excel files for testing
├── .env.example              # Template for environment variables
├── .gitignore                # Git ignore rules
├── README.md                 # This file
├── sample_data/              # Generated sample Excel files (after running the script)
│   ├── 1st_CSE-A.xlsx
│   ├── 1st_CSE-B.xlsx
│   └── ...
├── .env                      # Your credentials (not tracked by git)
└── *.xlsx                    # Real student data files (not tracked by git)
```

## Getting Started

Follow these steps to set up and run the project on your machine:

### Step 1: Clone the Repository

```bash
git clone https://github.com/NarendarKoratla/CSE_Internship_Project.git
cd CSE_Internship_Project
```

### Step 2: Install Dependencies

Make sure you have Python 3 installed, then run:

```bash
pip install openpyxl python-dotenv
```

### Step 3: Set Up Email Credentials

1. Copy the example environment file:
   ```bash
   # On Windows
   copy .env.example .env

   # On Mac/Linux
   cp .env.example .env
   ```

2. Open the `.env` file and fill in your details:
   ```
   SENDER_EMAIL=your_email@gmail.com
   SENDER_PASSWORD=your_app_password_here
   EXCEL_FILE_DIR=C:\path\to\your\excel\files
   ```

> **Important:** You must use a [Google App Password](https://myaccount.google.com/apppasswords), not your regular Gmail password. To generate one:
> 1. Go to [Google App Passwords](https://myaccount.google.com/apppasswords)
> 2. Select **Mail** as the app
> 3. Generate and copy the 16-character password
> 4. Paste it as `SENDER_PASSWORD` in your `.env` file

### Step 4: Generate Sample Data (For Testing)

Run the sample data generator to create Excel files with dummy student data:

```bash
python generate_sample_data.py
```

This creates 12 Excel files inside the `sample_data/` folder (one per year-section combination). Update `EXCEL_FILE_DIR` in your `.env` to point to this folder:

```
EXCEL_FILE_DIR=C:\path\to\CSE_Internship_Project\sample_data
```

### Step 5: Run the Application

```bash
python app.py
```

## How It Works

1. **Page 1** — Select the **Academic Year** (1st–4th) and **Section** (CSE-A/B/C), then click **Next**.
2. **Page 2** — Select the **Subject**, enter the **Number of Absentees**, and type their **Roll Numbers** separated by spaces.
3. Click **Submit Attendance** — the app will:
   - Update the absence count in the corresponding Excel file
   - Send automated email alerts to students based on their absence level
   - Notify staff if any student exceeds the leave limit

## Excel File Format

Each Excel file must follow this exact column layout:

| Column A | Column B | Column C | Column D | Column E | Column F |
|----------|----------|----------|----------|----------|----------|
| Roll No  | Name     | Email    | Subject 1 absences | Subject 2 absences | Subject 3 absences |

**Example for `1st_CSE-A.xlsx`:**

| Roll No | Name         | Email              | M1 | EC | PPS |
|---------|--------------|--------------------|----|----|----|
| 1       | Rahul Sharma | rahul@gmail.com    | 0  | 0  | 0  |
| 2       | Priya Reddy  | priya@gmail.com    | 0  | 0  | 0  |

> The subjects in columns D, E, F must match the order defined in `SUBJECTS` inside `app.py`.

## License

This project is open-source and available under the [MIT License](LICENSE).
