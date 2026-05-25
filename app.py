# app.py
# Email-Enabled Attendance Tracker (Tkinter GUI + Excel + Email)


import tkinter as tk
from tkinter import ttk, messagebox
import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from dotenv import load_dotenv

# Load environment variables from a .env file (if present)
load_dotenv()

# ---------------------------------------------
# CONFIGURATION SECTION
# ---------------------------------------------
# Default Excel directory: <Your Documents>/Email-Enabled Attendance Tracker
# You can override this with environment variable EXCEL_FILE_DIR
EXCEL_FILE_DIR = os.getenv(
    "EXCEL_FILE_DIR",
    os.path.join(os.path.expanduser("~"), "Documents", "Email-Enabled Attendance Tracker")
)

# Email credentials (set in .env or system env variables)
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
SENDER_PASSWORD = os.getenv("SENDER_PASSWORD")

# Subjects by year
SUBJECTS = {
    "1st Year": ["M1", "EC", "PPS"],
    "2nd Year": ["P&S", "OOPS", "DS"],
    "3rd Year": ["DAA", "PPL", "CN"],
    "4th Year": ["CC", "LP", "MC"]
}

# Staff emails (placeholders only — replace with real ones before running)
STAFF_EMAILS = {
    "1st Year": {"M1": "m1@gmail.com", "EC": "ec@gmail.com", "PPS": "pps@gmail.com"},
    "2nd Year": {"P&S": "ps@gmail.com", "OOPS": "oops@gmail.com", "DS": "ds@gmail.com"},
    "3rd Year": {"DAA": "daa@gmail.com", "PPL": "ppl@gmail.com", "CN": "cn@gmail.com"},
    "4th Year": {"CC": "cc@gmail.com", "LP": "lp@gmail.com", "MC": "mc@gmail.com"},
}

# ---------------------------------------------
# UTILITY FUNCTIONS
# ---------------------------------------------
def get_excel_file_path(academic_year, section):
    year_prefix = academic_year.split()[0]   # e.g. "1st"
    return os.path.join(EXCEL_FILE_DIR, f"{year_prefix}_{section}.xlsx")

def load_excel_file(academic_year, section):
    file_path = get_excel_file_path(academic_year, section)
    try:
        book = openpyxl.load_workbook(file_path)
        sheet = book.active
        return book, sheet
    except Exception as e:
        messagebox.showerror("Error", f"Error loading Excel file: {e}")
        return None, None

def save_excel_file(book, academic_year, section):
    file_path = get_excel_file_path(academic_year, section)
    try:
        book.save(file_path)
    except Exception as e:
        messagebox.showerror("Error", f"Error saving Excel file: {e}")

def send_email(to_email, subject, body, student_name=None, student_roll=None):
    try:
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = to_email
        msg['Subject'] = subject
        msg.attach(MIMEText(body, 'plain'))

        with smtplib.SMTP('smtp.gmail.com', 587) as server:
            server.starttls()
            server.login(SENDER_EMAIL, SENDER_PASSWORD)
            server.sendmail(SENDER_EMAIL, to_email, msg.as_string())

        if student_name and student_roll:
            return f"Email sent to {student_name} (Roll No: {student_roll}) at {to_email}\n"
        else:
            return f"Email sent to staff at {to_email}\n"
    except smtplib.SMTPAuthenticationError:
        return f"Authentication error for {to_email}. Check credentials.\n"
    except Exception as e:
        return f"Failed to send email to {to_email}: {e}\n"

def process_attendance():
    try:
        academic_year = academic_year_var.get()
        section = section_var.get()
        subject = subject_var.get()
        num_absentees = int(num_absentees_var.get())
        roll_numbers = list(map(int, roll_numbers_var.get().split()))

        book, sheet = load_excel_file(academic_year, section)
        if not book or not sheet:
            return

        subject_index = SUBJECTS[academic_year].index(subject) + 1

        email_logs = []
        exceeded_rolls = []
        exceeded_names = []

        for student_roll in roll_numbers:
            for i in range(2, sheet.max_row + 1):
                if sheet.cell(row=i, column=1).value == student_roll:
                    col_index = 3 + subject_index
                    current_days = sheet.cell(row=i, column=col_index).value or 0
                    current_days += 1
                    sheet.cell(row=i, column=col_index).value = current_days
                    save_excel_file(book, academic_year, section)

                    student_email = sheet.cell(row=i, column=3).value
                    student_name = sheet.cell(row=i, column=2).value

                    if current_days == 1:
                        msg = "You have taken leave for the first time. Please be cautious."
                        email_logs.append(send_email(student_email, f"First-Time Leave - {subject}",
                                                     f"Dear {student_name},\n\n{msg}\n\nRegards,\nAttendance Team",
                                                     student_name, student_roll))
                    elif current_days == 2:
                        msg = "Warning! You can take only one more leave."
                        email_logs.append(send_email(student_email, f"Warning - {subject}",
                                                     f"Dear {student_name},\n\n{msg}\n\nRegards,\nAttendance Team",
                                                     student_name, student_roll))
                    elif current_days == 3:
                        msg = "Alert! This is your last allowed leave."
                        email_logs.append(send_email(student_email, f"Final Leave Alert - {subject}",
                                                     f"Dear {student_name},\n\n{msg}\n\nRegards,\nAttendance Team",
                                                     student_name, student_roll))
                    elif current_days > 3:
                        msg = "You have exceeded your leave limit. Meet HoD."
                        email_logs.append(send_email(student_email, f"Exceeded Limit - {subject}",
                                                     f"Dear {student_name},\n\n{msg}\n\nRegards,\nAttendance Team",
                                                     student_name, student_roll))
                        exceeded_rolls.append(str(student_roll))
                        exceeded_names.append(student_name)

        if exceeded_rolls:
            staff_email = STAFF_EMAILS[academic_year][subject]
            staff_msg = f"The following students exceeded the leave limit in {subject}:\nRolls: {', '.join(exceeded_rolls)}\nNames: {', '.join(exceeded_names)}"
            email_logs.append(send_email(staff_email, f"Leave Exceeded - {subject}", staff_msg))

        messagebox.showinfo("Email Summary", "\n".join(email_logs))
        ask_for_another_entry()

    except ValueError:
        messagebox.showerror("Input Error", "Please enter valid roll numbers separated by spaces.")

def ask_for_another_entry():
    response = messagebox.askyesno("Continue", "Do you want to enter attendance for another subject?")
    if response:
        show_first_page()
    else:
        root.destroy()

def update_subjects():
    selected_year = academic_year_var.get()
    subject_dropdown['values'] = SUBJECTS[selected_year]
    subject_var.set(SUBJECTS[selected_year][0])

def show_first_page():
    for widget in root.winfo_children():
        widget.destroy()
    create_first_page()

def create_first_page():
    frame = tk.Frame(root, padx=30, pady=30)
    frame.pack(expand=True)

    tk.Label(frame, text="Attendance Tracker", font=("Arial", 32, "bold")).pack(pady=30)

    tk.Label(frame, text="Select Academic Year:", font=("Arial", 16)).pack(anchor="w", pady=10)
    academic_year_dropdown = ttk.Combobox(frame, textvariable=academic_year_var, values=list(SUBJECTS.keys()), font=("Arial", 14))
    academic_year_dropdown.pack(fill="x")
    academic_year_dropdown.current(0)

    tk.Label(frame, text="Select Section:", font=("Arial", 16)).pack(anchor="w", pady=10)
    section_dropdown = ttk.Combobox(frame, textvariable=section_var, values=["CSE-A", "CSE-B", "CSE-C"], font=("Arial", 14))
    section_dropdown.pack(fill="x")
    section_dropdown.current(0)

    tk.Button(frame, text="Next", command=show_second_page, bg="#28a745", fg="white", font=("Arial", 16), padx=10, pady=5).pack(pady=30)

def show_second_page():
    for widget in root.winfo_children():
        widget.destroy()

    frame = tk.Frame(root, padx=30, pady=30)
    frame.pack(expand=True)

    tk.Label(frame, text="Attendance Tracker", font=("Arial", 32, "bold")).pack(pady=30)

    tk.Label(frame, text="Select Subject:", font=("Arial", 16)).pack(anchor="w", pady=10)
    global subject_dropdown
    subject_dropdown = ttk.Combobox(frame, textvariable=subject_var, font=("Arial", 14))
    subject_dropdown.pack(fill="x")
    update_subjects()

    tk.Label(frame, text="Number of Absentees:", font=("Arial", 16)).pack(anchor="w", pady=10)
    tk.Entry(frame, textvariable=num_absentees_var, font=("Arial", 14)).pack(fill="x")

    tk.Label(frame, text="Enter Roll Numbers (separated by space):", font=("Arial", 16)).pack(anchor="w", pady=10)
    tk.Entry(frame, textvariable=roll_numbers_var, font=("Arial", 14)).pack(fill="x")

    tk.Button(frame, text="Submit Attendance", command=process_attendance, bg="#28a745", fg="white", font=("Arial", 16), padx=10, pady=5).pack(pady=20)
    tk.Button(frame, text="Back", command=show_first_page, bg="#007bff", fg="white", font=("Arial", 16), padx=10, pady=5).pack()

# ---------------------------------------------
# MAIN APPLICATION
# ---------------------------------------------
root = tk.Tk()
root.title("Attendance Tracker")
root.geometry("1000x700")

# Tkinter variables
academic_year_var = tk.StringVar()
section_var = tk.StringVar()
subject_var = tk.StringVar()
num_absentees_var = tk.StringVar()
roll_numbers_var = tk.StringVar()

# Warn if credentials are missing
def _warn_missing_credentials():
    if not SENDER_EMAIL or not SENDER_PASSWORD:
        message = (
            "SENDER_EMAIL and/or SENDER_PASSWORD not found.\n\n"
            "Create a .env file in the project folder with:\n"
            "SENDER_EMAIL=your_email@gmail.com\n"
            "SENDER_PASSWORD=your_app_password\n\n"
            "Or set environment variables. See README for details."
        )
        messagebox.showwarning("Missing credentials", message)

root.after(200, _warn_missing_credentials)

create_first_page()
root.mainloop()
