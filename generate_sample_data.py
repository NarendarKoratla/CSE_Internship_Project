# generate_sample_data.py
# Run this script to generate sample Excel files with dummy student data for testing.
# Usage: python generate_sample_data.py

import openpyxl
import os

SUBJECTS = {
    "1st": ["M1", "EC", "PPS"],
    "2nd": ["P&S", "OOPS", "DS"],
    "3rd": ["DAA", "PPL", "CN"],
    "4th": ["CC", "LP", "MC"],
}

SECTIONS = ["CSE-A", "CSE-B", "CSE-C"]

SAMPLE_STUDENTS = [
    (1, "Rahul Sharma", "rahul.sample@gmail.com"),
    (2, "Priya Reddy", "priya.sample@gmail.com"),
    (3, "Amit Kumar", "amit.sample@gmail.com"),
    (4, "Sneha Patel", "sneha.sample@gmail.com"),
    (5, "Vikram Singh", "vikram.sample@gmail.com"),
    (6, "Ananya Rao", "ananya.sample@gmail.com"),
    (7, "Karthik Nair", "karthik.sample@gmail.com"),
    (8, "Divya Joshi", "divya.sample@gmail.com"),
    (9, "Rohan Mehta", "rohan.sample@gmail.com"),
    (10, "Kavya Iyer", "kavya.sample@gmail.com"),
]

output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sample_data")
os.makedirs(output_dir, exist_ok=True)

for year_prefix, subjects in SUBJECTS.items():
    for section in SECTIONS:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Attendance"

        # Header row
        headers = ["Roll No", "Name", "Email"] + subjects
        ws.append(headers)

        # Student data rows (all absence counts start at 0)
        for roll, name, email in SAMPLE_STUDENTS:
            row = [roll, name, email] + [0] * len(subjects)
            ws.append(row)

        # Auto-adjust column widths for readability
        for col in ws.columns:
            max_length = max(len(str(cell.value or "")) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = max_length + 4

        filename = f"{year_prefix}_{section}.xlsx"
        filepath = os.path.join(output_dir, filename)
        wb.save(filepath)
        print(f"Created: {filepath}")

print(f"\nAll sample files generated in: {output_dir}")
print("Copy these files to your EXCEL_FILE_DIR to start testing.")
