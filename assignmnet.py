import os
import re
import pdfplumber
from openpyxl import Workbook

def extract_info_from_cv(cv_path):
    # Initialize variables to store extracted information
    email = ""
    phone_number = ""
    text = ""

    # Extract text from the CV
    with pdfplumber.open(cv_path) as pdf:
        for page in pdf.pages:
            text += page.extract_text()

    # Extract email using regex
    email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
    email_match = re.search(email_pattern, text)
    if email_match:
        email = email_match.group(0)

    # Extract phone number using regex
    phone_pattern = r'\d{3}-\d{3}-\d{4}'
    phone_match = re.search(phone_pattern, text)
    if phone_match:
        phone_number = phone_match.group(0)

    return email, phone_number, text

def main():
    # Directory containing CV files
    cv_directory = r"C:\Users\user\Downloads\Sample2"
    cv_filename = "AarushiRohatgi.pdf"
    cv_path = os.path.join(cv_directory, cv_filename)

    # Extract info from CV
    email, phone_number, text = extract_info_from_cv(cv_path)

    # Print extracted info
    print("Email:", email)
    print("Phone Number:", phone_number)
    print("Text:", text)

    # Save the extracted info to an Excel file
    wb = Workbook()
    ws = wb.active
    ws.append(["Email", "Phone Number", "Text"])
    ws.append([email, phone_number, text])
    wb.save("cv_information.xlsx")

if __name__ == "__main__":
    main()
