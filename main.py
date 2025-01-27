import openpyxl
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os


# loading data sheet
data = openpyxl.load_workbook('/Users/srivenkatreddy/Documents/attendence.xlsx')
#choseing sheet
sheet = data['Sheet1']

r = sheet.max_row
c = sheet.max_column

# list of students who need to be reminded of attendance and lack of attendance
std_mail_id = []
lack_attd = []

FROM_EMAIL = os.getenv('FROM_EMAIL','srivenkatstock@gmail.com')
EMAIL_PASSWORD = os.getenv('EMAIL_PASSWORD','Sri@196768')

staff_mails = ['srivenkatreddy28@gmail.com','srivenkatreddy2208@gmail.com','srivenkatreddy28@gmail.com','srivenkatreddy2208@gmail.com']

warnings = {
    1: "Warning! You can take only one more day leave for ML class.",
    2: "Warning! You can take only one more day leave for DSA class.",
    3: "Warning! You can take only one more day leave for DBMS class.",
    4: "Warning! You can take only one more day leave for PYTHON class."
}

def savefile():
    data.save(r'/Users/srivenkatreddy/Library/CloudStorage/OneDrive-Personal/attendence.xlsx')
response = 1

subjects = {"jp":3,"dld":4,"dbms":5,"vegc":6}

while response is 1:
    sub = input("Enter the subject :")
    n = int(input('Number of absenties : '))
    print("Enter the Rollno's of absenties")
    absenties_roll_no = list(map(int,input("roll nos : ").split()))

    for rno in absenties_roll_no:
        for i in range(2,r+1):
            if sheet.cell(row = i,column=subjects[sub]).value != rno:
                print()

    response = int(input('Entering the attendance ? 1 ---> yes 0 ---> no'))