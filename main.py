import openpyxl
import smtplib

# loading data sheet
data = openpyxl.load_workbook('/Users/srivenkatreddy/Library/CloudStorage/OneDrive-Personal/attendence.xlsx')

#choseing sheet
sheet = data['Sheet1']

r = sheet.max_row
c = sheet.max_column

# list of students who need to be reminded of attendence and lack of attendence
std_remind = []
lack_attd = []

staff_mail = 'srivenkatstock@gmail.com'

def savefile():
    data.save(r'/Users/srivenkatreddy/Library/CloudStorage/OneDrive-Personal/attendence.xlsx')
response = 1
while response is 1:
    n = int(input('Number of absenties : '))
    absenties = list(map(int,input("roll nos : ").split()))

    for rno in absenties:
        for i in range(2,r+1):
            if sheet.cell(row = i,column=1).value != rno:
                for j in range(3,c+1):
                    # temp = sheet.cell(row=i,column=j).value
                    sheet.cell(row=i, column=j).value += 1
                    savefile()
    response = int(input('Entering the attendence ? 1 ---> yes 0 ---> no'))