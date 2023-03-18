import os
from openpyxl import Workbook,load_workbook
from Email import send_email

wb= load_workbook('E:\FaceRecognition\ATTENDANCE .xlsx')
ws= wb.active

perecentagecolumn= 'AK'
names= 'A'
emails= 'C'
people=10
i='2'
while int(i)<people+1:
    if int(ws[perecentagecolumn+i].value)<75:
        send_email(
          subject="Attendance Reminder",
          name= str(ws[names+i].value),
          receiver_email=str(ws[emails+i].value),
          percentage=str(ws[perecentagecolumn+i].value)    
        )
    i=str(int(i)+1)