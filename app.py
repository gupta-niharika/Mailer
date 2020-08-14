import smtplib, ssl
import pandas as pd
import xlrd 

port = 465 #gmail

sender = 'sender@gmail.com' #sender's address
password = 'password' #password

loc = ("mail_list.xlsx") 
  
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 

msg = """\
Subject: Python Mailer
Hey this is a mail sent from the python mail bot.  :P
"""
context = ssl.create_default_context()

print("Starting to send")

with smtplib.SMTP_SSL("smtp.gmail.com",port, context=context ) as server:
    server.login(sender, password)
    #for j in range(1,11):  #this loop is for spamming 10 times
    for i in range(sheet.nrows): 
        server.sendmail(sender, sheet.cell_value(i, 0), msg)
print("sent")