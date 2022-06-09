# report
import smtplib
import os
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart



def gmail_login(user_name,password):
    smtp_server='smtp.gmail.com'
    port_number=587
    server=smtplib.SMTP(host=smtp_server,port=port_number)
    server.connect(host=smtp_server,port=port_number)
    server.starttls()
    server.login(user=username,password=password)
    return server

mgs = MIMEMultipart()
mgs['From'] = 'tasnimbkh1825006f@gmail.com'
mgs['To'] = 'tasnim1825006f@gmail.com'
mgs['Subject'] = 'test mail'


body="""
dear all,<br><br>Please send mail automatically work
"""
mgs.attach(MIMEText(body,'html')) 

# create attatchment
file_path=r"E:\Selenium"
file_name="result.xlsx"
file=open(file_path+"\\"+file_name,"rb")
payload=MIMEBase('application','octet-stream')
payload.set_payload(file.read())
file.close()
encoders.encode_base64(payload)
payload.add_header('content-desposition','attatchment',filename=file_name)
mgs.attach(payload)
server=gmail_login(user_name='tasnimbkh1825006f@gmail.com',password='bkh1825006f')
server.send_message(mgs)
server.quit()
