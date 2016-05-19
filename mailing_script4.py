# -*- coding: utf-8 -*-
import smtplib
import email
import time
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.header import Header
from email.utils import formataddr
import excel


fromadr = 'web-journals@bits-apogee.org'
username = fromadr
pwd="" #Add Password here  
body = """
<html>
  <head></head>
       <body>  
       <p>Greetings %s!</p>
       <p>Wish you a Happy New Year and a wonderful 2016!</p>
       <p>We would like to inform you that the deadline to submit your article for <b>APOGEE Tech Review</b>
       web journal has been shifted to <b>10th January 2016</b>. Please adhere to the given deadline so that we have
       enough time to edit and upload your submissions.</p>
       <p>Thanks for your cooperation, and good luck for for APOGEE 2016!</p>
       <p>Organizing Team, <b>APOGEE 2016</b></p>
        </body>
</html>
"""  
list_email=excel.lEmail4('mail.xlsx') #Give the name of excel sheet as argument before running the program
server=smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(username,pwd)

for detail in list_email:
    name=detail[0]
    email=detail[1]
    msg=MIMEMultipart()
    msg['From']=formataddr((str(Header('APOGEE 2016, BITS Pilani', 'utf-8')), 'web-journals@bits-apogee.org'))
    msg['To']=email
    msg['Subject']='Introducing APOGEE Tech Review- Web Journal' 
    msg.attach(MIMEText(body %(name.encode('ascii')), 'html'))
    toadr=email
    try:
       server.sendmail(fromadr,toadr,msg.as_string())
       print "Mail sent to "+toadr
         
    except smtplib.SMTPRecipientsRefused:
       print "Mail not sent to "+toadr
    
    except smtplib.SMTPServerDisconnected:
       print "Server disconneted"
       print "MAIL NOT SENT TO "+toadr
       print "Connecting Again... "
       server=smtplib.SMTP('smtp.gmail.com',587)
       server.starttls()
       server.login(username,pwd)
   

server.quit
    
