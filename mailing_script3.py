# -*- coding: utf-8 -*-
import smtplib
import email
import time
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.header import Header
from email.utils import formataddr
import excel


fromadr = 'pep@bits-apogee.org'
username = fromadr
pwd=""   #Add Password here 
body = """
<html>
  <head></head>
       <body>  
       <p>Congratulations %s</p>
       <p>Your paper titiled <b>'%s'</b> has bagged <b>%s</b> postion in the <b>%s</b> category.</p>
       <p>Please provide your Bank a/c details along with IIFSc number so that we can send
       you the prize money.</p>
       <p>Best wishes for your future endeavours</p> 
       <p>Regards,<br></p>
        </body>
</html>
"""  
#list_email=excel.lEmail3('mail.xlsx') #Give the name of excel sheet as argument before running the program
list_email=excel.lEmail6('mail.xlsx')
print "Details of participants"
print list_email

server=smtplib.SMTP('smtp.gmail.com', 587)
print "Connected to server"
server.starttls()
server.login(username,pwd)
#qwerty=raw_input("Press any key to proceed")
for detail in list_email:
    '''title=detail[0] #for lEmail3
    category=detail[1]
    author_name=detail[2]
    author_email=detail[3]
    coauthor_name=detail[4]
    coauthor_email=detail[5]
    ref_no=detail[6]
    date=detail[7]
    time=detail[8]
    room=detail[9]'''
    category=detail[0]
    position=detail[1]
    title=detail[2]
    author_name=detail[3]
    author_email=detail[4]
    coauthor_name=detail[5]
    coauthor_email=detail[6]
    msg=MIMEMultipart()
    msg['From']=formataddr((str(Header('Paper Presentation Team, BITS Pilani', 'utf-8')), 'pep@bits-apogee.org'))
    msg['To']=author_email
    msg['Subject']='Paper Presentation, APOGEE 2016' 
    msg.attach(MIMEText(body %(author_name.encode('ascii'), title.encode('ascii'), position.encode('ascii'), category.encode('ascii')), 'html'))
    toadr=author_email
    try:
       server.sendmail(fromadr,toadr,msg.as_string())
       #time.sleep(30)
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
    if coauthor_name!='***':
               msg=MIMEMultipart()
               msg['From']=formataddr((str(Header('Paper Presentation Team, BITS Pilani', 'utf-8')), 'pep@bits-apogee.org'))
               msg['To']=coauthor_email
               msg['Subject']='Paper Presentation, APOGEE 2016' 
               msg.attach(MIMEText(body %(coauthor_name.encode('ascii'), title.encode('ascii'), position.encode('ascii'), category.encode('ascii')), 'html'))
               toadr=coauthor_email
               try:
                 server.sendmail(fromadr,toadr,msg.as_string())
                 #time.sleep(30)
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