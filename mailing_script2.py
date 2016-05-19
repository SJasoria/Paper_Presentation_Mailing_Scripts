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
pwd=""  #Add Password here
body = """\
<html>
  <head></head>
       <body>  
       <p>Dear %s!</p>
       <p>Thank you for submitting your abstract for Paper Presentation 2016. However, I regret to inform you that your abstract titled <b>"%s"</b> in the category <b>%s</b> has been not been selected for the next round.</p>
       <p>We highly appreciate your interest in research work and wish you godspeed in all your further endeavors.</p>
       <p>Keep checking our <a href='http://goo.gl/SCx9ub'>website</a> and <a href='https://goo.gl/8XzD6A'>Facebook Page</a> for other competitions. Hope to see you in <b>APOGEE 2016</b>.</p>
       <p>All the best for your exams.</p>
       <p>Best Regards,<br>
        </body>
</html>
"""  
list_email=excel.lEmail3('Rejected Abstracts.xlsx') #Give the name of excel sheet as argument before running the program
server=smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(username,pwd)

for detail in list_email:
    title=detail[0]
    category=detail[1]
    author_name=detail[2]
    author_email=detail[3]
    coauthor_name=detail[4]
    coauthor_email=detail[5]
    msg=MIMEMultipart()
    msg['From']=formataddr((str(Header('Paper Presentation Team, BITS Pilani', 'utf-8')), 'pep@bits-apogee.org'))
    msg['To']=author_email
    msg['Subject']='Paper Presentation- APOGEE 2016' 
    msg.attach(MIMEText(body %(author_name.encode('ascii'), title.encode('ascii'), category.encode('ascii')), 'html'))
    #fp=open('FINAL_PAPER_FORMAT.doc','rb')
    #doc = MIMEText(fp.read())
    #doc.add_header('Content-Disposition', 'attachment ;filename="Final Paper Format.doc"')
    #msg.attach(doc)
    #fp.close()
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
               msg['Subject']='Paper Presentation- APOGEE 2016' 
               msg.attach(MIMEText(body %(coauthor_name.encode('ascii'), title.encode('ascii'), category.encode('ascii')), 'html'))
               #fp=open('FINAL_PAPER_FORMAT.doc','rb')
               #doc = MIMEText(fp.read())
               #doc.add_header('Content-Disposition', 'attachment ;filename="Final Paper Format.doc"')
               #msg.attach(doc)
               #fp.close()
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
    
