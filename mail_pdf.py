# -*- coding: utf-8 -*-
import smtplib
import email
import time
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMENonMultipart import MIMENonMultipart
from email.header import Header
from email.utils import formataddr
import excel
import base64


fromadr = 'pep@bits-apogee.org'
username = fromadr
pwd="" #Add Password here 
body = """
<html>
  <head></head>
       <body>  
       <p>Hello %s!</p>
       <p> Thanks a lot for submitting your paper for Paper Presentation- APOGEE 2016. However, I regret to inform you that your 
       paper titled <b>'%s'</b> in the category <b>%s</b> has not been selected for the the final round.Please find attached 
       plagiarism check report of your paper.</p>
       <p>We highly appreciate your interest in research work and wish you godspeed in all your future endeavors.</p>
       <p>Visit our <a href='http://bits-apogee.org/2016/'>website</a> and <a href='https://www.facebook.com/bitsapogee/'>Facebook page</a> 
       to participate in other competitions. Hope to see you in <b>APOGEE 2016</b>.</p>
       <p>Best Regards,<br></p>
        </body>
</html>
"""  
list_email=excel.lEmail3('Papers_rejected.xlsx') #Give the name of excel sheet as argument before running the program
server=smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(username,pwd)
print "Connected to server"

for detail in list_email:
    title=detail[0]
    category=detail[1]
    author_name=detail[2]
    author_email=detail[3]
    coauthor_name=detail[4]
    coauthor_email=detail[5]
    ref_no=detail[6]
    
    msg=MIMEMultipart()
    msg['From']=formataddr((str(Header('Paper Presentation Team, BITS Pilani', 'utf-8')), 'pep@bits-apogee.org'))
    msg['To']=author_email
    msg['Subject']='Paper Presentation, APOGEE 2016' 
    msg.attach(MIMEText(body %(author_name.encode('ascii'),title.encode('ascii'),category.encode('ascii')), 'html'))
    #attach PDF
    fp=open((ref_no+'.pdf').decode('ascii'),'rb')
    attach = MIMENonMultipart('application', 'pdf')
    payload = base64.b64encode(fp.read()).decode('ascii')
    attach.set_payload(payload)
    attach['Content-Transfer-Encoding'] = 'base64'
    fp.close()
    attach.add_header('Content-Disposition', 'attachment', filename = ref_no+'_report.pdf')
    msg.attach(attach)
    toadr=author_email
    try:
       server.sendmail(fromadr,toadr,msg.as_string())
       time.sleep(30)
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
               msg.attach(MIMEText(body %(coauthor_name.encode('ascii'),title.encode('ascii'),category.encode('ascii')), 'html'))
               #PDF Attachment
               fp=open(ref_no+'.pdf','rb')
               attach = MIMENonMultipart('application', 'pdf')
               payload = base64.b64encode(fp.read()).decode('ascii')
               attach.set_payload(payload)
               attach['Content-Transfer-Encoding'] = 'base64'
               fp.close()
               attach.add_header('Content-Disposition', 'attachment', filename = ref_no+'_report.pdf')
               msg.attach(attach)
               toadr=coauthor_email
               try:
                 server.sendmail(fromadr,toadr,msg.as_string())
                 time.sleep(30)
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
    
