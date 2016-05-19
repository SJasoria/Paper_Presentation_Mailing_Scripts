# -*- coding: utf-8 -*-
import smtplib
import email
import time
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.header import Header
from email.utils import formataddr
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
import excel


fromadr = 'pep@bits-apogee.org'
username = fromadr
pwd=""  #Add Password here
body = """\
<html>
  <head></head>
       <body>  
       <p>Greetings %s!</p>
        <p>We are proud to announce and invite you to Paper Presentation, APOGEE 2016. </p>
        <p>First, we take this opportunity to express our gratitude for students like you for <br>
        participating in the event and making APOGEE 2015 a resounding success. </p>
        The event, which will be held during APOGEE, from the 25th to 28th February 2016, is <br>
        an excellent opportunity to showcase your research work and compete with the best <br>
        minds of the country. Also, with lots of  prizes to be won and a chance to get your <br>
        paper published in a reputed journal, this definitely, is one of the most exciting <br>
        and interesting events of APOGEE.</p>
        <p>However, this year the paper presentation event will follow a slightly different <br>
        timeline. We’ll invite abstracts in the beginning of the month of October. The result <br>
        for this round will be announced by mid November. The selected candidates will have <br>
        to send their papers by mid January 2016. After the plagiarism check of the papers,<br>
        shortlisted participants will be invited for the final event of Paper Presentation<br>
        at APOGEE 2016.</p>
        <p>For further notifications keep checking our <a href=http://goo.gl/SCx9ub> website </a> and 
        <a href=https://goo.gl/8XzD6A> Facebook page </a> </p>
        <p>Looking forward to your participation.</p>
        <p>Best Regards,<br></p>
        </body>
</html>
"""  

#body=body.decode('utf-8')

list_email=excel.lEmail('Completelist2015Abstracts_edit.xlsx') #Give the name of excel sheet as argument before running the program
sent={}
not_sent={}

server=smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(username,pwd)

for key,value in list_email.items():
    toadr=value
    msg=MIMEMultipart()
    msg['From']=formataddr((str(Header('Department of PEP', 'utf-8')), 'pep@bits-apogee.org'))
    msg['To']=toadr
    msg['Subject']='Paper Presentation- APOGEE 2016' 
    msg.attach(MIMEText(body %(key.encode('ascii')), 'html'))
    #image=MIMEApplication(open('pep poster.jpg','rb').read())
    #msg.add_header('Content-Disposition', 'attachment ;filename="Paper Presentation.jpg"')
    #msg.attach(image)
    fp=open('pep poster.jpg','rb')
    img = MIMEImage(fp.read())
    img.add_header('Content-Disposition', 'attachment ;filename="Paper Presentation.jpg"')
    msg.attach(img)
    fp.close()
    try:
       server.sendmail(fromadr,toadr,msg.as_string())
       time.sleep(60)
       print "Mail sent to "+toadr
       sent[key]=value
         
    except smtplib.SMTPRecipientsRefused:
       print "Mail not sent to "+toadr
       not_sent[key]=value
    
    except smtplib.SMTPServerDisconnected:
       print "Server disconneted"
       print "MAIL NOT SENT TO "+toadr
       not_sent[key]=value
       print "Connecting Again... "
       server=smtplib.SMTP('smtp.gmail.com',587)
       server.starttls()
       server.login(username.pwd)


excel.mail_log(sent, not_sent)

server.quit
    
