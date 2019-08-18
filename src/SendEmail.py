from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Python code to illustrate Sending mail from  alstom email id 
import smtplib, datetime 

def sendSummaryEmail(message, ExpiryDays, fromEmailId, pwdOfEmail, toEmailId, LicenseAdminName, ccEmailId):

    try: 
        print("\nSending an email. Please wait...")        

        # creates SMTP session
        s = smtplib.SMTP('outlook.office365.com', 587) 

        # start TLS for security 
        s.starttls()

        # Authentication
        s.login(fromEmailId, pwdOfEmail) 
 
        # message to be sent
        dt= datetime.datetime.now()
        msg = MIMEMultipart('alternative') 
        msg['Subject'] = "Licensed application expiry details on: "+str(dt)
        msg['From'] = fromEmailId
        msg['To'] = toEmailId
        if len(ccEmailId) > 0:
            msg['cc'] = ccEmailId      

        EMailMessage = """\
        <html>
        <head></head>
        <body>
            Hi,<br><br>
                There are no application which are expiring in next """+ str(ExpiryDays)+""" days<br><br>
            Regards,<br>
            """+LicenseAdminName+"""<br>       
        </body>
        </html>
        """
             
        
        if len(message) > 2:
            EMailMessage = """\
            <html>
            <head></head>
            <body>
                Hi,<br>
                """+ message+""" days<br><br>
                Regards,<br>
                """+LicenseAdminName+"""<br>       
            </body>
            </html>
            """

            #msg = "Hi, \n"+EMailMessage+"\n\nRegards,\nRaghavendra's support team"
            #EMailMessage = 'Subject: {}\n\n{}'.format("Licensed application expiry details on: "+str(dt), bodyOfMessage) 

        message_mime = MIMEText(EMailMessage, 'html')
        msg.attach(message_mime)  

        # sending the mail 
        s.sendmail(fromEmailId, toEmailId, msg.as_string()) 
  
        # terminating the session 
        s.quit()
        print("\nEmail sent.")        
    except Exception as e:
        print("\nException raised While sending email: "+str(e))