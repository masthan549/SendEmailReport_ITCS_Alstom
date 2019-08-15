import ReadConfiguration

# Python code to illustrate Sending mail from  alstom email id 
import smtplib, datetime 

def sendEmail(message, ExpiryDays):

    try: 
        print("\nSending an email. Please wait...")        
        #Read email id user name and password
        fromEmailId = ReadConfiguration.userFromEmailId
        toEmailId = ReadConfiguration.userToEmailId
        pwdOfEmail = ReadConfiguration.userPwd

        # creates SMTP session
        s = smtplib.SMTP('outlook.office365.com', 587) 

        # start TLS for security 
        s.starttls()

        # Authentication
        s.login(fromEmailId, pwdOfEmail) 
 
        # message to be sent
        dt= datetime.datetime.now()
        EMailMessage = 'Subject: {}\n\n{}'.format("Licensed application expiry details on: "+str(dt), "There are no application which are expiring in next "+ str(ExpiryDays)+" days")
        
        if len(message) > 2:
            bodyOfMessage = "Hi, \n"+message+"\n\nRegards,\nRaghavendra's support team"
            EMailMessage = 'Subject: {}\n\n{}'.format("Licensed application expiry details on: "+str(dt), bodyOfMessage) 

        # sending the mail 
        s.sendmail(fromEmailId, toEmailId, EMailMessage) 
  
        # terminating the session 
        s.quit()
        print("\nEmail sent.")        
    except Exception as e:
        print("\nException raised While sending email: "+str(e))