import ReadConfiguration

# Python code to illustrate Sending mail from  alstom email id 
import smtplib, datetime 

try: 
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

    message = 'Subject: {}\n\n{}'.format("Email from Masthan: "+str(dt), "TEXT") 
  
    # sending the mail 
    s.sendmail(fromEmailId, toEmailId, message) 
  
    # terminating the session 
    s.quit()
except Exception as e:
    print("\nException raised While sending email: "+str(e))