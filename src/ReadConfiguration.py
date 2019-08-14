import os

try:
    emailID_from = open("Configuration.txt").readlines()[0]
    emailID_to = open("Configuration.txt").readlines()[1]
    pwd = open("Configuration.txt").readlines()[2]
    sheetName = open("Configuration.txt").readlines()[3]
    sheetName = open("Configuration.txt").readlines()[3]    

    userFromEmailId = emailID_from.split("::")[1].strip()
    userToEmailId = emailID_to.split("::")[1].strip() 
    userPwd = pwd.split("::")[1].strip()
    LicenseSpreadSheetLoc = sheetName.split("::")[1].strip()
    
    #Check if spread sheet exist
    if not os.path.exists(LicenseSpreadSheetLoc):
        print("License file doesnt exist in given path!")

except Exception as e:
    print("\nException thrown while reading data from configuration file: "+str(e))