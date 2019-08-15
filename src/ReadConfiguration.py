import os

try:
    emailID_from = open("Configuration.txt").readlines()[0]
    emailID_to = open("Configuration.txt").readlines()[1]
    pwd = open("Configuration.txt").readlines()[2]
    workbookName = open("Configuration.txt").readlines()[3]
    sheetName = open("Configuration.txt").readlines()[4]
    appNameInSheetColumn = open("Configuration.txt").readlines()[5]
    dateOfExpOfApplication = open("Configuration.txt").readlines()[6]
    appExp = open("Configuration.txt").readlines()[7]
                   
    userFromEmailId = emailID_from.split("::")[1].strip()
    userToEmailId = emailID_to.split("::")[1].strip() 
    userPwd = pwd.split("::")[1].strip()
    ApplicationExpiryWorkbookLoc = workbookName.split("::")[1].strip()
    sheetNameInWorkbook = sheetName.split("::")[1].strip()
    appNameColumnInWorkbook = appNameInSheetColumn.split("::")[1].strip()
    appExpDate = dateOfExpOfApplication.split("::")[1].strip()
    appExpiryDaysBefore = appExp.split("::")[1].strip()
    
    #Check if spread sheet exist
    if not os.path.exists(ApplicationExpiryWorkbookLoc):
        print("License file doesnt exist in given path!")

except Exception as e:
    print("\nException thrown while reading data from configuration file: "+str(e))