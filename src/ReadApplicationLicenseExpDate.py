import ReadConfiguration, xlrd, sys
from datetime import date
from SendEmail import sendSummaryEmail

def prepareEmailDetails(mailType):
    
    rowIndx = 1
    appplicationExpiryInfo = "" 
    
    while rowIndx < numOfRows:
        AppName = sheetIndx.cell_value(rowIndx, appCoulumnNameNum)
        AppExpDate = sheetIndx.cell_value(rowIndx, appDateOfExp)
        ProjectNameColWhereLicenseAppUsed_val = sheetIndx.cell_value(rowIndx, ProjectNameColWhereLicenseAppUsed)
        ProjectManagerNameCol_val = sheetIndx.cell_value(rowIndx, ProjectManagerNameCol)
        ProjectManagerEmailIdCol_val = sheetIndx.cell_value(rowIndx, ProjectManagerEmailIdCol)
        #ProjectManagerEmailIdCol_val = ProjectManagerEmailIdCol_val.split(",")
        
        if(len(AppName.strip()) == 0):
            AppName = "App name is empty at row number: "+str(rowIndx+1)
            
        if(len(ProjectNameColWhereLicenseAppUsed_val.strip()) == 0):
            ProjectNameColWhereLicenseAppUsed_val = "project name is empty at row number: "+str(rowIndx+1)
            
        if(len(ProjectManagerNameCol_val.strip()) == 0):
            ProjectManagerNameCol_val = "project manager name is empty at row number: "+str(rowIndx+1)
            
        if(len(ProjectManagerEmailIdCol_val) == 0):
            ProjectManagerEmailIdCol_val = "project manager email is empty at row number: "+str(rowIndx+1)                                    

        if len(str(AppExpDate).strip()) > 0:
            y,m,d,h,i,s = xlrd.xldate_as_tuple(int(AppExpDate), 0) #These are the dates from Excel
            
            #Compare the dates
            datesDiff =  date(y,m,d) - date(int(todayDate.split("-")[0]),int(todayDate.split("-")[1]),int(todayDate.split("-")[2]))
            if datesDiff.days < 0:
                appExpDetails = "\t\t<br><br><font color=\"red\">Application Name: \""+AppName+"\" used in program \""+ProjectNameColWhereLicenseAppUsed_val+"\" by Manager \""+ProjectManagerNameCol_val+"\", Expired on date: "+str(date(y,m,d))+"</font>"               
                appplicationExpiryInfo = appplicationExpiryInfo+appExpDetails 
            elif datesDiff.days <= appExpiryDaysBefore:
                appExpDetails = "\t\t<br><br>Application Name: \""+AppName+"\" used in program \""+ProjectNameColWhereLicenseAppUsed_val+"\" by Manager \""+ProjectManagerNameCol_val+"\", Expires on date: "+str(date(y,m,d))                
                appplicationExpiryInfo = appplicationExpiryInfo+appExpDetails
        else:
            appExpDetails = "\t\t<br><br><font color=\"red\">Application Name: \""+AppName+"\" used in program \""+ProjectNameColWhereLicenseAppUsed_val+"\" by Manager \""+ProjectManagerNameCol_val+"\", Expiry date not mentioned. </font>"                
            appplicationExpiryInfo = appplicationExpiryInfo+appExpDetails

        if (mailType == "Individual") and len(ProjectManagerEmailIdCol_val) > 0:
            sendSummaryEmail(appplicationExpiryInfo, appExpiryDaysBefore, ReadConfiguration.AlstomEmailID, ReadConfiguration.AlstomEmailPassword, ProjectManagerEmailIdCol_val, ReadConfiguration.LicenseAdminName, ReadConfiguration.AlstomEmailID)
            appplicationExpiryInfo = ""    
        elif (mailType == "Individual") and len(ProjectManagerEmailIdCol_val) == 0:
            sendSummaryEmail("woner Email is empty for Application: "+AppName, appExpiryDaysBefore, ReadConfiguration.AlstomEmailID, ReadConfiguration.AlstomEmailPassword, ReadConfiguration.AlstomEmailID, ReadConfiguration.LicenseAdminName, "")            
            appplicationExpiryInfo = ""
            
        rowIndx = rowIndx+1
        
    if mailType == "Summary":        
        sendSummaryEmail(appplicationExpiryInfo, appExpiryDaysBefore, ReadConfiguration.AlstomEmailID, ReadConfiguration.AlstomEmailPassword, ReadConfiguration.AlstomEmailID, ReadConfiguration.LicenseAdminName, "")

#Read workbook
workbookPtr = xlrd.open_workbook(ReadConfiguration.AlstomLicenseWorkbookTrackerPath)
sheetNames = workbookPtr.sheet_names()

#check if sheet name exist in workbook
if not (ReadConfiguration.SheetNameInWorkbook).strip() in sheetNames:
    print("There is no sheet with name given in Configuration file: "+ReadConfiguration.AlstomLicenseWorkbookTrackerPath)
    sys.exit()
    
#Read globals from Configuration file
appCoulumnNameNum = ord((ReadConfiguration.ApplicationNameColNameInSheet).upper())-65 
appDateOfExp = ord((ReadConfiguration.DateOfExpiryColOfAnApplication).upper())-65
appExpiryDaysBefore = int(ReadConfiguration.LicenseExpireNumberDaysBefore)
ProjectNameColWhereLicenseAppUsed = ord(ReadConfiguration.ProjectNameColWhereLicenseAppUsed)-65
ProjectManagerNameCol = ord(ReadConfiguration.ProjectManagerNameCol)-65
ProjectManagerEmailIdCol = ord(ReadConfiguration.ProjectManagerEmailIdCol)-65

sheetIndx = workbookPtr.sheet_by_name((ReadConfiguration.SheetNameInWorkbook).strip())
numOfRows = sheetIndx.nrows
todayDate = str(date.today())

print("\nReading application details from workbook. Please wait...")
prepareEmailDetails("Summary") #This sends only one summary email to admin of all the licenses
prepareEmailDetails("Individual") #This sends only one summary email to admin of all the licenses
 