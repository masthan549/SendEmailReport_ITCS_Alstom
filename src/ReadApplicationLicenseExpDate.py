import ReadConfiguration, xlrd
from datetime import date
import datetime
from SendEmail import sendEmail



workbookPtr = xlrd.open_workbook(ReadConfiguration.ApplicationExpiryWorkbookLoc)

sheetNames = workbookPtr.sheet_names()

print("\nReading application details from workbook. Please wait...")
if (ReadConfiguration.sheetNameInWorkbook).strip() in sheetNames:
    sheetIndx = workbookPtr.sheet_by_name((ReadConfiguration.sheetNameInWorkbook).strip())
    appCoulumnNameNum = ord((ReadConfiguration.appNameColumnInWorkbook).upper())-65 
    appDateOfExp = ord((ReadConfiguration.appExpDate).upper())-65
    appExpiryDaysBefore = int(ReadConfiguration.appExpiryDaysBefore)
    numOfRows = sheetIndx.nrows
    rowIndx = 1
    todayDate = str(date.today())    
    
    appplicationExpiryInfo = ""
    
    while rowIndx < numOfRows:
        AppName = sheetIndx.cell_value(rowIndx, appCoulumnNameNum)
        AppExpDate = sheetIndx.cell_value(rowIndx, appDateOfExp)

        if len(str(AppExpDate).strip()) > 0:
            y,m,d,h,i,s = xlrd.xldate_as_tuple(int(AppExpDate), 0) #These are the dates from Excel
            
            #Compare the dates
            datesDiff =  date(y,m,d) - date(int(todayDate.split("-")[0]),int(todayDate.split("-")[1]),int(todayDate.split("-")[2]))
            if datesDiff.days < 0:
                appExpDetails = "\n\t\tApplication Name: \""+AppName+"\" Expired on date: "+str(date(y,m,d))                
                appplicationExpiryInfo = appplicationExpiryInfo+appExpDetails 
            elif datesDiff.days <= appExpiryDaysBefore:
                appExpDetails = "\n\t\tApplication Name: \""+AppName+"\" Expires on date: "+str(date(y,m,d))                
                appplicationExpiryInfo = appplicationExpiryInfo+appExpDetails

        rowIndx = rowIndx+1
        
    sendEmail(appplicationExpiryInfo, appExpiryDaysBefore)
    
else:
    print("There is no sheet with name given in Configuration file: "+ReadConfiguration.sheetNameInWorkbook)
