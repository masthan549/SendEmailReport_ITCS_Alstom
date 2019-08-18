import xml.etree.ElementTree as ET
import os

try:

    tree = ET.parse("Configuration.xml")
    root = tree.getroot()
   
    #XML tag names
    LicenseAdminName = ""
    AlstomEmailID = ""
    AlstomEmailPassword = ""
    AlstomLicenseWorkbookTrackerPath = ""    
    SheetNameInWorkbook = ""
    ApplicationNameColNameInSheet = ""
    DateOfExpiryColOfAnApplication = ""
    ProjectNameColWhereLicenseAppUsed = ""
    ProjectManagerNameCol = ""
    ProjectManagerEmailIdCol = ""
    LicenseExpireNumberDaysBefore = ""

    for element in root:
        if element.tag == "LicenseAdminName": LicenseAdminName = element.text
        if element.tag == "AlstomEmailID": AlstomEmailID = element.text
        if element.tag == "AlstomEmailPassword": AlstomEmailPassword = element.text
        if element.tag == "AlstomLicenseWorkbookTrackerPath": AlstomLicenseWorkbookTrackerPath = element.text
        if element.tag == "SheetNameInWorkbook": SheetNameInWorkbook = element.text
        if element.tag == "ApplicationNameColNameInSheet": ApplicationNameColNameInSheet = element.text
        if element.tag == "DateOfExpiryColOfAnApplication": DateOfExpiryColOfAnApplication = element.text
        if element.tag == "ProjectNameColWhereLicenseAppUsed": ProjectNameColWhereLicenseAppUsed = element.text
        if element.tag == "ProjectManagerNameCol": ProjectManagerNameCol = element.text
        if element.tag == "ProjectManagerEmailIdCol": ProjectManagerEmailIdCol = element.text
        if element.tag == "LicenseExpireNumberDaysBefore": LicenseExpireNumberDaysBefore = element.text                                                                                
    
    #Check if spread sheet exist
    if not os.path.exists(AlstomLicenseWorkbookTrackerPath):
        print("License file doesnt exist in given path!")


except Exception as e:
    print("\nException thrown while reading data from configuration file: "+str(e))