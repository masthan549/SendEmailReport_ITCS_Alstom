import ReadConfiguration, xlrd

workbookPtr = xlrd.open_workbook(ReadConfiguration.LicenseSpreadSheetLoc)
numOfSheets = workbookPtr.nsheets
sheetNames = workbookPtr.sheet_names()
if "LicenseFile" in sheetNames:
sheetIndx = workbookPtr.sheet_by_name()


