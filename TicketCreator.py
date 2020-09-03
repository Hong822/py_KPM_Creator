# -*- coding: utf-8 -*-
import os
import time
import timeit
import re
# import openpyxl
import win32timezone
import win32com.client as win32
import easygui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.ie.options import Options

import WebControl
import ActionItem 
import CreatorFunctions

def FindEditBoxValue(SheetName, KeyWord):
    for nRow in range(1, SheetName.UsedRange.Rows.Count + 1):
        for nCol in range (1, 50):
            if SheetName.Cells(nRow, nCol).Value == KeyWord:
                return SheetName.Cells(nRow, nCol + 1).Value


print("Ticket Creating has been started.\nPlease select an Excel file.")
nTotalStart = timeit.default_timer()

#Get Excel File Path
sExcelPath = easygui.fileopenbox('Select Ticket Creator')

if sExcelPath == None:    
    os._exit(0)
else:
    isExcel = re.search(".xlsm", sExcelPath)
    if isExcel == None:
        sErrorString = "Program will be terminated.\n\nWrong excel file: " + str(sExcelPath)
        easygui.msgbox(sErrorString, "Warning.")
        os._exit(0)
    else:
        # fExcelFile = openpyxl.load_workbook(filename = sExcelPath, data_only=True)
        excel = win32.gencache.EnsureDispatch('Excel.Application')
        try:
            fExcelFile = excel.Workbooks.Open(sExcelPath)
        except:
            print ("You have to Save excel file before run this program. ")
        fExcelFile.Save()
        
        RecordingSheet = fExcelFile.Worksheets('Recording')
        KpmCreateSheet = fExcelFile.Worksheets('kpmcreate')        
        
        if FindEditBoxValue(RecordingSheet, "Excel Visible") == 'O':
            excel.Visible = True
        else:
            excel.Visible = False
    
#ColumeDictionary Setting
RecColDict = {}
KpmColDict = {}
CreatorFunctions.DictionarySetting(RecordingSheet, RecColDict, 8, True)
CreatorFunctions.DictionarySetting(KpmCreateSheet, KpmColDict, 1, True)

print("Loading WebBrowser")
bBrowserType = FindEditBoxValue(RecordingSheet, "Browser Type")
if bBrowserType == "Firefox":
    KPMwebbrowser = webdriver.Firefox(executable_path=r"C:\WebDrivers\geckodriver.exe")
elif bBrowserType == "Chrome":
    KPMwebbrowser = webdriver.Chrome(executable_path=r"C:\WebDrivers\chromedriver.exe")
else:#ie
    # KPMwebbrowser = webdriver.Ie(executable_path=r"C:\WebDrivers\IEDriverServer_32bit.exe")
    # caps = webdriver.DesiredCapabilities.INTERNETEXPLORER
    # # capabilities['ie.enableFullPageScreenshot'] = False
    # # capabilities['ie.ensureCleanSession'] = True
    # caps["requireWindowFocus"] = True
    # caps["ignoreProtectedModeSettings"] = True
    # caps['ignoreZoomSetting'] = True
    # caps["javascriptEnabled"] = True
    # caps['INTRODUCE_FLAKINESS_BY_IGNORING_SECURITY_DOMAINS']=True
    ie_options = Options()
    ie_options.ignore_protected_mode_settings = True
    KPMwebbrowser = webdriver.Ie(executable_path=r"C:\WebDrivers\IEDriverServer_64bit.exe", options=ie_options)
    # KPMwebbrowser = webdriver.Ie(executable_path=r"C:\WebDrivers\IEDriverServer_32bit.exe", options=ie_options)

KPMwebbrowser.implicitly_wait(10)


if FindEditBoxValue(RecordingSheet, "WebType") == 'B2C':
    url = "https://quasi.vw.vwg/kpm/kpmweb/Index.action"
    WebControl.GoToURL(KPMwebbrowser, url)
    easygui.msgbox("Wait!!! Login and Go to Main page. And then press OK", "Please Login KPM")
    if FindEditBoxValue(RecordingSheet, "Brand") == 'AU':
        ActionListSheet = fExcelFile.Worksheets('AU_B2C') 
    else:
        ActionListSheet = fExcelFile.Worksheets('PO_B2C') 
    
else:
    # if bBrowserType == "Firefox" or bBrowserType == "Chrome":
    sID = FindEditBoxValue(RecordingSheet, "B2B Site ID")
    sPW = FindEditBoxValue(RecordingSheet, "B2B Site PW")
    if sID != None and sPW != None:
        url = "https://"+sID+":"+sPW+"@sso.volkswagen.de/kpmweb/index.do;jsessionid=ydr9irMbCOcxffBE27_cuxwg-64Lit_sYNMYCCLgz9PAWFo67z4s!761829767"            
        WebControl.GoToURL(KPMwebbrowser, url)
    else:
        url = "https://sso.volkswagen.de/kpmweb/index.do;jsessionid=ydr9irMbCOcxffBE27_cuxwg-64Lit_sYNMYCCLgz9PAWFo67z4s!761829767"
        WebControl.GoToURL(KPMwebbrowser, url)
        easygui.msgbox("Wait!!! Login and Go to Main page. And then press OK", "Please Login KPM")
    if FindEditBoxValue(RecordingSheet, "Brand") == 'AU':
        ActionListSheet = fExcelFile.Worksheets('AU_B2B')
    else:
        ActionListSheet = fExcelFile.Worksheets('PO_B2B')

pHNDLInfo = ActionItem.WindowHandleInfo(KPMwebbrowser.current_window_handle, None, KPMwebbrowser.current_window_handle)

#Create tickets!
lActionList = ActionItem.GetActionItemList(1, "StartEvent", ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict)      
CreatorFunctions.ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, 0)

for nRow in range(2, KpmCreateSheet.UsedRange.Rows.Count + 1):
    nRecordingCol = nRow + 7
    if KpmCreateSheet.Cells(nRow, 1).Value == None : # No KPM Number
        if RecordingSheet.Cells(nRecordingCol, 1).Value != None: # Timestamp is Ok
            # Then Create Ticket
            lActionList = ActionItem.GetActionItemList(nRow, "CreateTicket", ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict)      
            CreatorFunctions.ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, nRow)
            
            fExcelFile.Save() #Save Excel File
            
            if KpmCreateSheet.Cells(nRow, CreatorFunctions.FindDictVal(KpmColDict, "Re-Upload")).Value != "X":
                CreatorFunctions.UploadAttachment(  KPMwebbrowser, lActionList, pHNDLInfo, 
                                                    KpmCreateSheet, RecordingSheet, ActionListSheet, 
                                                    nRecordingCol, nRow,
                                                    RecColDict, KpmColDict )
            
            #Back to main screen
            lActionList = ActionItem.GetActionItemList(nRow, "Finish", ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict)      
            CreatorFunctions.ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, nRow)

            fExcelFile.Save() #Save Excel File
        else:
            pass
    else:
        #Only upload attachments
        if KpmCreateSheet.Cells(nRow, CreatorFunctions.FindDictVal(KpmColDict, "Re-Upload")).Value == "O":
            print ("Upload Attachment ticket / ", KpmCreateSheet.Cells(nRow, 1).Value)
            
            lActionList = ActionItem.GetActionItemList(nRow, "SearchTicket", ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict)      
            CreatorFunctions.ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, nRow)

            CreatorFunctions.UploadAttachment(  KPMwebbrowser, lActionList, pHNDLInfo, 
                                                KpmCreateSheet, RecordingSheet, ActionListSheet, 
                                                nRecordingCol, nRow,
                                                RecColDict, KpmColDict )
            
            #Back to main screen
            lActionList = ActionItem.GetActionItemList(nRow, "Finish", ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict)      
            CreatorFunctions.ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, nRow)

            fExcelFile.Save() #Save Excel File
        else:
             print ("Ticket is created. / ", KpmCreateSheet.Cells(nRow, 1).Value)

nTotalTime = timeit.default_timer() - nTotalStart
nHr = int(nTotalTime / 3600)
nMin = int((nTotalTime - 3600 * nHr) / 60)
nSec = int(nTotalTime % 60)

sLastString = "[Ticket Creation is Done!] Spend time = " + str(nHr) +"h, " + str(nMin) + "min, " + str(nSec) + "sec."
# easygui.msgbox(sLastString, "Finish")
print(sLastString)
