# -*- coding: utf-8 -*-
import openpyxl
import easygui
from selenium import webdriver
import WebControl
import ActionItem

def DictionarySetting(sSheetName, dDict, nRowOrColIdx, bCol):
    allData = sSheetName.UsedRange
    MaxRows = allData.Rows.Count + 1
    MaxCol = allData.Columns.Count + 1

    if bCol == True:
        for ColIdx in range(1, MaxCol):
            if sSheetName.Cells(nRowOrColIdx, ColIdx).Value != None:
                if sSheetName.Cells(nRowOrColIdx, ColIdx).Value in dDict:
                    sErrorString = "Sheet Name: " + sSheetName.title + " / Duplicated Key: " + sSheetName.Cells(nRowOrColIdx, ColIdx).Value
                    easygui.msgbox(sErrorString, "Duplicated Key")
                else:
                    dDict[sSheetName.Cells(nRowOrColIdx, ColIdx).Value] = ColIdx
    else:
        for RowIdx in range(1, MaxRows):
             if sSheetName.Cells(RowIdx, nRowOrColIdx).Value != None:
                if sSheetName.Cells(RowIdx, nRowOrColIdx).Value in dDict:
                    sErrorString = "Sheet Name: " + sSheetName.title + " / Duplicated Key: " + sSheetName.Cells(RowIdx, nRowOrColIdx).Value
                    easygui.msgbox(sErrorString, "Duplicated Key")
                else:
                    dDict[sSheetName.Cells(RowIdx, nRowOrColIdx).Value] = RowIdx
   
def FindDictVal(dDict, sKey):
    if sKey in dDict:
        return dDict[sKey]
    else:
        sErrorString = "No Key in Dictionary!!!" + sKey
        easygui.msgbox(sErrorString, "Duplicated Key")


def ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, RowinKPMCreateSheet):
    for item in lActionList:
        if item.ActionType == "CLICK":
           WebControl.ClickButton(KPMwebbrowser, item, pHNDLInfo)        
        elif item.ActionType == "DROPBOX":
            WebControl.SelectFromDropbox(KPMwebbrowser, item)
        elif item.ActionType == "INPUT_TEXT":
            WebControl.InputTexts( KPMwebbrowser, item)
        elif item.ActionType == "CopyToExcel":
            WebControl.CopyToExcel( KPMwebbrowser, item, KpmCreateSheet, RowinKPMCreateSheet)

def UploadAttachment(   KPMwebbrowser, lActionList, pHNDLInfo, 
                        KpmCreateSheet, RecordingSheet, ActionListSheet, 
                        nRecordingCol, nKPMCreateRow,
                        RecColDict, KpmColDict):
    #And Upload Attachment
    lActionList = ActionItem.GetActionItemList(nKPMCreateRow, "GoToUpload", ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict)      
    ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, nKPMCreateRow)

    #UploadFiles
    lActionList = ActionItem.GetActionItemList(nKPMCreateRow, "UploadFiles", ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict)      
    ExecuteActionList(KPMwebbrowser, lActionList, pHNDLInfo, KpmCreateSheet, nKPMCreateRow)

    #Uncheck Re-upload button
    RecordingSheet.Cells(nRecordingCol, FindDictVal(RecColDict, "Re-Upload")).Value = "X"

    