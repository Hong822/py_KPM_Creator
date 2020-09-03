# -*- coding: utf-8 -*-
import CreatorFunctions
import os

class WindowHandleInfo():
    Parent_Handle = ""
    Child_Handle = ""
    Cur_Handle = ""

    def __init__(self, Parent_Handle = None, Child_Handle = None, Cur_Handle = None):
        self.Parent_Handle = Parent_Handle
        self.Child_Handle = Child_Handle
        self.Cur_Handle = Cur_Handle

class ActionItem ():
    ComponentID = ""
    ActionType = ""    
    ComponentType = ""
    ClickAfter = ""
    SearchType = ""
    InputString = ""
    SplitIdx = ""
    HandleMoveState = ""
    WaitTime = ""
    
    def __init__(self, ComponentID = None, ComponentType = None, ClickType = "ID", ActionType = "CLICK", ClickAfter = None, SearchType = "Text", InputString = None, SplitIdx = None, HandleMoveState = "STAY", ChildWindowItem = None, WaitTime = 0):
        self.ComponentID = ComponentID
        self.ComponentType = ComponentType
        self.ActionType = ActionType
        self.ClickAfter = ClickAfter
        self.SearchType = SearchType
        self.InputString = InputString
        self.SplitIdx = SplitIdx
        self.HandleMoveState = HandleMoveState       
        self.WaitTime = WaitTime

def FillinActionItem(TempActionItem, nRow, nCol, ActionListSheet, nKPMCreateRow, KpmColDict, KpmCreateSheet, sUploadComment, sUploadDocument, DocIndex):
    if ActionListSheet.Cells(1, nCol).Value == "Component":
        TempActionItem.ComponentID = ActionListSheet.Cells(nRow, nCol).Value    
    elif ActionListSheet.Cells(1, nCol).Value == "ActionType":
        TempActionItem.ActionType = ActionListSheet.Cells(nRow, nCol).Value
    elif ActionListSheet.Cells(1, nCol).Value == "ComponentType":
        TempActionItem.ComponentType = ActionListSheet.Cells(nRow, nCol).Value
    elif ActionListSheet.Cells(1, nCol).Value == "ClickAfter":
        TempActionItem.ClickAfter = ActionListSheet.Cells(nRow, nCol).Value
    elif ActionListSheet.Cells(1, nCol).Value == "SearchType":
        TempActionItem.SearchType = ActionListSheet.Cells(nRow, nCol).Value
    elif ActionListSheet.Cells(1, nCol).Value == "InputString":
        if ActionListSheet.Cells(nRow, nCol).Value == "Doc Comment":
            TempActionItem.InputString = str(sUploadComment[DocIndex])
        elif ActionListSheet.Cells(nRow, nCol).Value == "Documents":
            TempActionItem.InputString = str(sUploadDocument[DocIndex])
        else:
            TempVal = ActionListSheet.Cells(nRow, nCol).Value               
            nCol = CreatorFunctions.FindDictVal(KpmColDict, str(TempVal))
            CellString = KpmCreateSheet.Cells(nKPMCreateRow, nCol).Value
            if type(CellString) == float:                    
                CellString = int(CellString)
            TempActionItem.InputString = CellString
    elif ActionListSheet.Cells(1, nCol).Value == "SplitIndex":
        if ActionListSheet.Cells(nRow, nCol).Value != None:
            SplitString = TempActionItem.InputString.split(" ")                    
            SplitIdx = ActionListSheet.Cells(nRow, nCol).Value - 1
            if SplitIdx < 0: SplitIdx = 0
            SplitIdx = int(SplitIdx)                                        
            TempActionItem.InputString = SplitString[SplitIdx]
        TempActionItem.SplitIndex = ActionListSheet.Cells(nRow, nCol).Value
    elif ActionListSheet.Cells(1, nCol).Value == "HandleMoveState":
        TempActionItem.HandleMoveState = ActionListSheet.Cells(nRow, nCol).Value
    elif ActionListSheet.Cells(1, nCol).Value == "WaitTime":
        # if WaitSpeed == "Slow":
        TempActionItem.WaitTime = ActionListSheet.Cells(nRow, nCol).Value
        # elif WaitSpeed == "Fast":
        #     TempActionItem.WaitTime = ActionListSheet.Cells(nRow, nCol).Value
        # else:
        #     TempActionItem.WaitTime = ActionListSheet.Cells(nRow, nCol).Value
    else:     
        pass

def GetActionItemList( nKPMCreateRow, ActionType, ActionListSheet, RecordingSheet, KpmCreateSheet, RecColDict, KpmColDict):
    LTempList = list()
    allData = ActionListSheet.UsedRange
    MaxRows = allData.Rows.Count + 1
    MaxCol = allData.Columns.Count + 1    

    nStepCol = nExecuteCol = nStartCol = nEndCol = 0
    for nCol in range (1, MaxCol):
        if ActionListSheet.Cells(1, nCol).Value == "Step":
            nStepCol = nCol
        elif ActionListSheet.Cells(1, nCol).Value == "Execute":
            nExecuteCol = nCol
        elif ActionListSheet.Cells(1, nCol).Value == "Component":
            nStartCol = nCol
        elif ActionListSheet.Cells(1, nCol).Value == "Comment":
            nEndCol = nCol

    sUploadComment = sUploadDocument = None
    if ActionType == "UploadFiles":
        nKPMCreateCol = CreatorFunctions.FindDictVal(KpmColDict, "Doc Comment")
        if type(KpmCreateSheet.Cells(nKPMCreateRow, nKPMCreateCol).Value) == float:
            TempComment = int(KpmCreateSheet.Cells(nKPMCreateRow, nKPMCreateCol).Value)
        else:
            TempComment = KpmCreateSheet.Cells(nKPMCreateRow, nKPMCreateCol).Value
        sUploadComment = (str(TempComment)).split('\n')
        nKPMCreateCol = CreatorFunctions.FindDictVal(KpmColDict, "Documents")
        sUploadDocument = (KpmCreateSheet.Cells(nKPMCreateRow, nKPMCreateCol).Value).split('\n')
        nActionIdx = len(sUploadComment)    # To repeat the number of attachment file.
    else:
        nActionIdx = 1

    for DocIndex in range(0, nActionIdx):
        for nRow in range(2, MaxRows):        
            if ActionListSheet.Cells(nRow, nStepCol).Value != ActionType or ActionListSheet.Cells(nRow, nExecuteCol).Value == 'X':
                continue                        
            TempActionItem = ActionItem()
            
            #Fill in action item with whole column values
            for nCol in range (nStartCol, nEndCol):
                if ActionListSheet.Cells(nRow, nCol).Value == None:
                    continue            
                FillinActionItem(TempActionItem, nRow, nCol, ActionListSheet, nKPMCreateRow, KpmColDict, KpmCreateSheet, sUploadComment, sUploadDocument, DocIndex)
                
            LTempList.append(TempActionItem)  #Add to action item list
    return LTempList