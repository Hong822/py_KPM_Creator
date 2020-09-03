# _*_ coding: utf-8 _*_
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import InvalidArgumentException
from selenium.common.exceptions import WebDriverException
from selenium.common.exceptions import TimeoutException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import timeit
import os

def GoToURL(WebDriver, url):
    try: 
        WebDriver.get(url)
    except WebDriverException as e:
        print (e)
        os._exit(0)    
    except :
        print ("GoToURL Error")
        os._exit(0)


# import win32gui, win32api, win32con
def WaitForChildWindow(KPMwebbrowser):
    all_windows = KPMwebbrowser.window_handles 
    nStart = timeit.default_timer()
    while len(all_windows) < 2:
        WaitTime = 20
        nEnd = timeit.default_timer()
        if nEnd - nStart > WaitTime:
            print ("Wait for ", WaitTime, " sec but fail to open Child window. ")
            return False
        all_windows = KPMwebbrowser.window_handles 
        time.sleep(1)
        print (" - WaitForChildWindow", int(nEnd - nStart), "sec")

def WaitForChildClose(KPMwebbrowser, pHNDLInfo):
    bFind = True
    nStart = timeit.default_timer()
    while bFind:
        WaitTime = 600
        nEnd = timeit.default_timer()
        if nEnd - nStart > WaitTime:
            print ("Wait for ", WaitTime, " sec but fail to Close Child window. ")
            os._exit(0)

        if pHNDLInfo.Child_Handle in KPMwebbrowser.window_handles:
            bFind = True
            time.sleep(1)
            print (" - WaitForChildClose", int(nEnd - nStart), "sec")
        else:
            bFind = False 

def WaitForStaleness(KPMwebbrowser, bButton):
    print("- WaitForStaleness")
    #Wait 60 sec for refresh
    if bButton:
        try:
            WebDriverWait(KPMwebbrowser, 20).until(EC.staleness_of(bButton))
        except Exception as e:
            print ("WaitForStaleness Fail!", e)
            return False
    else:
        print("No Button?")

def WaitForClickable(KPMwebbrowser, InputID, ComponentType):
    print("- WaitForClickable")
    if InputID:
        try:
            if ComponentType == "ID":
                WebDriverWait(KPMwebbrowser, 20).until(EC.element_to_be_clickable((By.ID, InputID)))
            elif ComponentType == "XPATH":
                WebDriverWait(KPMwebbrowser, 20).until(EC.element_to_be_clickable((By.XPATH, InputID)))
            elif ComponentType == "LINK_TEXT":
                WebDriverWait(KPMwebbrowser, 20).until(EC.element_to_be_clickable((By.LINK_TEXT, InputID)))
            elif ComponentType == "NAME":
                WebDriverWait(KPMwebbrowser, 20).until(EC.element_to_be_clickable((By.NAME, InputID)))
            elif ComponentType == "TAG_NAME":
                WebDriverWait(KPMwebbrowser, 20).until(EC.element_to_be_clickable((By.TAG_NAME, InputID)))
            elif ComponentType == "CLASS_NAME":
                WebDriverWait(KPMwebbrowser, 20).until(EC.element_to_be_clickable((By.CLASS_NAME, InputID)))
            elif ComponentType == "CSS_SELECTOR":
                WebDriverWait(KPMwebbrowser, 20).until(EC.element_to_be_clickable((By.CSS_SELECTOR, InputID)))
        except Exception as e:
            print ("WaitForClickable Fail! ", e)
            return False
    else:
        print("No Button?")

def ClickButtonAction(bButton):
    try:
        bButton.click()
    except Exception as e:
        print ("ClickButtonAction Fail!", e)
        return False

def FindElement(KPMwebbrowser, ComponentID, Type):
    if Type == "ID" or Type == "JAVASCRIPT":
        return KPMwebbrowser.find_element_by_id(ComponentID)
    elif  Type == "Xpath":
        return KPMwebbrowser.find_element_by_xpath(ComponentID)

def ClickButton(KPMwebbrowser, item, pHNDLInfo):
    print("***Click button ID: ", item.ComponentID, " / Wait Time:", item.WaitTime)    
    # if item.ComponentID == "btnGeraetetypStammdaten_middle":
    #     print ("Check")
    try:
        if WaitForClickable(KPMwebbrowser, item.ComponentID, item.ComponentType) == False:
            ClickButton(KPMwebbrowser, item, pHNDLInfo)
        
        if item.ClickType == "ID":
            bButton = FindElement(KPMwebbrowser, item.ComponentID, item.ClickType)
            if ClickButtonAction(bButton) == False:
                ClickButton(KPMwebbrowser, item, pHNDLInfo)
        elif item.ClickType == "JAVASCRIPT":
            print("***Click JavaScriptID ID: ", item.JavaScriptID) 
            KPMwebbrowser.execute_script(item.JavaScriptID)
            bButton = FindElement(KPMwebbrowser, item.ComponentID, item.ClickType)
        elif item.ClickType == "Xpath":
            print("***Click Xpath ID: ", item.XPathID) 
            bButton = FindElement(KPMwebbrowser, item.XPathID, item.ClickType)
            if ClickButtonAction(bButton) == False:
                ClickButton(KPMwebbrowser, item, pHNDLInfo)        
    except NoSuchElementException as e:        
        print (e)
        ClickButton(KPMwebbrowser, item, pHNDLInfo)
    except:
        ClickButton(KPMwebbrowser, item, pHNDLInfo)
    
    time.sleep(item.WaitTime)

    if item.HandleMoveState == "GoToChildWindow":
        if WaitForChildWindow(KPMwebbrowser) == False:
            ClickButton(KPMwebbrowser, item, pHNDLInfo)

        for window in KPMwebbrowser.window_handles:
            if window != pHNDLInfo.Parent_Handle:
                child_window = window
        pHNDLInfo.Child_Handle = child_window
        pHNDLInfo.Cur_Handle = child_window

        KPMwebbrowser.switch_to.window(child_window) 
    elif item.HandleMoveState == "BackToParentWindow":
        WaitForChildClose(KPMwebbrowser, pHNDLInfo)

        pHNDLInfo.Child_Handle = None
        pHNDLInfo.Cur_Handle = pHNDLInfo.Parent_Handle
        KPMwebbrowser.switch_to.window(pHNDLInfo.Parent_Handle)

        # WaitForClickable(KPMwebbrowser, "problem_middle")
        # bButton = KPMwebbrowser.find_element_by_id("problem_middle")
        # WaitForStaleness(KPMwebbrowser, bButton)
    else:
        if item.ClickAfter == "LOADING":
            if WaitForStaleness(KPMwebbrowser, bButton) == False:
                ClickButton(KPMwebbrowser, item, pHNDLInfo)

    

def SelectFromDropbox(KPMwebbrowser, item):
    if item.InputString != None:  
        item.InputString = str(item.InputString)
        print("***Dropbox ID: ", item.ComponentID, ", Input Text: ", item.InputString, " / Wait Time:", item.WaitTime)
        dropdown = None
        try:
            if WaitForClickable(KPMwebbrowser, item.ComponentID) == False:
                SelectFromDropbox(KPMwebbrowser, item)
            dropdown = Select(KPMwebbrowser.find_element_by_id(item.ComponentID))
            dropdown = Select(KPMwebbrowser.find(item.ComponentID))
        except NoSuchElementException as e:
            print (e)
        except:
            SelectFromDropbox(KPMwebbrowser, item)

        if item.SearchType == "Text":
            dropdown.select_by_visible_text(item.InputString)    
        elif item.SearchType == "Value":
            dropdown.select_by_value(item.InputString)  

        time.sleep(item.WaitTime)

def InputTexts(KPMwebbrowser, item):
    if item.InputString != None: 
        print("***InputText ID: ", item.ComponentID, ", Input Text: ", item.InputString, " / Wait Time:", item.WaitTime)
        try:
            if WaitForClickable(KPMwebbrowser, item.ComponentID) == False:
                InputTexts(KPMwebbrowser, item)

            TargetBox = KPMwebbrowser.find_element_by_id(item.ComponentID)
        except NoSuchElementException as e:
            InputTexts(KPMwebbrowser, item)
        except:
            InputTexts(KPMwebbrowser, item)

        TargetBox.clear()      
        try:
            TargetBox.send_keys(item.InputString)
        except InvalidArgumentException as e:
            # print ("No File!!! Please check again. This program will be terminated.", item.InputString)
            # os._exit(0)
            InputTexts(KPMwebbrowser, item)
        except:
            InputTexts(KPMwebbrowser, item)

        time.sleep(item.WaitTime)

def CopyToExcel(KPMwebbrowser, item, KpmCreateSheet, RowinKPMCreateSheet):
    print("***CopyToExcel / Editbox ID: ", item.ComponentID, " / Wait Time:", item.WaitTime)
    try:
        if WaitForClickable(KPMwebbrowser, item.ComponentID) == False:
            CopyToExcel(KPMwebbrowser, item, KpmCreateSheet, RowinKPMCreateSheet)
        TargetBox = KPMwebbrowser.find_element_by_id(item.ComponentID)
    except NoSuchElementException as e:
        print (e)
    except:
        CopyToExcel(KPMwebbrowser, item, KpmCreateSheet, RowinKPMCreateSheet)

    KpmNumber = TargetBox.get_attribute('value')      
    KpmCreateSheet.Cells(RowinKPMCreateSheet, 1).Value = KpmNumber
    print("\n     [Ticket Number]: ", KpmNumber, "\n")
    time.sleep(item.WaitTime)
    

# def get_window_by_caption(caption):
#     """
#     finds the window by caption and returns handle (int)
#     """
#     try:
#         hwnd = win32gui.FindWindow(None, caption)
#         return hwnd
#     except Exception as ex:
#         print('error calling win32gui.FindWindow ' + str(ex))
#         return -1


# def Find_Window_Handle_By_Class(className):
#     handle = win32gui.FindWindow(className, None)
#     return handle

# def Find_Window_By_Class(classname):
#     """
#     Finds the window with the given classname
#     """ 
#     try:
 
#         return Window(win32gui.FindWindow(classname, None))
 
#     except win32gui.error:
 
#         logging.exception("Error while finding the window")
 
#         return None


# def Find_Window_By_Caption(Caption):
#     """
#     Finds the window with the given classname
#     """ 
#     try:
 
#         return Window(win32gui.FindWindow(None, Caption))
 
#     except win32gui.error:
 
#         logging.exception("Error while finding the window")
 
#         return None

# def SetForegroundWindow(hwnd):
#      win32gui.SetForegroundWindow(hwnd)