# To work this code properly, install and import necessary modules.

import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys             # This will alow us to use 'Enter'/'ESC' (any Keyboard keys) in automation.
from selenium.webdriver.common.by import By                 # Without this, [By] in [browser.find_element(By.CSS_SELECTOR] will not work.
import time                                                 # This is for using delays in this code.
import openpyxl                                             # This is for working on excel file.

import ChangePart                                           # This will import File Names and Locations which must be changed from system to system.

####################        These set of code is for quiting Excel Aplication if already opened.
import win32com.client as win32                             # [pip install pywin32] command is needed for this line of code.
xlApp = win32.gencache.EnsureDispatch('Excel.Application')  # Quits Excel app if already open else skips the process (if not opened).
try:                                                        # This [try & except] clause saves + closes excel app if not saved else closes excel if no need of 'save'.
    xlApp.ActiveWorkbook.Save()
    xlApp.Application.Quit() 
except:
    xlApp.Application.Quit() 

####################        These sets of code is for opening chrome browser.
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
os.environ['PATH'] += ChangePart.ChromeLocation
driver.maximize_window()

####################        Below code will open excel file and read/write data on it.
xlFile = ChangePart.xlFileName                              # 'r' converts string to a 'raw string'.
workbook = openpyxl.load_workbook(filename= xlFile)
worksheet = workbook.active                                 # or   worksheet = workbook['SheetName']
word_Column = ChangePart.word_col                           # Here we can specify the column number where words are located to search.


###################################################
def bdword():
    driver.get("https://www.bdword.com/english-to-bengali-meaning-" + word)
    srch_bdword = driver.find_element(By.CSS_SELECTOR, "div.align_text2")         # We must have to use 'div' in ["div.align_text2"] which we got on hovering over the class name.
    cell.offset(row = 0, column = 1).value = srch_bdword.text                     # [.offset] method will move to next Column, [.value] will paste the word to next column, [.text] will convert searched word to text.


###################################################
for cell in worksheet[word_Column]:
    word = cell.value
    try:
        bdword()
        
    except:
        pass


driver.close()

####################     Saves and closes the workbook which is opened by the driver.
workbook.save(xlFile)
workbook.close()

####################     Opens the excel workbook in user-view mode.
open_wb = xlApp.Workbooks.Open(xlFile)
opnxl = open_wb.Worksheets(ChangePart.activeWorkSheet) 
xlApp.Visible = True