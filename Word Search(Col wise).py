# These codes needs moderation as per your need. So go through the code and and modify it accordingly.
# To work this code properly, install and import bellow modules and [Very Very Important]=> Check "Chrome Driver" and "word" file path + location.
# [pip install openpyxl, pywin32]

import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys             # This will alow us to use 'Enter'/'ESC' (any Keyboard keys) in automation.
from selenium.webdriver.common.by import By                 # Without this, [By] in [browser.find_element(By.CSS_SELECTOR] will not work.
import time                                                 # This is for using delays in this code.
import openpyxl                                             # This is for working on excel file.

# These set of code is for quiting Excel Aplication if already opened.
import win32com.client as win32                             # [pip install pywin32] command is needed for this line of code.
xlApp = win32.gencache.EnsureDispatch('Excel.Application')  # Quits Excel app if already open else skips the process (if not opened).
try:                                                        # This [try & except] clause saves + closes excel app if not saved else closes excel if no need of 'save'.
    xlApp.ActiveWorkbook.Save()
    xlApp.Application.Quit() 
except:
    xlApp.Application.Quit() 

# These sets of code is for opening chrome browser.
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
# Check PATH location Bellow.................................................................................................................
os.environ['PATH'] += r"C:/Chrome Driver"
driver.maximize_window()

# Below code will open excel file and read/write data on it.
# Check xlfile Name Bellow...................................................................................................................
xlFile = r"F:\Projects\Vocab-Auto-Search\1500 Words.xlsx"    # 'r' converts string to a 'raw string'.
workbook = openpyxl.load_workbook(filename= xlFile)
worksheet = workbook.active                                  # or   worksheet = workbook['SheetName']
word_Column = 'A'                                            # Here we can specify the column number where words are located to search.

##################################################################
def bdword():
    try:
        word = cell.value
        driver.get("https://www.bdword.com/english-to-bengali-meaning-" + word)
        srcWord = driver.find_element(By.CSS_SELECTOR, "div.align_text2")         # We must have to use 'div' in ["div.align_text2"] which we got on hovering over the class name.
        resultWord = srcWord.text
        nxtCol = cell.offset(row = 0, column = 1)                                 # This code will offset one Column.
        nxtCol.value = resultWord
        # print(resultWord)
    except:
        pass

################################################################
for cell in worksheet[word_Column]:
    bdword()


driver.close()

################################################################
workbook.save(xlFile)                                        # Saves and closes the workbook which is opened by the driver.
workbook.close()

################################################################
open_wb = xlApp.Workbooks.Open(xlFile)                       # Opens the excel workbook in user-view mode.
# Check Sheet Name Bellow....................................................................................................................
opnxl = open_wb.Worksheets('1500 Words') 
xlApp.Visible = True