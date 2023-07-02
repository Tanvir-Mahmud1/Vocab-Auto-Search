
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
# This will alow us to use 'Enter' or 'ESC' keys in automation process.
from selenium.webdriver.common.by import By
# We have import this "By" otherwise [browser.find_element(By.CSS_SELECTOR] will not work.
import time
# This can delay our Programmee as we want.
import openpyxl
#This is for working on excel file.

# This set of code is for quiting Excel Aplication if already opened.
# But to run [win32com.client] we first need to install [pip install pywin32] in command prompt.
import win32com.client as win32
# Quit the Excel application if already opened or else will skip the process if not opened. Exactly these two line of code is needed to do so.
xlApp = win32.gencache.EnsureDispatch('Excel.Application')
# This [try and except] clause will save and close excel application if not saved/close excel application if no changes was made.
try:
    xlApp.ActiveWorkbook.Save()
    xlApp.Application.Quit() 
except:
    xlApp.Application.Quit() 

chrmDriverPath = "C:\Chrome Driver\chromedriver.exe"
# Here Chrome Driver is located.
browser = webdriver.Chrome(chrmDriverPath)
# This will open Chrome browser.

# Below code will open excel file and read or write data on it.
xlFile = r"C:\Users\AC (Land)\Desktop\mahmud\Browser\words.xlsx"
# 'r' converts string to a raw string. Where single quote is needed like '', we can use [r" "] which is same and no error will occure.
workbook = openpyxl.load_workbook(filename= xlFile)
worksheet = workbook.active
# or
#worksheet = workbook['Sheet1']

ColNum = 'A'
# Here we can specify the column number where words are located to search.

for cell in worksheet[ColNum]:
    try:
        word = cell.value
        browser.get("https://www.bdword.com/english-to-bengali-meaning-" + word)

        srcWord = browser.find_element(By.CSS_SELECTOR, "div.align_text2")
        # We must have to use 'div' in ["div.align_text2"] which we got on hovering over the class name (which we want) while inspecting the browser content.
        resultWord = srcWord.text
        # Below code will offset one Column and the next line will write the searched word in Offseted Cell.
        nxtCol = cell.offset(row = 0, column = 1)
        nxtCol.value = resultWord

        print(resultWord)

    except:
        pass

workbook.save(xlFile)
workbook.close()

# These lines of code will open WorkBook (+WorkSheet) where searched data was saved.
open_wb = xlApp.Workbooks.Open(xlFile)
opnxl = open_wb.Worksheets('Sheet1') 
xlApp.Visible = True