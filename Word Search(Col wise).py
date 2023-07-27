# To work this properly, import these modules and Very Very Important Check "Chrome Driver" path and "word" file path.
# pip install openpyxl
# pip install pywin32
# import os



import os
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


# This set of code is for opening chrome browser.
options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
os.environ['PATH'] += r"C:/Chrome Driver"


# Below code will open excel file and read or write data on it.
# Check your workbook name........................... and write it bellow. Otherwise it will throw an error.
xlFile = r"D:\Projects\Vocab Search\1500_words.xlsx"
# 'r' converts string to a raw string. Where single quote is needed like '', we can use [r" "] which is same and no error will occure.

workbook = openpyxl.load_workbook(filename= xlFile)
worksheet = workbook.active
# or
#worksheet = workbook['Sheet1']

ColNum = 'A'
# Here we can specify the column number where words are located to search.

for cell in worksheet[ColNum]:
    try:
        driver.maximize_window()
        word = cell.value
        driver.get("https://www.bdword.com/english-to-bengali-meaning-" + word)

        srcWord = driver.find_element(By.CSS_SELECTOR, "div.align_text2")
        # We must have to use 'div' in ["div.align_text2"] which we got on hovering over the class name (which we want) while inspecting the browser content.
        resultWord = srcWord.text
        # Below code will offset one Column and the next line will write the searched word in Offseted Cell.
        nxtCol = cell.offset(row = 0, column = 1)
        nxtCol.value = resultWord

        print(resultWord)

    except:
        pass
    
driver.close()

workbook.save(xlFile)
workbook.close()

# These lines of code will open WorkBook (+WorkSheet) where searched data was saved.
open_wb = xlApp.Workbooks.Open(xlFile)

# Check your workbook Sheet name........................... and write it bellow. Otherwise it will throw an error.
opnxl = open_wb.Worksheets('1500 Words') 
xlApp.Visible = True