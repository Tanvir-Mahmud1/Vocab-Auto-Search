# These codes needs moderation as per your need. So go through the code and and modify it accordingly.

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
xlApp.Application.Quit()

# chrmDriverPath = "C:/Chrome Driver/chromedriver.exe"
# # Here Chrome Driver is located.
# browser = webdriver.Chrome(chrmDriverPath)
# # This will open Chrome browser.


# Above code is creatin problem for versions. So i used bellow code.

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options)
os.environ['PATH'] += r"C:/Chrome Driver"


# Below code will open excel file and read or write data on it.
xlFile = r"D:\Projects\Vocab Search\words.xlsx"
# 'r' converts string to a raw string. Where single quote is needed like '', we can use [r" "] which is same and no error will occure.
workbook = openpyxl.load_workbook(filename= xlFile)
worksheet = workbook.active
# or
#worksheet = workbook['Sheet1']

for row in worksheet.iter_rows():
# Here, iter_rows() method returns a generator that iterates over all the rows in the worksheet. 
# You can specify the range of rows and columns to iterate over using the min_row, max_row, min_col, and max_col parameters.
    for cell in row:

        word = cell.value
        driver.maximize_window()
        driver.get("https://www.bdword.com/english-to-bengali-meaning-" + word)
        # print(cell.value)

        srcWord = driver.find_element(By.CSS_SELECTOR, "div.align_text2")
        # We must have to use 'div' in ["div.align_text2"] which we got on hovering over the class name (which we want) 
        # while inspecting the browser content.
        resultWord = srcWord.text
        print(resultWord)


driver.close()

workbook.save(xlFile)
workbook.close()

# These lines of code will open WorkBook (+WorkSheet) where searched data was saved.
open_wb = xlApp.Workbooks.Open(xlFile)
opnxl = open_wb.Worksheets('Sheet1') 
xlApp.Visible = True