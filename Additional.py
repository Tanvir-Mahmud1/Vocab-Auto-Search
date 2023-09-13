# This file contains Some miscellaneous code.
# These codes needs moderation as per your need. So go through the code and and modify it accordingly.

# import win32com.client as win32                                # [pip install pywin32, comtypes] command is needed for this line of code.
import comtypes.client

####################  These set of code is for quiting Excel Aplication if already opened.
# xlApp = win32.gencache.EnsureDispatch('Excel.Application')     # Quits Excel app if already open else skips the process (if not opened).
xlApp = comtypes.client.CreateObject('Excel.Application')
# try:                                                           # This [try & except] clause saves + closes excel app if not saved else closes excel if no need of 'save'.
#     xlApp.ActiveWorkbook.Save()
#     xlApp.Application.Quit() 
# except:
#     xlApp.Application.Quit() 


###################         Different web browser code.
from selenium import webdriver
import os

# These 2 lines of code written bellow is to not close Chrome browser window after the task is completed.
# And also in [driver = webdriver.Chrome(options=options)], [(options=options)] is also for that purpose.

def chrome_driver(Driver_Location):
    options = webdriver.ChromeOptions()
    options.add_experimental_option("detach", True)
    driver = webdriver.Chrome(options=options)
    os.environ['PATH'] += Driver_Location
    return driver
    
def firefox_driver(Driver_Location):
    options = webdriver.FirefoxOptions()
    options.add_argument("--detach")
    driver = webdriver.Firefox(options=options)
    os.environ['PATH'] += Driver_Location
    return driver
    
def edge_driver(Driver_Location):
    options = webdriver.EdgeOptions()
    options.add_argument("--detach")
    driver = webdriver.Edge(options=options)
    os.environ['PATH'] += Driver_Location
    return driver