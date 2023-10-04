
# This module is totally independent, only import is "browseSelect". Variable part is separated, so don't worry.

from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchWindowException, NoSuchElementException, InvalidSessionIdException 
# import comtypes.client
import openpyxl

import Additional


# Chage Part-------------------------------------------------------------------------------------
Driver_Location = r"C:\BrowserDriver\msedgedriver.exe"
excel_file_path = r"F:\Projects\Vocab-Auto-Search\xlFiles\1500 Words_vocabulary.xlsx"
sheet_name = '1500 Words'
word_col = 'B'
start_row = 1493
end_row = 1495
col_offset_short = 1
col_offset_long = 2
# ------------------------------------------------------------------------------------------------


xlApp = Additional.xlApp
workbook = openpyxl.load_workbook(excel_file_path)
sheet = workbook[sheet_name]


for row in range(start_row, end_row+1):
    try:
        driver = Additional.edge_driver(Driver_Location)   # Chage Part---------------------------------------
        
        driver.set_window_size(500, 300)
        driver.set_window_position(1536 - 500, 864 - 300)      # [My Display Window Size: 1036x659;Screen Size: 1536x864, that's why i subtract it from the window size to get window position at the right bottom corner]

        cell = sheet[f"{word_col}{row}"]
        cell_value = cell.value

        print(f'[{row}:{cell_value}] is searching...')
        
        driver.get("https://www.vocabulary.com/dictionary/" + cell_value)

        src_word = driver.find_element(By.CSS_SELECTOR, "p.short")
        cell.offset(row = 0, column= col_offset_short).value = src_word.text

        src_word = driver.find_element(By.CSS_SELECTOR, "p.long")
        cell.offset(row = 0, column= col_offset_long).value = src_word.text

        driver.close()
    
    except NoSuchElementException:
        print(f'[{row}:{cell_value}] is not found in this dictionary.')
        
    except NoSuchWindowException:
        print("Browser window is closed by user.")

    except InvalidSessionIdException:
        print(f"An error occurred: {InvalidSessionIdException}")

    except Exception as e:
        print(f"An error occurred: {e}")
        
    finally:
        start_row = row + 1

try:
    workbook.save(excel_file_path)
    workbook.close()
except Exception as e:
    print(f"An error occurred: {e}")
        

open_wb = xlApp.Workbooks.Open(excel_file_path)
opnxl = open_wb.Worksheets(sheet_name)
xlApp.Visible = True