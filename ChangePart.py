# This file contains [File's Name, Location, output type, Dictionary (to chose from) etc.]
# These codes needs moderation as per your need. So go through the code and and modify it accordingly.

import Additional

# Driver_Location = r"C:/BrowserDriver"
# Driver_Location = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
# Driver_Location = r"C:/BrowserDriver/geckodriver.exe"
Driver_Location = r"C:\BrowserDriver\msedgedriver.exe"

# Driver_Select = Additional.chrome_driver(Driver_Location)
Driver_Select = Additional.edge_driver(Driver_Location)

xlFileName = r"F:\Projects\Vocab-Auto-Search\xlFiles\Kabil_Vocabulary_Memorized2.xlsx"              # 'r' converts string to a 'raw string'.
activeWorkSheet = 'Kabil'

# Here we can specify the column number where words are located to search.
row_from = 2
row_to = 5
col_from = 1
col_to = 1

# Here we can set where the output will be stored, more specifically Column no. next to the word containing cell.
offset_output = 6

# Available Dictionaries are [bdword, eng2ban, OED, Merrium, collings, Merrium_Sent]
def dictionary_name(word):
    import Dict_Code
    # I have imported [Dict_Code] module in [ChangePart] module and [ChangePart] module in [Dict_Code] module which is problematic and throws an exception. 
    # So try to avoid this type of circular import next time. But works if we import inside this function (Got the idea form ChatGPT "how to get rid of circular import error in python").
    return  Dict_Code.britannica(word)
