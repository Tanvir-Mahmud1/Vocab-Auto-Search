# got this code from ChatGPT


from selenium.common.exceptions import NoSuchWindowException
from selenium.common.exceptions import InvalidSessionIdException
from selenium.common.exceptions import NoSuchElementException


def handle_exceptions(row, cell_value, e):
    
    if isinstance(e, NoSuchElementException):
        print(f'[{row}:{cell_value}] is not found in this dictionary.')
        
    elif isinstance(e, NoSuchWindowException):
        print("Browser window is closed by user.")
        
    elif isinstance(e, InvalidSessionIdException):
        print(f"An error occurred: {InvalidSessionIdException}")
        
    else:
        print(f"An error occurred: {e}")
