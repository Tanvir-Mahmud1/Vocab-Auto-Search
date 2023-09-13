# This file contains Dictionary code.
# These codes needs moderation as per your need. So go through the code and and modify it accordingly.


from selenium.webdriver.common.by import By                                     # Without this, [By] in [browser.find_element(By.CSS_SELECTOR] will not work.
from selenium.webdriver.common.keys import Keys                                 # This will alow us to use 'Enter'/'ESC' (any Keyboard keys) in automation.
# import time                                                                   # This is for using delays in this code.
import ChangePart

driver = ChangePart.Driver_Select

###################################################
def bdword(word):
    driver.get("https://www.bdword.com/english-to-bengali-meaning-" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.align_text2")         # We must have to use 'div' in ["div.align_text2"] which we got on hovering over the class name.
    return srch_word

def eng2ban(word):
    driver.get("https://www.english-bangla.com/dictionary/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "span.format1")
    return srch_word
    
def OED(word):
    driver.get("https://www.oed.com/search/dictionary/?scope=Entries&q=" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.snippet")
    return srch_word
    
def Merrium(word):
    driver.get("https://www.merriam-webster.com/dictionary/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.vg")
    return srch_word
    
def collings(word):
    driver.get("https://www.collinsdictionary.com/dictionary/english/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.hom")
    return srch_word
