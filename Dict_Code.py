# This file contains Dictionary code.
# These codes needs moderation as per your need. So go through the code and and modify it accordingly.


from selenium.webdriver.common.by import By                                     # Without this, [By] in [browser.find_element(By.CSS_SELECTOR] will not work.
from selenium.webdriver.common.keys import Keys                                 # This will alow us to use 'Enter'/'ESC' (any Keyboard keys) in automation.
import time                                                                     # This is for using delays in this code.
import ChangePart

driver = ChangePart.Driver_Select

# ------------------------------------------------------------------------------------------------
def oxford_learner_def(word):     # Gets defination from oxfordlearner dictionary.
    driver.get(f"https://www.oxfordlearnersdictionaries.com/definition/english/{word.lower()}?q={word.lower()}")    # [word.lower()] is used as Capitalized word cannot find in the dictionary.
    srch_word = driver.find_element(By.CSS_SELECTOR, "span.def")
    return srch_word

def oxford_learner_exmp(word):     # Gets examples from oxfordlearner dictionary.
    driver.get(f"https://www.oxfordlearnersdictionaries.com/definition/english/{word.lower()}?q={word.lower()}")    # [word.lower()] is used as Capitalized word cannot find in the dictionary.
    srch_word = driver.find_element(By.CSS_SELECTOR, "ul.examples")
    return srch_word

# ------------------------------------------------------------------------------------------------
def britannica(word):     # Gets Definition and example sentences from britannica.com dictionary.
    driver.get("https://www.britannica.com/dictionary/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.sblocks")
    return srch_word

# ------------------------------------------------------------------------------------------------
def vocabulary_pShort(word):     # Gets Short summary of the word from vocabulary.com dictionary.
    driver.get("https://www.vocabulary.com/dictionary/" + word)
    # time.sleep(5)
    srch_word = driver.find_element(By.CSS_SELECTOR, "p.short") 
    return srch_word

def vocabulary_pLong(word):     # Gets Long summary of the word from vocabulary.com dictionary.
    driver.get("https://www.vocabulary.com/dictionary/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "p.long")
    return srch_word

# ------------------------------------------------------------------------------------------------
def Wordreference(word):     # Gets synonyms from wordreference dictionary.
    driver.get("https://www.wordreference.com/synonyms/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.clickable.engthes.clickTranslate.noTapHighlight")
    return srch_word

# ------------------------------------------------------------------------------------------------
def thesaurus_anti(word):     # Gets antonyms from thesaurus.com dictionary.
    driver.get("https://www.thesaurus.com/browse/" + word)
    srch_word = driver.find_element(By.XPATH, '//*[@id="root"]/div/main/div[2]/section/section[2]/section[5]/ul')
    return srch_word

# ------------------------------------------------------------------------------------------------
def bdword(word):   # Gets Bangla Meanings.
    driver.get("https://www.bdword.com/english-to-bengali-meaning-" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.align_text2")         # We must have to use 'div' in ["div.align_text2"] which we got on hovering over the class name.
    return srch_word

def eng2ban(word):  # Gets Bangla Meanings with (a few).
    driver.get("https://www.english-bangla.com/dictionary/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "span.format1")
    return srch_word

# ------------------------------------------------------------------------------------------------
def OED(word):  # Gets Defination from Oxford dictionary.
    driver.get("https://www.oed.com/search/dictionary/?scope=Entries&q=" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.snippet")
    return srch_word

# ------------------------------------------------------------------------------------------------
def Merrium(word):  # Gets Meaning from Merrium-webster dictionary.
    driver.get("https://www.merriam-webster.com/dictionary/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.vg")
    return srch_word

def Merrium_Sent(word):     # Gets Recent Examples used on the Web in Merrium dictionary.
    driver.get("https://www.merriam-webster.com/dictionary/" + word + "#examples")
    srch_word = driver.find_element(By.CSS_SELECTOR, "span.t.has-aq")
    return srch_word
    
# ------------------------------------------------------------------------------------------------
def collings(word): # Gets Meaning from Collings dictionary.
    driver.get("https://www.collinsdictionary.com/dictionary/english/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.hom")
    return srch_word

# ------------------------------------------------------------------------------------------------
# yourdictionary.com:
def thes_youdict_syn_ant(word):     # Gets synonyms and antonyms from thesaurus.yourdictionary.com dictionary (everything).
    driver.get("https://thesaurus.yourdictionary.com/" + word)
    srch_word = driver.find_element(By.CSS_SELECTOR, "div.definition-section")
    return srch_word

# Both dictionaries bellw is nedded as it is aparted two parts.
def thes_youdict_syn_ant1(word):     # Gets synonyms and antonyms from thesaurus.yourdictionary.com dictionary (one part).
    driver.get("https://thesaurus.yourdictionary.com/" + word)
    srch_word = driver.find_element(By.XPATH, '//*[@id="__layout"]/div/div[2]/div[1]/main/div[3]/div[1]')
    return srch_word

def thes_youdict_syn_ant2(word):     # Gets synonyms and antonyms from thesaurus.yourdictionary.com dictionary (another part but if exists).
    driver.get("https://thesaurus.yourdictionary.com/" + word)
    srch_word = driver.find_element(By.XPATH, '//*[@id="__layout"]/div/div[2]/div[1]/main/div[3]/div[2]')
    return srch_word
