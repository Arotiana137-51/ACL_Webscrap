import undetected_chromedriver as uc
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import time
import random
import requests
from urllib.parse import urlparse, parse_qs
import re
from datetime import datetime

def setup_chrome_driver():
    print("Setting up the Chrome driver")
    service = Service(r'C:\\Users\\Arotiana\\Documents\\CromeDriver\\chromedriver-win64\\chromedriver.exe')
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")

    driver = uc.Chrome(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    print("Chrome driver initialized")
    return driver

def main():
    print("Main function execution")
    driver = setup_chrome_driver()

    # Your scraping logic here
    driver.get("https://www.tom-tailor.eu/women/clothing/blouses")
    print("Page loaded")
    
    try:
        element = WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.ID, " div.product-list__list > div > div > div > a"))
        )
        print("Element found")
        # Perform actions on the element here
    except TimeoutException:
        print("Loading took too much time!")
    except NoSuchElementException:
        print("Element not found!")

    driver.quit()
    print("Driver quit")

if __name__ == "__main__":
    main()
