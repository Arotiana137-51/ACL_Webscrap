#import  undetected_chromedriver as uc
from selenium import webdriver
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
    service = Service(r'C:\\Users\\Arotiana\\Documents\\CromeDriver\\chromedriver-win64\\chromedriver.exe')
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
   # options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.53 Safari/537.36')

    # Initialize undetected Chrome driver
    # driver = uc.Chrome(service=service, options=options)


    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    
    return driver


def random_delay(driver):
    WebDriverWait(driver,random.uniform(1, 5))
    

def random_scroll(driver, intensity=1.0):
    """Intensity-controlled scrolling with height change detection"""
    # Get current scroll position and height
    previous_scroll_y = driver.execute_script("return window.scrollY")
    previous_total_height = driver.execute_script("return document.body.scrollHeight")
    
    # Calculate scroll distance
    window_height = driver.execute_script("return window.innerHeight")
    scroll_ratio = random.uniform(1.3,2)
    scroll_distance = scroll_ratio * intensity * window_height
    
    # Execute scroll
    driver.execute_script(f"window.scrollBy(0, {scroll_distance});")
    
     # Calculate dynamic sleep time (intensity-based)
    min_sleep = 0.3 + (1 - intensity) * 0.7  # 0.3-1.0s range
    max_sleep = 1.2 + (1 - intensity) * 2.3  # 1.2-3.5s range
    time.sleep(random.uniform(min_sleep, max_sleep))
    
    # Check if scroll was effective
    new_scroll_y = driver.execute_script("return window.scrollY")
    new_total_height = driver.execute_script("return document.body.scrollHeight")
    
    # Return True if either the scroll position or page height changed
    return (new_scroll_y != previous_scroll_y) or (new_total_height != previous_total_height)


def scroll_to_bottom(driver, max_retries=15, scroll_pause_time=5):
    retries = 0
    previous_height = 0
    
    while True:
        # Scroll down to bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        
        # Wait for content to load
        time.sleep(scroll_pause_time)
        
        # Calculate new scroll height
        new_height = driver.execute_script("return document.body.scrollHeight")
        
        # Check if height has changed
        if new_height == previous_height:
            retries += 1
            # If no change for max_retries, stop scrolling
            if retries >= max_retries:
                print("Reached bottom of the page.")
                break
        else:
            retries = 0  # Reset counter if height changed
            previous_height = new_height


def get_last_path_element(url):
    """Get the element after the last slash in a URL"""
    try:
        parsed_url = urlparse(url)
        path = parsed_url.path
        last_element = path.rstrip('/').split('/')[-1]
        return last_element
    except Exception as e:
        return str(e)

# Function to check for redirection
def check_redirection(url):
    response = requests.get(url, allow_redirects=False)
    if response.status_code in (301, 302):
        print("Redirection detected to:", response.headers["Location"])
        return response.headers["Location"]
    return url


def get_elements_text(letter,driver, selector):
    """Get text from a matching element or from all matching elements"""
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        if elements:
            # If only one element is found, return its text
            if len(elements) == 1:
                return elements[0].text.strip()
            # If multiple elements are found, return a list of their texts
            else:
                return [elem.text.strip() for elem in elements if elem.text.strip()], True
        else:
            print(f"{letter}: No elements found with the selector: {selector}")
            #raise NoSuchElementException(f"No elements found with the selector: {selector}")
            return None
          
    except Exception as e:
        return str(e)
       
    
def click_btn(driver, selector):
    """Click the description button and wait for content to load"""
    try:
        description_button = WebDriverWait(driver, 5).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        )
        description_button.click()
        time.sleep(3)  # Wait for content to load
        return True
    except Exception as e:
        print(f"Error clicking description button: {str(e)}")
        return False

def get_elements_attribute(letter,driver, selector, attribute_name):
    """Get attribute values from all matching elements"""
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        return [elem.get_attribute(attribute_name).strip() for elem in elements if elem.get_attribute(attribute_name)]
    except Exception as e:
        print(f"{letter}: Error getting elements attribute: {str(e)}")
        return None

def click_if_needed(driver, selector):
    """Click element if it exists"""
    try:
        element = driver.find_element(By.CSS_SELECTOR, selector)
        element.click()
        time.sleep(3)  # Wait for any animations/loading
    except:
        pass

def extract_page_number(url):
    """Extract page number from URL"""
    try:
        parsed_url = urlparse(url)
        query_params = parse_qs(parsed_url.query)
        page = query_params.get('page', ['1'])[0]
        return int(page)
    except (ValueError, AttributeError, KeyError):
        return 1  # Default to page 1 if no page number found
    
def extract_currency_symbol(extracted_content):
    # Define the regex pattern to match non-digit, non-dot, non-comma characters
    pattern = r'[^\d,.]+'
    match = re.search(pattern, extracted_content)

    if match:
        symbol = match.group().strip()
    else:
        symbol = None  # Set to None if no symbol is found

    return symbol

def extract_number(extracted_content):
    # Define the regex pattern to match digit, dot, and comma characters
    pattern = r'[\d,.]+'
    match = re.search(pattern, extracted_content)
    
    if match:
        number_str = match.group().strip()
        # Remove any commas and replace comma with dot if comma is used as a decimal separator
        number_str = number_str.replace(',', '.').replace('.', '', number_str.count('.') - 1)
        return float(number_str)
    else:
        return None  

def click_center_of_screen(driver):
    # Get screen width and height
    screen_width = driver.execute_script("return window.innerWidth;")
    screen_height = driver.execute_script("return window.innerHeight;")
    
    # Calculate the center coordinates
    center_x = screen_width / 2
    center_y = screen_height / 2
    
    # Create an action chain to move to the center and click
    actions = ActionChains(driver)
    actions.move_by_offset(center_x, center_y).click().perform()
    
    # Move the mouse back to the initial position
    actions.move_by_offset(-center_x, -center_y).perform()


def click_and_extract_text(letter, driver, click_selector, text_selector):
    extracted_texts = []

    try:
        # Find elements by CSS selector that you want to click
        elements_to_click = driver.find_elements(By.CSS_SELECTOR, click_selector)

        for element in elements_to_click:
            # Click the element
            element.click()
            # Wait for the click action to take effect
            time.sleep(1)  # Consider using explicit waits for better reliability

            # Extract and store the displayed text for each clicked element
            text_display_element = driver.find_element(By.CSS_SELECTOR, text_selector)
            displayed_text = text_display_element.text
            extracted_texts.append(displayed_text)

    except Exception as e:
        print(f"{letter}: An error occurred: {e}")

    return extracted_texts
    
def hover_and_extract_text(letter,driver, hover_selector, text_selector):
    extracted_texts = []

    try:
        # Find elements by CSS selector that you want to hover over
        elements_to_hover = driver.find_elements(By.CSS_SELECTOR, hover_selector)

        # Create ActionChains object
        actions = ActionChains(driver)

        for element in elements_to_hover:
            # Hover over each element
            actions.move_to_element(element).perform()

            # Wait for the hover action to take effect
            time.sleep(1)  # Adjust hover timeout as needed

            # Extract and store the displayed text for each hovered element
            text_display_element = driver.find_element(By.CSS_SELECTOR, text_selector)
            displayed_text = text_display_element.text
            extracted_texts.append(displayed_text)

    except Exception as e:
        print(f"{letter}:An error occurred: {e}")

    return extracted_texts

def scrape_product_page(driver, url, position):
    try:
        driver.get(url)
        WebDriverWait(driver,3)
        
        # Extract page number from the URL
        page_number = extract_page_number(url)
        
        # Scrape all fields, handling multiple elements
        initial_data = {
            "A. UNIQUE_ID":datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
            "B. APPAREL_TYPOLOGY": get_elements_text("B",driver," div.breadcrumb.pdp__breadcrumb.new-pdp__breadcrumb > div > a > small > span"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": "Tom Tailor", #get_elements_text("D",driver,"div.ProductHeading > div.ProductHeading__brand--container > a"),
            "E. WEBSITE_URL": "https://www.tom-tailor.eu/",
            "F. PRODUCT_URL": url,
            "G. CITY": "Garstedter Weg 14, 22453 Hamburg",
            "H. WEBSITE_COUNTRY": "Germany",
            "I. GENDER":"Woman", # get_elements_text("I",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-child(3)"),
            "J. SEASON": "unknown", #get_elements_text(driver,"season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text("K",driver," div.breadcrumb.pdp__breadcrumb.new-pdp__breadcrumb > div:last-child> a > small > span"),
            "L. PRODUCT_NAME": get_elements_text("L",driver," div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.product-info__content > div.product-tile__content > h5 > span"),
            "M. PRODUCT_REFERENCE":get_elements_text("M",driver,"div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div > ul > li:last-child"),
            "N. PRODUCT_DESCRIPTION": "In Peripherals",#get_elements_text("N",driver,"div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.ProductHeroDescription.padt8.padb6 > div > div > div > p"),
            "O. PATTERN": get_elements_text("O",driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div > ul > li:nth-child(1)'),
            "P. STYLE": get_elements_text("P",driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div > ul > li:nth-child(2)'),
            "S. PERIPHERALS": get_elements_text("S",driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div > ul > li'),
            "U. MAIN_BODY_COLOUR": get_elements_text("U",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.variants.product-info__variants > h5"),
            "V. LIST_OF_COLORS":click_and_extract_text("V",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.variants.product-info__variants > div > div > div.variant.new-variant.variants__variant", "div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.variants.product-info__variants > h5"),
            "Z. RETAIL_CURRENCY":"€",# extract_currency_symbol(get_elements_text(driver," div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZA. PRICE": extract_number(get_elements_text("ZA-Promo",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.product-info__content > div.product-tile__content > div > span.product-tile__price.product-tile__price--dashed > span") or get_elements_text("ZA-Normal",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.product-info__content > div.product-tile__content > div > span.product-tile__price")),
            "ZB. ON_SALE": "Yes",#get_elements_text(driver,"sale-selector"),
            "ZC. SALE_TYPE":get_elements_text("ZC",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.product-info__content > div.flex.flex--justify-between.hide-mobile > div > div.flag.flag--danger.flag--sm.flag--ordered.flag--sale.new-flag > p > span") or "Normal sale",
            "ZD. SALE_AMOUNT": extract_number(get_elements_text("ZD-a",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.product-info__content > div.product-tile__content > div > span.product-tile__price.product-tile__price--sale") or get_elements_text("ZD-b",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.product-info__content > div.product-tile__content > div > span.product-tile__price")),
            "ZF. FEATURED_PRODUCT": get_elements_text("ZF",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.product-info__content > div.flex.flex--justify-between.hide-mobile > div > div.flag.flag--secondary.flag--sm.flag--ordered.flag--be-part.new-flag > p > span") or "Product not featured",
            "ZG. MAIN_IMAGE_URL": get_elements_attribute("ZG",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-gallery.pdp__gallery.new-product-gallery > div.hide-mobile.product-gallery--flex.pdp__gallery.new-product-gallery.product-gallery--flex-ct-7 > div:nth-child(2) > img", "src"),
            "ZH. IMAGE_META_TAGS": get_elements_attribute("ZH",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-gallery.pdp__gallery.new-product-gallery > div.hide-mobile.product-gallery--flex.pdp__gallery.new-product-gallery.product-gallery--flex-ct-7 > div:nth-child(2) > img", "alt"),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number, # Now using the extracted page number
            "Shirts and Trousers":position
        }
            # Click description button and get remaining elements
        # features_data = {} 
        # if click_btn(driver,'[data-testid="product-info-navigation-details-description-button"]'):
        #     features_data = {
        #         "S. PERIPHERALS": get_elements_text("S",driver,'#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(1) > div.AccordionItem__items > div > div > div > ul > li'),
        #         "ZE. CIEL_TEX_PRODUCT": "unknown",#get_elements_text("ZE",driver,"#product-materials-suppliers > div > div > ul > li > ul > li:nth-of-type(2) > p"),
        #         "Y. KNITTING_WOVEN":"unknown",#get_elements_text(driver,""),
        #     }

        materials_data = {}
        if (click_btn(driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div:nth-last-child(2) > button')):
            random_delay(driver)
            materials_data = {
                "W. FABRIC_COMPOSITION": get_elements_text("W",driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div:nth-child(1) > ul > li:nth-child(2)'),
                "R. NON_IRON":get_elements_attribute("R. Attr",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div:nth-child(2) > div > div:nth-child(1) > div > span > img","alt")+  get_elements_text("R",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div:nth-child(2) > div > div:nth-child(2) > ul > li"),
                "T. COUNTRY_OF_ORIGIN": get_elements_text("T",driver,"#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(2) > div.AccordionItem__items > div > div > div > ul > li:nth-child(2)"),
            }

        size_data = {}
        if click_btn(driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.size-buttons.new-size-buttons.product-info__sizes > div.size-buttons__text-container > div > div > span'):
            random_delay(driver)
            size_data = {
                "ZI. SIZE_SET": get_elements_text("ZI",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.size-buttons.new-size-buttons.product-info__sizes > div.size-buttons__text-container > div > div > div.modal.size-guides-modal.modal--centered.modal-enter-done > div.modal__body > div > div > table > thead > tr > th:not(:first-child)"),
                "ZJ. SIZE_AVAILABILITY": get_elements_text("ZJ",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.size-buttons.new-size-buttons.product-info__sizes > div.size-buttons__text-container > div > div > div.modal.size-guides-modal.modal--centered.modal-enter-done > div.modal__body > div > div > table > thead > tr > th:not(:first-child)"),
            }
            click_btn(driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.size-buttons.new-size-buttons.product-info__sizes > div.size-buttons__text-container > div > div > div.modal.size-guides-modal.modal--centered.modal-enter-done > div.modal__header > svg') # Close size selector
        
        fit_data = {}
        if click_btn(driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div:nth-child(6) > button'):
            random_delay(driver)
            fit_data = {
                "Q. FIT": get_elements_text("Q",driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div > ul > li:nth-child(1)'),
            }
        sustainability_data = {}
        if click_btn(driver,'#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div:nth-child(7) > button'):
            random_delay(driver)
            sustainability_data = {
                "X. IDENTIFIED AS SUSTAINABLE":get_elements_text("X",driver,"#newPdp > div.pdp__flex.new-pdp__flex > div.product-info.new-product-info.pdp__info.product-info--sticks-to-bottom.new-product-info--sticks-to-bottom > div.more-info.accordion.accordion--transparent.more-info--small > div > div > div > div > p:nth-child(2) > span") or "Unknown",
            }
        
      

        # Merge dictionaries while maintaining alphabetical order by prefix
        # Step 1: Combine all dictionaries into one
        all_data = {}
        all_data.update(initial_data)
        all_data.update(fit_data)
        all_data.update(sustainability_data)
        all_data.update(materials_data)
        all_data.update(size_data)
    

        # Step 2: Define prefix extraction logic
        def get_prefix(key):
            if '.' in key:
                prefix = key.split('.')[0]
                return prefix
            return key

        # Step 3: Create a new dictionary with sorted keys
        product_data = {}
        for key in sorted(all_data.keys(), key=get_prefix):
            product_data[key] = all_data[key]

        return product_data
        
    except TimeoutException:
        print(f" Timeout or selector not found on  {url} - possible CAPTCHA")
        input("Please solve CAPTCHA and press Enter to continue...")
        return scrape_product_page(driver, url, position)  # Retry after CAPTCHA
    except Exception as e:
        print(f"Error on {url}: {str(e)}")
        return None


def main():
    product_name = "TomTailor_Women_Blouses"
    product_link_selector = "a.product-tile--linked"
    driver = setup_chrome_driver()
    all_product_links = []
    seen_hrefs = set()
    idx = 1
    retries = 0
    max_retries = 8  # Reduced from 20 since we're more efficient
    scroll_attempts = 0

    try:
        main_url = check_redirection("https://www.tom-tailor.eu/women/clothing/blouses")
        driver.get(main_url)
        random_delay(driver)  # Reduced initial delay

        while retries < max_retries:
            # Smart scrolling with diminishing intensity
            scroll_intensity = max(0.3, 1 - (scroll_attempts * 0.1))
            random_scroll(driver, intensity=scroll_intensity)
            scroll_attempts += 1
            
            # Get links through JavaScript in one call
            hrefs = driver.execute_script(
                f"return Array.from(document.querySelectorAll('{product_link_selector}'))"
                ".map(a => a.href).filter(h => h);"
            )
            
            new_links = 0
            for href in hrefs:
                if href not in seen_hrefs:
                    all_product_links.append((href, idx))
                    seen_hrefs.add(href)
                    idx += 1
                    new_links += 1

            # Break early if no new links
            if new_links == 0:
                retries += 1
                time.sleep(1.5)  # Reduced delay between checks
            else:
                retries = 0
                time.sleep(0.8)  # Short delay when finding new links

        # Rest of your scraping logic remains the same
        current_batch = []
        all_products= []
        batch_size = 100

        for url, position in all_product_links:
            random_delay(driver)
            product_data = scrape_product_page(driver, url, position)
            print(url)
            if product_data:
                current_batch.append(product_data)
            
            # Save batch to Excel and clear memory
            if len(current_batch) >= batch_size:
                df = pd.DataFrame(current_batch)
                df.to_excel(f'{product_name}_batch_{position//batch_size}.xlsx', index=False)
                all_products.extend(current_batch)
                current_batch = []
                
            time.sleep(3)  # Polite delay between requests
        
        # Save any remaining products
        if current_batch:
            df = pd.DataFrame(current_batch)
            df.to_excel(f'{product_name}_batch_final.xlsx', index=False)
            all_products.extend(current_batch)
        
        # Save complete dataset
        df_all = pd.DataFrame(all_products)
        df_all.to_excel(f'{product_name}.xlsx', index=False)
            
    except Exception as e:
        print(f"Main error: {str(e)}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()