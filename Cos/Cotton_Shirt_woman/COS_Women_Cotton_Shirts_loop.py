from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.common.action_chains import ActionChains
import pandas as pd
import os
import random
import time
import requests
from urllib.parse import urlparse, parse_qs
import re
from selenium import webdriver
from datetime import datetime


def setup_chrome_driver():
    try:
        service = Service(r'C:\\Users\\Arotiana\\Documents\\CromeDriver\\chromedriver-win64\\chromedriver.exe',
                )
        
        options = Options()
        
        # Essential Headless Configuration
        options.add_argument(f"--user-data-dir=C:\\Users\\Arotiana\\AppData\\Local\\Google\\Chrome\\User Data")        
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_experimental_option("detach", True) 
        
        # Stealth Configuration
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
        options.add_experimental_option("useAutomationExtension", False)
        
        
        # Network Optimization
        options.add_argument("--disable-extensions")
        options.add_argument("--dns-prefetch-disable")
        options.add_argument("--disable-translate")
        
        # Memory Management
        options.add_argument("--single-process")
        options.add_argument("--disable-accelerated-2d-canvas")
        
        # Disable features that consume resources
        prefs = {
                "profile.default_content_setting_values": {
                    "images": 2,  # Disable images
                    "plugins": 2,
                    "popups": 2,
                    "geolocation": 2
                },
                "credentials_enable_service": False,  # Disable password saving popups
                "profile.password_manager_enabled": False
            }
        options.add_experimental_option("prefs", prefs)

        driver = webdriver.Chrome(service=service, options=options)
        
        #9. Enhanced Stealth Measures
        driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
                "source": """
                    const originalQuery = window.navigator.permissions.query;
                    window.navigator.permissions.query = (parameters) => (
                        parameters.name === 'notifications' ?
                            Promise.resolve({ state: Notification.permission }) :
                            originalQuery(parameters)
                    );
                    
                    Object.defineProperty(navigator, 'webdriver', {
                        get: () => false
                    });
                    
                    window.chrome = {
                        app: {},
                        runtime: {},
                        webstore: {}
                    };
                """
            })
        
        return driver

    except Exception as e:
        print(f"Driver initialization failed: {str(e)}")
        raise

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
                return [elem.text.strip() for elem in elements if elem.text.strip()]
        else:
            print ((f"{letter}: No elements found with the selector: {selector}"))
            return "Unknown"
    except Exception as e:
            return "Unknown"

    
def click_btn(driver, selector):
    """Click the description button and wait for content to load"""
    try:
        description_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        )
        description_button.click()
        time.sleep(3)  # Wait for content to load
        return True
    except Exception as e:
        print(f"Error clicking button: {selector}")
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

def get_base_url(long_url):
    # Parse the URL into its components
    parsed_url = urlparse(long_url)
    
    # Construct base URL from scheme and netloc components
    base_url = f"{parsed_url.scheme}://{parsed_url.netloc}/"
    return base_url


def extract_number(extracted_content):
    # Define the regex pattern to match digit, dot, and comma characters
    pattern = r'[\d,.]+'
    match = re.search(pattern, extracted_content)

    if match:
        number = match.group().strip()
    else:
        number = None  # Set to None if no number is found

    return number

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
    
def random_delay(driver):
    WebDriverWait(driver,random.uniform(1, 5))

def safe_get_text(debug_label, driver, selector):
    """Always returns a list of text values (empty list if none found)"""
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        texts = []
        for elem in elements:
            text = elem.text.strip()
            if text:  # Skip empty text
                texts.append(text)
        return texts  # Always a list
    except Exception as e:
        print(f"Error ({debug_label}): {e}")
        return []  # Return empty list on failure

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

def scrape_product_page(driver, url, position):
    try:
        driver.get(url)
        # Wait for page stability
        WebDriverWait(driver, 5).until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )
        # Extract page number from the URL
        page_number = extract_page_number(url)
        
        # Scrape all fields, handling multiple elements
        initial_data = {
            "A. UNIQUE_ID":datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
            "B. APPAREL_TYPOLOGY": get_elements_text("B",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-last-child(-n+5)"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": "COS",
            "E. WEBSITE_URL": get_base_url(url),
            "F. PRODUCT_URL": url,
            "G. CITY": "Stockholm",
            "H. WEBSITE_COUNTRY": "Master Samuelsgatan 46 A, Stockholm, SE, 10638, SE",
            "I. GENDER": get_elements_text("I",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-child(3)"),
            "J. SEASON": "SUMMER 2025", #get_elements_text(driver,"season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text("K",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-child(7)"),
            "L. PRODUCT_NAME": get_elements_text("L",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.flex.flex-col.items-start.lg\:px-0 > h1"),
            "U. MAIN_BODY_COLOUR": get_elements_text("U",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.pb-12.lg\:order-2.lg\:col-span-4.lg\:mt-12.lg\:border-0.lg\:pl-2.-mx-3\.75.lg\:mx-0.lg\:pr-0 > div.border-b-0\.5.border-faint-divider.px-3\.75.pb-7\.5.lg\:border-0.lg\:px-0.lg\:pb-\[3\.625rem\] > div > h2"),
            "V. LIST_OF_COLORS":get_elements_attribute("V",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.pb-12.lg\:order-2.lg\:col-span-4.lg\:mt-12.lg\:border-0.lg\:pl-2.-mx-3\.75.lg\:mx-0.lg\:pr-0 > div.border-b-0\.5.border-faint-divider.px-3\.75.pb-7\.5.lg\:border-0.lg\:px-0.lg\:pb-\[3\.625rem\] > div > div > ul > li > a > span > img", "alt"),
            "Z. RETAIL_CURRENCY":"USD",# extract_currency_symbol(get_elements_text(driver," div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZA. PRICE": extract_number(get_elements_text("ZA",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.flex.flex-col.items-start.lg\:px-0 > div > span")),
            "ZB. ON_SALE": "Yes",#get_elements_text(driver,"sale-selector"),
            "ZC. SALE_TYPE":"Normal Sale",#get_elements_text(driver,"sale-type-selector"),
            "ZD. SALE_AMOUNT": extract_number(get_elements_text("ZD",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.flex.flex-col.items-start.lg\:px-0 > div > span")),
            "ZF. FEATURED_PRODUCT": "unknown",#get_elements_text(driver,"featured-selector"),
            "ZG. MAIN_IMAGE_URL": get_elements_attribute("ZG",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.flex-1.flex-shrink-0.overflow-auto.lg\:flex-0.lg\:z-30.lg\:col-span-6.lg\:h-\[calc\(100vh_-_var\(--header-height\)\)\].lg\:overflow-visible.lg\:col-start-1.lg\:row-start-1.lg\:-ml-4.lg\:-mr-2.lg\:min-h-full > div > div > div.relative.h-full.overflow-y-clip.theme-slider-light > div.no-scrollbar.h-full.block.overflow-y-auto.lg\:space-y-0\.\[125rem\].space-x-\[0\.125rem\].lg\:space-x-0.lg\:space-y-\[0\.125rem\] > button:nth-child(1) > img", "src"),
            "ZH. IMAGE_META_TAGS": get_elements_attribute("ZH",driver,r"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.flex-1.flex-shrink-0.overflow-auto.lg\:flex-0.lg\:z-30.lg\:col-span-6.lg\:h-\[calc\(100vh_-_var\(--header-height\)\)\].lg\:overflow-visible.lg\:col-start-1.lg\:row-start-1.lg\:-ml-4.lg\:-mr-2.lg\:min-h-full > div > div > div.relative.h-full.overflow-y-clip.theme-slider-light > div.no-scrollbar.h-full.block.overflow-y-auto.lg\:space-y-0\.\[125rem\].space-x-\[0\.125rem\].lg\:space-x-0.lg\:space-y-\[0\.125rem\] > button:nth-child(1) > img", "alt"),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number, # Now using the extracted page number
            "Shirts and Trousers":position
        }
            # Click description button and get remaining elements
        description_data = {} 
        if click_btn(driver,'[data-testid="details-et-description-heading"]'):
                 
                # For "Q. FIT"
                q_fit_parts = []
                q_fit_parts += safe_get_text("Q2", driver, '[data-testid="accordion-item-0"] > div > ul > li:nth-child(1)')
                

                # For "S. PERIPHERALS"
                s_peripherals_parts = []
                s_peripherals_parts += safe_get_text("S1", driver, '[data-testid="accordion-item-0"] > div > ul > li')
                s_peripherals_parts += safe_get_text("S2", driver, '[data-testid="accordion-item-0"] > div > p:not(:last-child):not(:first-child)')
                s_peripherals = ", ".join(s_peripherals_parts) if s_peripherals_parts else "Unknown"


                description_data = {
                "N. PRODUCT_DESCRIPTION": get_elements_text("N",driver,'[data-testid="accordion-item-0"] > div > p:nth-child(1)'),#disclosure-\:rn\:  [data-testid="accordion-item-0"]
                "S. PERIPHERALS": s_peripherals,
                "Y. KNITTING_WOVEN":"unknown",#get_elements_text(driver,""),
            }

        details_data = {}
        if click_btn(driver,'[data-testid="accordion-button-product-details"]'):
    
                #For fit ihany
                q_fit_parts += safe_get_text("Q1", driver, 'dl > dt[data-testid="attribute-item-tour-de-taille"] + dd') # tsy mandeha fa afindra fotsiny
                q_fit = ", ".join(q_fit_parts) if q_fit_parts else "Regular Fit"  # Special fallback

                # For "O. PATTERN"
                o_pattern_parts = []
                o_pattern_parts += safe_get_text("O1", driver, 'dl > dt[data-testid^="attribute-item-longueur"]+ dd')
                o_pattern_parts += safe_get_text("O2", driver, 'dl > dt[data-testid="attribute-item-longueur-du-vetement"] + dd')
                o_pattern_parts += safe_get_text("O3", driver, 'dl > dt[data-testid="attribute-item-longueur-des-manches"] + dd')
                o_pattern = ", ".join(o_pattern_parts) if o_pattern_parts else "Unknown"

                # For "P. STYLE" 
                p_style_parts = []
                p_style_parts += safe_get_text("P1", driver, 'dl > dt[data-testid^="attribute-item-style"] + dd')
                p_style_parts += safe_get_text("P2", driver, 'dl > dt[data-testid="attribute-item-style-de-col"] + dd')
                p_style_parts += safe_get_text("P3", driver, 'dl > dt[data-testid="attribute-item-style-de-vetement"] + dd') 
                p_style = ", ".join(p_style_parts) if p_style_parts else "Unknown"

                details_data = {
                "M. PRODUCT_REFERENCE": get_elements_text("M",driver,'[data-testid="attribute-item-numero-de-produit"] + dd'),
                "Q. FIT": q_fit, #sady hita desc no hita aty 
                "O. PATTERN": o_pattern,         	
                "P. STYLE": p_style,
                "W. FABRIC_COMPOSITION": get_elements_text("W",driver,'data-testid="attribute-item-composition"+ dd'),
                }
                compliance_data = {}
                if click_btn(driver,'span.underline.font_s_regular'):
                    compliance_data = {
                        "X. IDENTIFIED AS SUSTAINABLE":get_elements_text("X",driver,'h5[data-testid="heading"] + ul > li > p') or "Unknown",
                        "T. COUNTRY_OF_ORIGIN": get_elements_text("T",driver,r"#headlessui-dialog-panel-\:rh\: > div > div.pointer-events-auto.flex.w-full.flex-col.overflow-auto.bg-main.md\:\[\&\:first-child\]\:-mr-2\.5.lg\:w-1\/3.xl\:w-\[25\%\].xl\:\[\&\:first-child\]\:w-\[calc\(25\%_\+_1\.25rem\)\] > div.flex.flex-1.flex-col.justify-between.overflow-auto.px-3\.75.pb-4.lg\:px-4 > div > div:nth-child(4) > ul > li > p") or "Unknown"
                    }
                    click_btn(driver,'[data-testid="compliance-details-drawer-close-btn"]')

        care_data = {}
        if click_btn(driver,'[data-testid="accordion-button-product-care-guide"]'):
             care_data = {
                "R. NON_IRON": get_elements_text("R",driver,'[data-testid="accordion-item-2"]> ul > li:nth-last-child(2) > p') or "Unknown",
            }
        supplier_data = {}
        if click_btn(driver,'[data-testid="accordion-button-product-materials-suppliers"]'):
             supplier_data = {
                "ZE. CIEL_TEX_PRODUCT": get_elements_text("ZE",driver,'[data-testid="accordion-item-3"]> div> ul > li > ul > li:nth-child(2) > p ') or "Unknown",
            }  
        click_center_of_screen(driver)

        size_data = {}
        if click_btn(driver,'[data-testid="trouver-ma-taille-heading"]'):
             size_data = {
                "ZI. SIZE_SET": get_elements_text("ZI",driver,r"div > div > div:nth-child(2) > div > div > div > div.max-w-full.overflow-x-auto > table > tbody> tr > td.border-main-chart.px-1.py-4.text-center.font_mini_regular_uc.group-hover\:border-t.group-hover\:border-t-main-chart-hover"),
                "ZJ. SIZE_AVAILABILITY": get_elements_text("ZJ",driver,r"div > div > div:nth-child(2) > div > div > div > div.max-w-full.overflow-x-auto > table > tbody> tr > td.border-main-chart.px-1.py-4.text-center.font_mini_regular_uc.group-hover\:border-t.group-hover\:border-t-main-chart-hover"),
            }

        # Merge dictionaries while maintaining alphabetical order by prefix
        # Step 1: Combine all dictionaries into one
        all_data = {}
        all_data.update(initial_data)
        all_data.update(description_data)
        all_data.update(details_data)
        all_data.update(care_data)
        all_data.update(compliance_data)
        all_data.update(supplier_data)
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
        print(f"Error on {url}: {str(e)} ")
        return None

def get_all_product_links(driver):
    all_product_links = []
    global_position = 0  # Track position across all pages
    #Scraping product 915/2276 (Global position: 915)
    #Error on https://www.cos.com/fr-fr/men/menswear/jeans/regular-fit/product/pillar-jeans-tapered-blue-1052587027: Message: unknown error: net::ERR_INTERNET_DISCONNECTED
    #(Session info: chrome=134.0.6998.89)
   
    # Loop through pages 1 to 89  MISY KNITWEAR AO FA MILA JERENA A PARTIR AN'LE LIENS
    for page in range(0,2): 
        url = f"https://www.cos.com/en_usd/women/shirts.html?qualityName=Cotton&sort=bestMatch&page={page}"
        driver.get(url)
        #driver.implicitly_wait(10)
        click_if_needed(driver, "#onetrust-accept-btn-handler")
        random_delay(driver)
        click_if_needed(driver, '[aria-label="FERMER"]')
        
        # Get all product links with GLOBAL positions
        product_elements = driver.find_elements(By.CSS_SELECTOR, "div.o-product.productTrack> div.description > a")
        for elem in product_elements:
            global_position += 1
            all_product_links.append( (elem.get_attribute("href"), global_position) )
        
        print(f"Page {page}: Collected {len(product_elements)} products (Global position: {global_position})")
        time.sleep(2)
    
    return all_product_links  # List of tuples (url, global_position)

def main():
    driver = setup_chrome_driver()
    all_products = []
    current_batch = []
    batch_size = 90

    save_directory = "C:/Users/Arotiana/Documents/Scrap/Cos/"  

    try:
        # Create the directory if it doesn't exist
        os.makedirs(save_directory, exist_ok=True)
        print(f"Files will be saved in: {save_directory}")

        # Step 1: Get all product links WITH GLOBAL POSITIONS
        print("Collecting product links...")
        product_links = get_all_product_links(driver)  # Returns (url, position) tuples
        print(f"Total product links collected: {len(product_links)}")
        
        # Step 2: Scrape with global positions
        print("Scraping product details...")
        for idx, (url, global_pos) in enumerate(product_links):
            print(f"Scraping product {idx+1}/{len(product_links)} (Global position: {global_pos})")
            product_data = scrape_product_page(driver, url, global_pos)
            
            if product_data:
                current_batch.append(product_data)
                print(f"batched: {url}")

                  # Batch saving
            if len(current_batch) >= batch_size:
                print(f"Batch size reached: {len(current_batch)}")
                df = pd.DataFrame(current_batch)
                file_name = f'Cos_WOMEN_shirts_batch_{(idx//batch_size)+1}.xlsx'
                file_path = os.path.join(save_directory, file_name)  # Save to custom directory
                try:
                    df.to_excel(file_path, index=False)
                    print(f"Saved batch to: {file_path}")
                except Exception as e:
                    print(f"Error saving file {file_path}: {str(e)}")
                    import traceback
                    traceback.print_exc()
                all_products.extend(current_batch)
                current_batch = []
            
            time.sleep(3)
        
        
        # Final save
        if current_batch:
            print(f"Final batch size: {len(current_batch)}")
            df = pd.DataFrame(current_batch)
            file_name = 'products_batch_Cos_final.xlsx'
            file_path = os.path.join(save_directory, file_name)  # Save to custom directory
            try:
                df.to_excel(file_path, index=False)
                print(f"Saved final batch to: {file_path}")
            except Exception as e:
                print(f"Error saving final file {file_path}: {str(e)}")
                import traceback
                traceback.print_exc()
            all_products.extend(current_batch)
        
        # Save all products
        df_all = pd.DataFrame(all_products)
        all_products_file = os.path.join(save_directory, 'cos_WOMEN_shirts.xlsx')  # Save to custom directory
        try:
            df_all.to_excel(all_products_file, index=False)
            print(f"Saved all products to: {all_products_file}")
        except Exception as e:
            print(f"Error saving all products file: {str(e)}")
            import traceback
            traceback.print_exc()
            
    except Exception as e:
        print(f"Main error: {str(e)}")
        import traceback
        traceback.print_exc()
    finally:
        driver.quit()

if __name__ == "__main__":
    main()