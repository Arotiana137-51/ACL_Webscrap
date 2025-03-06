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

def get_base_url(long_url):
    # Parse the URL into its components
    parsed_url = urlparse(long_url)
    
    # Construct base URL from scheme and netloc components
    base_url = f"{parsed_url.scheme}://{parsed_url.netloc}/"
    return base_url

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
        
        initial_data = {
            "A. UNIQUE_ID": time.now().strftime("%Y%m%d%H%M%S") + str(position),
            "B. APPAREL_TYPOLOGY": get_elements_text("",driver,"div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:last-of-type"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": get_elements_text("",driver,"div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.brand-name"),
            "E. WEBSITE_URL": get_base_url(url),
            "F. PRODUCT_URL": url,
            "G. CITY": "New York",
            "H. WEBSITE_COUNTRY": "USA",
            "I. GENDER": get_elements_text("",driver,"#product-content > div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:nth-child(2)"),
            "J. SEASON": "unknown", #get_elements_text("",driver,"season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text("",driver,"#product-content > div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:last-of-type"),
            "L. PRODUCT_NAME": get_elements_text("",driver," div.product-favorite-cont > h1"),
            "N. PRODUCT_DESCRIPTION": get_elements_text("",driver," div.pdp-product-details > div.pdp-details-description.js-clamped > div"),
            "R. NON_IRON": "unknown",# get_elements_text("",driver,"non-iron-selector"),
            "S. PERIPHERALS": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li"),
            "U. MAIN_BODY_COLOUR": get_elements_text("",driver," div.product-variations.pdp-content.dfrefreshcont > ul > li.js-attribute-wrapper.attribute.colorname > div > div.attribute-top-links > span.js-selected-value-wrapper.selected-value.select-attribute.selected-color"),
            "V. LIST_OF_COLORS": get_elements_text("",driver," div.s7viewer-swatches.swiper-container.color-swatches-wrapper > ul > li.variations-attribute.selectable> div > a"),
            "X. IDENTIFIED AS SUSTAINABLE":"unknown",# get_elements_text("",driver,"sustainable-selector"),
            "Z. RETAIL_CURRENCY": extract_currency_symbol(get_elements_text("",driver," div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZA. PRICE": extract_number(get_elements_text("",driver," div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZB. ON_SALE": "Yes",#get_elements_text("",driver,"sale-selector"),
            "ZC. SALE_TYPE":"Normal Sale",#get_elements_text("",driver,"sale-type-selector"),
            "ZD. SALE_AMOUNT": extract_number(get_elements_text("",driver," div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZE. CIEL_TEX_PRODUCT": "unknown",#get_elements_text("",driver,"ciel-tex-selector"),
            "ZF. FEATURED_PRODUCT": "unknown",#get_elements_text("",driver,"featured-selector"),
            "ZG. MAIN_IMAGE_URL": get_elements_attribute("#pdpMain > div.pdp-top-cont.js-product-top-content-container > div.js-product-images-section.pdp-left-col > div > div.pdp-media-container.js-pdp-media-container.js-pdp-video-container > div > div.swiper-wrapper > div:nth-child(1) > div > picture > source:nth-child(1)", "src"),
            "ZH. IMAGE_META_TAGS": get_elements_attribute(" div.pdp-top-cont.js-product-top-content-container > div.js-product-images-section.pdp-left-col > div > div.pdp-media-container.js-pdp-media-container.js-pdp-video-container > div > div.swiper-wrapper > div:nth-child(1) > div > picture > img", "alt"),
            "ZI. SIZE_SET": get_elements_text("",driver," li.js-attribute-wrapper.attribute.primarysize.sizing > div > ul > li> a > span > bdi"),
            "ZJ. SIZE_AVAILABILITY": get_elements_text("",driver," li.js-attribute-wrapper.attribute.primarysize.sizing > div > ul > li> a > span > bdi"),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number, # Now using the extracted page number
            "Shirts and Trousers":position
        }
            # Click description button and get remaining elements
        description_data = {} 
        if click_btn(driver,"#product-content > div.pdp-product-details > ul > li:nth-child(1) > button"):
            description_data = {
                "M. PRODUCT_REFERENCE": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li.style-number.js-extra > span.screen-reader-digits"),
                "O. PATTERN": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(4)"),
            	"P. STYLE": get_elements_text("",driver,"#app-wrapper > div > main > div.grid-x.margin-top-xs > div > div > div > div.large-offset-1.xlarge-5.xlarge-offset-0.cell.medium-6.large-5 > div.title > h1 > span"),
                "Q. FIT": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(1)"),
                "S. PERIPHERALS": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li"),
                "T. COUNTRY_OF_ORIGIN": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(8)"),
                "W. FABRIC_COMPOSITION": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(6)"),
                "Y. KNITTING_WOVEN": get_elements_text("",driver,"body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(4)"),
            }
        
      

        # Merge dictionaries while maintaining alphabetical order by prefix
        # Step 1: Combine all dictionaries into one
        all_data = {}
        all_data.update(initial_data)
        all_data.update(description_data )


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
    product_link_selector = "div.product-name-row.favorite-enabled > div.product-name > a"
    driver = setup_chrome_driver()
    all_product_links = []
    seen_hrefs = set()
    idx = 1
    retries = 0
    max_retries = 8  # Reduced from 20 since we're more efficient
    scroll_attempts = 0

    try:
        main_url = check_redirection("https://www.ralphlauren.com/women-clothing-sweaters?webcat=content-women-clothing-sweaters&ab=en_US_WLP_Slot_2_S1_Image_SHOP")
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