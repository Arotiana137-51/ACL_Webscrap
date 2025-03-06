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
import requests
from urllib.parse import urlparse, parse_qs
import re
from selenium import webdriver
from datetime import datetime


def setup_chrome_driver():
    service = Service(r'C:\\Users\\Arotiana\\Documents\\CromeDriver\\chromedriver-win64\\chromedriver.exe')
    options = Options()
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    #options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Add these options to make detection harder
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Chrome(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver  

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
            print(f"{letter}: No elements found with the selector: {selector}")
            raise NoSuchElementException(f"No elements found with the selector: {selector}")
    except Exception as e:
        return str(e)

    
def click_description_button(driver, selector):
    """Click the description button and wait for content to load"""
    try:
        description_button = WebDriverWait(driver, 10).until(
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
    


def scrape_product_page(driver, url, position):
    try:
        driver.get(url)
        WebDriverWait(driver,3)
        
        # Extract page number from the URL
        page_number = extract_page_number(url)
        
        # Scrape all fields, handling multiple elements
        initial_data = {
            "A. UNIQUE_ID":datetime.now().strftime("%Y-%m-%d-%H-%M-%S"),
            "B. APPAREL_TYPOLOGY": get_elements_text("B",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-last-child(-n+5)"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": "COS",
            "E. WEBSITE_URL": "https://www.cos.com/en_usd",
            "F. PRODUCT_URL": url,
            "G. CITY": "New York",
            "H. WEBSITE_COUNTRY": "USA",
            "I. GENDER": get_elements_text("I",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-child(3)"),
            "J. SEASON": "unknown", #get_elements_text(driver,"season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text("K",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-child(7)"),
            "L. PRODUCT_NAME": get_elements_text("L",driver," #main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.flex.flex-col.items-start.lg\:px-0 > h1"),
            "U. MAIN_BODY_COLOUR": get_elements_text("U",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.pb-12.lg\:order-2.lg\:col-span-4.lg\:mt-12.lg\:border-0.lg\:pl-2.-mx-3\.75.lg\:mx-0.lg\:pr-0 > div.border-b-0\.5.border-faint-divider.px-3\.75.pb-7\.5.lg\:border-0.lg\:px-0.lg\:pb-\[3\.625rem\] > div > h2"),
            "V. LIST_OF_COLORS":get_elements_attribute("V",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.pb-12.lg\:order-2.lg\:col-span-4.lg\:mt-12.lg\:border-0.lg\:pl-2.-mx-3\.75.lg\:mx-0.lg\:pr-0 > div.border-b-0\.5.border-faint-divider.px-3\.75.pb-7\.5.lg\:border-0.lg\:px-0.lg\:pb-\[3\.625rem\] > div > div > ul > li > a > span > img", "alt"),
            "X. IDENTIFIED AS SUSTAINABLE":"unknown", #get_elements_text("X",driver,"#product-materials-suppliers > div > div.mb-3.grid.grid-cols-2.gap-x-7\.5.gap-y-4.font_s_regular > p"),
            "Z. RETAIL_CURRENCY":"USD",# extract_currency_symbol(get_elements_text(driver," div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZA. PRICE": extract_number(get_elements_text("ZA",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.flex.flex-col.items-start.lg\:px-0 > div > span")),
            "ZB. ON_SALE": "Yes",#get_elements_text(driver,"sale-selector"),
            "ZC. SALE_TYPE":"Normal Sale",#get_elements_text(driver,"sale-type-selector"),
            "ZD. SALE_AMOUNT": extract_number(get_elements_text("ZD",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.flex.flex-col.items-start.lg\:px-0 > div > span")),
            "ZF. FEATURED_PRODUCT": "unknown",#get_elements_text(driver,"featured-selector"),
            "ZG. MAIN_IMAGE_URL": get_elements_attribute("ZG",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.flex-1.flex-shrink-0.overflow-auto.lg\:flex-0.lg\:z-30.lg\:col-span-6.lg\:h-\[calc\(100vh_-_var\(--header-height\)\)\].lg\:overflow-visible.lg\:col-start-1.lg\:row-start-1.lg\:-ml-4.lg\:-mr-2.lg\:min-h-full > div > div > div.relative.h-full.overflow-y-clip.theme-slider-light > div.no-scrollbar.h-full.block.overflow-y-auto.lg\:space-y-0\.\[125rem\].space-x-\[0\.125rem\].lg\:space-x-0.lg\:space-y-\[0\.125rem\] > button:nth-child(1) > img", "src"),
            "ZH. IMAGE_META_TAGS": get_elements_attribute("ZH",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.flex-1.flex-shrink-0.overflow-auto.lg\:flex-0.lg\:z-30.lg\:col-span-6.lg\:h-\[calc\(100vh_-_var\(--header-height\)\)\].lg\:overflow-visible.lg\:col-start-1.lg\:row-start-1.lg\:-ml-4.lg\:-mr-2.lg\:min-h-full > div > div > div.relative.h-full.overflow-y-clip.theme-slider-light > div.no-scrollbar.h-full.block.overflow-y-auto.lg\:space-y-0\.\[125rem\].space-x-\[0\.125rem\].lg\:space-x-0.lg\:space-y-\[0\.125rem\] > button:nth-child(1) > img", "alt"),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number, # Now using the extracted page number
            "Shirts and Trousers":position
        }
            # Click description button and get remaining elements
        description_data = {} 
        if click_description_button(driver,'[data-testid="product-info-navigation-details-description-button"]'):
            description_data = {
                "N. PRODUCT_DESCRIPTION": get_elements_text("N",driver,"#product-description > div > div > p:nth-child(1)"),
                "M. PRODUCT_REFERENCE": get_elements_text("M",driver,'[data-testid="attribute-item-product-no"] + dd,#product-details > div > dl > dd:last-child'),
                "O. PATTERN": get_elements_text("O",driver,'[data-testid="attribute-item-sleeve-length"] + dd'),
            	"P. STYLE": get_elements_text("P",driver,'[data-testid="attribute-item-clothing-style"] + dd, [data-testid="attribute-item-collar-style"] + dd'),
                "Q. FIT": get_elements_text("Q",driver,'[data-testid="attribute-item-fit"] + dd'),
                "S. PERIPHERALS": get_elements_text("S",driver,'div[data-testid="accordion-item-0"] ul li ,#product-description > div > div > p:nth-child(n):not(:first-child):not(:last-child)'),
                "T. COUNTRY_OF_ORIGIN": get_elements_text("T",driver,"#product-materials-suppliers > div > div > ul > li > ul > li:nth-of-type(1) > p"),
                "W. FABRIC_COMPOSITION": get_elements_text("W",driver,'[data-testid="attribute-item-composition"] + dd'),
                "R. NON_IRON": get_elements_text("R",driver,"#product-care-guide > div > ul > li > p"),
                "ZE. CIEL_TEX_PRODUCT": get_elements_text("ZE",driver,"#product-materials-suppliers > div > div > ul > li > ul > li:nth-of-type(2) > p"),
                "Y. KNITTING_WOVEN":"unknown",#get_elements_text(driver,""),
            }
        click_center_of_screen(driver)

        size_data = {}
        if click_description_button(driver,'[data-testid="product-info-navigation-find-my-size-button"]'):
            size_data = {
                "ZI. SIZE_SET": get_elements_text("ZI",driver,"div > div > div:nth-child(2) > div > div > div > div.max-w-full.overflow-x-auto > table > tbody> tr > td.border-main-chart.px-1.py-4.text-center.font_mini_regular_uc.group-hover\:border-t.group-hover\:border-t-main-chart-hover"),
                "ZJ. SIZE_AVAILABILITY": get_elements_text("ZJ",driver,"div > div > div:nth-child(2) > div > div > div > div.max-w-full.overflow-x-auto > table > tbody> tr > td.border-main-chart.px-1.py-4.text-center.font_mini_regular_uc.group-hover\:border-t.group-hover\:border-t-main-chart-hover"),
            }

        # Merge dictionaries while maintaining alphabetical order by prefix
        # Step 1: Combine all dictionaries into one
        all_data = {}
        all_data.update(initial_data)
        all_data.update(description_data)
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
    driver = setup_chrome_driver()
    all_products = []
    current_batch = []
    batch_size = 50 # Save to Excel every n products
    
    try:
        # Get main product grid page
        main_url=check_redirection("https://www.cos.com/fr-fr/women/view-all")
        driver.get(main_url)
        # Get all product links with their positions
        product_elements = driver.find_elements(By.CSS_SELECTOR, "div.o-product.productTrack> div.description > a")#product-link-selector
        product_links = [(elem.get_attribute("href"), idx + 1) for idx, elem in enumerate(product_elements)]
        
        for url, position in product_links:
            product_data = scrape_product_page(driver, url, position)
            print(url)
            if product_data:
                current_batch.append(product_data)
            
            # Save batch to Excel and clear memory
            if len(current_batch) >= batch_size:
                df = pd.DataFrame(current_batch)
                df.to_excel(f'products_batch_{position//batch_size}.xlsx', index=False)
                all_products.extend(current_batch)
                current_batch = []
                
            time.sleep(3)  # Polite delay between requests
        
        # Save any remaining products
        if current_batch:
            df = pd.DataFrame(current_batch)
            df.to_excel(f'products_batch_final.xlsx', index=False)
            all_products.extend(current_batch)
        
        # Save complete dataset
        df_all = pd.DataFrame(all_products)
        df_all.to_excel('cos_women.xlsx', index=False)
            
    except Exception as e:
        print(f"Main error: {str(e)}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()