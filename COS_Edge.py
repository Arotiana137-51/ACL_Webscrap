from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
from urllib.parse import urlparse, parse_qs
import re
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options 
from datetime import datetime


def setup_edge_driver():
    service = Service(r'C:\\Users\\Arotiana\\Documents\\edgedriver_win64\\msedgedriver.exe')
    options = Options()
    #options.add_argument("user-data-dir=C:\\Users\\Arotiana\\AppData\\Local\\Microsoft\\Edge\\User Data")
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    #options.add_argument("--disable-blink-features=AutomationControlled")
    
    # Add these options to make detection harder
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    
    driver = webdriver.Edge(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver  


def get_elements_text(driver, selector):
    """Get text from all matching elements"""
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        return [elem.text.strip() for elem in elements if elem.text.strip()]
    except:
        return []
    
def click_description_button(driver, selector):
    """Click the description button and wait for content to load"""
    try:
        description_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, selector))
        )
        description_button.click()
        time.sleep()  # Wait for content to load
        return True
    except Exception as e:
        print(f"Error clicking description button: {str(e)}")
        return False

def get_elements_attribute(driver, selector, attribute_name):
    """Get attribute values from all matching elements"""
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        return [elem.get_attribute(attribute_name).strip() for elem in elements if elem.get_attribute(attribute_name)]
    except:
        return 

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
    
def scrape_product_page(driver, url, position):
    try:
        driver.get(url)
        # WebDriverWait(driver, 5).until(
        #     EC.presence_of_element_located((By.CSS_SELECTOR, "your-main-content-selector"))
        # )
        
        # Extract page number from the URL
        page_number = extract_page_number(url)
        
        #eto no asiana ny capthca per product sns,....
        # Click any elements needed to reveal data (example)
       # click_if_needed(driver, "show-more-button-selector")
       # click_if_needed(driver, "size-chart-selector")
        
        # Scrape all fields, handling multiple elements
        product_data= {
            "A. UNIQUE_ID": datetime.now(),
            "B. APPAREL_TYPOLOGY": get_elements_text(driver,"#main-content-section > div.pdpbreadcrumb > nav > ol > li:nth-last-child(-n+3)"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": "COS", #get_elements_text("div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.brand-name"),
            "E. WEBSITE_URL": "https://www.cos.com/en_usd",
            "F. PRODUCT_URL": url,
            "G. CITY": "London",
            "H. WEBSITE_COUNTRY": "United Kingdom",
            "I. GENDER": "Women", #get_elements_text("#product-content > div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:nth-child(2)"),
            "J. SEASON": "unknown", #get_elements_text("season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text(driver,"#main-content-section > div.pdpbreadcrumb > nav > ol > li:last-child"),
            "L. PRODUCT_NAME": get_elements_text(driver," h1.a-heading-1"),
            "M. PRODUCT_REFERENCE": get_elements_text(driver,"div.float-on-mobile > div.description-wrapper > div.o-accordion> div.accordion-content > div > div > span.article-number"),
            "N. PRODUCT_DESCRIPTION": get_elements_text(" #productInformation > form > div.float-on-mobile > div.description-wrapper > div.o-accordion > div.accordion-content > div > div.details-text.description-text > p:nth-child(1)"),
            "O. PATTERN": get_elements_text(driver,"#productInformation > form > div.float-on-mobile > div.description-wrapper > div.o-accordion > div.accordion-content > div > div > p:nth-child(4)"),
            "P. STYLE": get_elements_text(driver,"h1.a-heading-1"),
            "Q. FIT": get_elements_text(driver,"#perceivedSize > div > label > span:nth-child(2)"),
            "R. NON_IRON": get_elements_text(driver,"div.o-accordion > div.accordion-content > div > div > p.a-paragraph > span.careInstructions"),# get_elements_text("non-iron-selector"),
            "S. PERIPHERALS": get_elements_text(driver,"#productInformation > form > div.float-on-mobile > div.description-wrapper > div.o-accordion.is-visible > div.accordion-content > div > div > p:nth-last-child(-n+6)"),
            "T. COUNTRY_OF_ORIGIN": get_elements_text(driver,"#productInformation > form > div.float-on-mobile > div.description-wrapper > div.o-accordion > div.accordion-content > div > div > span.imported"),        
            "U. MAIN_BODY_COLOUR": get_elements_text(driver,"button.a-button-nostyle.placeholder"),
            "V. LIST_OF_COLORS": get_elements_attribute(driver,"li.a-option.color-list","data-value"),
            "W. FABRIC_COMPOSITION": get_elements_text(driver,"#productInformation > form > div.float-on-mobile > div.description-wrapper > div.o-accordion > div.accordion-content > div > div > p.a-paragraph > span.productCompositionSpan"),            
            "X. IDENTIFIED AS SUSTAINABLE":"unknown",# get_elements_text("sustainable-selector"),
            "Y. KNITTING_WOVEN": "unknown",#get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(4)"),
            "Z. RETAIL_CURRENCY": extract_currency_symbol(get_elements_text(driver,"#product-price > div > span")),
            "ZA. PRICE": extract_number(get_elements_text(driver,"#product-price > div > span")),
            "ZB. ON_SALE": "Yes",#get_elements_text("sale-selector"),
            "ZC. SALE_TYPE":"Normal Sale",#get_elements_text("sale-type-selector"),
            "ZD. SALE_AMOUNT": extract_number(get_elements_text(driver,"#product-price > div > span")),
            "ZE. CIEL_TEX_PRODUCT": "unknown",#get_elements_text("ciel-tex-selector"),
            "ZF. FEATURED_PRODUCT": "unknown",#get_elements_text("featured-selector"),
            "ZG. MAIN_IMAGE_URL": get_elements_attribute(driver,"#mainImageList > li:nth-child(1) > div > picture>img.a-image.default-image", "data-zoom-src"),
            "ZH. IMAGE_META_TAGS": get_elements_attribute(driver,"#mainImageList > li:nth-child(1) > div > picture>img.a-image.default-image", "alt"),
            "ZI. SIZE_SET": get_elements_text(driver," #sizes > div > button.size-options"),
            "ZJ. SIZE_AVAILABILITY": get_elements_text(driver," #sizes > div > button.size-options.in-stock"),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number, # Now using the extracted page number
            "Shirts and Trousers":position
            }
        
        return product_data
        
    except TimeoutException:
        print(f"Timeout on {url} - possible CAPTCHA")
        input("TIMEOUR OR CAPTCHA - Please solve CAPTCHA and press Enter to continue...")
        return scrape_product_page(driver, url, position)  # Retry after CAPTCHA
    except Exception as e:
        print(f"Error on {url}: {str(e)}")
        return None

def main():
    driver = setup_edge_driver()
    all_products = []
    current_batch = []
    batch_size = 200 # Save to Excel every 200 products
    
    try:
        # Get main product grid page
        driver.get("https://www.cos.com/en_usd/women/shirts.html?qualityName=Cotton&sort=bestMatch")
        
        # Get all product links with their positions
        product_elements = driver.find_elements(By.CSS_SELECTOR, " div.o-product.productTrack> div.description > a")#product-link-selector
        product_links = [(elem.get_attribute("href"), idx + 1) for idx, elem in enumerate(product_elements)]
        
        for url, position in product_links:
            product_data = scrape_product_page(driver, url, position)
            if product_data:
                current_batch.append(product_data)
            
            # Save batch to Excel and clear memory
            if len(current_batch) >= batch_size:
                df = pd.DataFrame(current_batch)
                df.to_excel(f'products_batch_{position//batch_size}.xlsx', index=False)
                all_products.extend(current_batch)
                current_batch = []
                
            time.sleep(2)  # Polite delay between requests
        
        # Save any remaining products
        if current_batch:
            df = pd.DataFrame(current_batch)
            df.to_excel(f'products_batch_final.xlsx', index=False)
            all_products.extend(current_batch)
        
        # Save complete dataset
        df_all = pd.DataFrame(all_products)
        df_all.to_excel('cos_women_shirts.xlsx', index=False)
            
    except Exception as e:
        print(f"Main error: {str(e)}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()