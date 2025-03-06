from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.edge.service import Service
#from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
import time
from urllib.parse import urlparse, parse_qs
import re

def setup_edge_driver():
    options = Options()
    options.add_argument("user-data-dir=C:\\Users\\Arotiana\\AppData\\Local\\Microsoft\\Edge\\User Data")  # Use your Edge profile
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--proxy brd.superproxy.io:33335 --proxy-user brd-customer-hl_6bb56026-zone-web_unlocker1:dnbmh3fcfj89")
    
    return webdriver.Edge(options=options)

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
        time.sleep(2)  # Wait for content to load
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
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "your-main-content-selector"))
        )
        
        # Extract page number from the URL
        page_number = extract_page_number(url)
        
        #eto no asiana ny capthca per product sns,....
        # Click any elements needed to reveal data (example)
       # click_if_needed(driver, "show-more-button-selector")
       # click_if_needed(driver, "size-chart-selector")
        
        # Scrape all fields, handling multiple elements
        initial_data = {
            "A. UNIQUE_ID": time.now().strftime("%Y%m%d%H%M%S") + str(position),
            "B. APPAREL_TYPOLOGY": get_elements_text("div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:last-of-type"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": get_elements_text("div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.brand-name"),
            "E. WEBSITE_URL": "https://www.ralphlauren.com/",
            "F. PRODUCT_URL": url,
            "G. CITY": "New York",
            "H. WEBSITE_COUNTRY": "USA",
            "I. GENDER": get_elements_text("#product-content > div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:nth-child(2)"),
            "J. SEASON": "unknown", #get_elements_text("season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text("#product-content > div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:last-of-type"),
            "L. PRODUCT_NAME": get_elements_text(" div.product-favorite-cont > h1"),
            "N. PRODUCT_DESCRIPTION": get_elements_text(" div.pdp-product-details > div.pdp-details-description.js-clamped > div"),
            "R. NON_IRON": "unknown",# get_elements_text("non-iron-selector"),
            "S. PERIPHERALS": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li"),
            "U. MAIN_BODY_COLOUR": get_elements_text(" div.product-variations.pdp-content.dfrefreshcont > ul > li.js-attribute-wrapper.attribute.colorname > div > div.attribute-top-links > span.js-selected-value-wrapper.selected-value.select-attribute.selected-color"),
            "V. LIST_OF_COLORS": get_elements_text(" div.s7viewer-swatches.swiper-container.color-swatches-wrapper > ul > li.variations-attribute.selectable> div > a"),
            "X. IDENTIFIED AS SUSTAINABLE":"unknown",# get_elements_text("sustainable-selector"),
            "Z. RETAIL_CURRENCY": extract_currency_symbol(get_elements_text(" div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZA. PRICE": extract_number(get_elements_text(" div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZB. ON_SALE": "Yes",#get_elements_text("sale-selector"),
            "ZC. SALE_TYPE":"Normal Sale",#get_elements_text("sale-type-selector"),
            "ZD. SALE_AMOUNT": extract_number(get_elements_text(" div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZE. CIEL_TEX_PRODUCT": "unknown",#get_elements_text("ciel-tex-selector"),
            "ZF. FEATURED_PRODUCT": "unknown",#get_elements_text("featured-selector"),
            "ZG. MAIN_IMAGE_URL": get_elements_attribute("#pdpMain > div.pdp-top-cont.js-product-top-content-container > div.js-product-images-section.pdp-left-col > div > div.pdp-media-container.js-pdp-media-container.js-pdp-video-container > div > div.swiper-wrapper > div:nth-child(1) > div > picture > source:nth-child(1)", "src"),
            "ZH. IMAGE_META_TAGS": get_elements_attribute(" div.pdp-top-cont.js-product-top-content-container > div.js-product-images-section.pdp-left-col > div > div.pdp-media-container.js-pdp-media-container.js-pdp-video-container > div > div.swiper-wrapper > div:nth-child(1) > div > picture > img", "alt"),
            "ZI. SIZE_SET": get_elements_text(" li.js-attribute-wrapper.attribute.primarysize.sizing > div > ul > li> a > span > bdi"),
            "ZJ. SIZE_AVAILABILITY": get_elements_text(" li.js-attribute-wrapper.attribute.primarysize.sizing > div > ul > li> a > span > bdi"),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number, # Now using the extracted page number
            "Shirts and Trousers":position
        }
            # Click description button and get remaining elements
        description_data = {} 
        if click_description_button(driver,"#product-content > div.pdp-product-details > ul > li:nth-child(1) > button"):
            description_data = {
                "M. PRODUCT_REFERENCE": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li.style-number.js-extra > span.screen-reader-digits"),
                "O. PATTERN": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(4)"),
            	"P. STYLE": get_elements_text("#app-wrapper > div > main > div.grid-x.margin-top-xs > div > div > div > div.large-offset-1.xlarge-5.xlarge-offset-0.cell.medium-6.large-5 > div.title > h1 > span"),
                "Q. FIT": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(1)"),
                "S. PERIPHERALS": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li"),
                "T. COUNTRY_OF_ORIGIN": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(8)"),
                "W. FABRIC_COMPOSITION": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(6)"),
                "Y. KNITTING_WOVEN": get_elements_text("body > div.rl-toaster.from-right.pdp-flyout.full-height.is-pdp-redesign.r24-form.js-details-flyout.opened > div.rl-toaster-content > div > div > div > div.bullet-list > ul > li:nth-child(4)"),
            }
        
        # Merge both dictionaries while maintaining order
        product_data = {**initial_data}
        for key in sorted(description_data.keys()):
            product_data[key] = description_data[key]
        
        return product_data
        
    except TimeoutException:
        print(f"Timeout on {url} - possible CAPTCHA")
        input("Please solve CAPTCHA and press Enter to continue...")
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
        driver.get("https://www.ralphlauren.com/women-clothing-sweaters?webcat=content-women-clothing-sweaters&ab=en_US_WLP_Slot_2_S1_Image_SHOP")
        
        # Get all product links with their positions
        product_elements = driver.find_elements(By.CSS_SELECTOR, " div.product-name-row.favorite-enabled > div.product-name > a")#product-link-selector
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
        df_all.to_excel('all_products.xlsx', index=False)
            
    except Exception as e:
        print(f"Main error: {str(e)}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()