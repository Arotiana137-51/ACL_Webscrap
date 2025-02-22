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

def setup_edge_driver():
    options = Options()
    options.add_argument("user-data-dir=C:\\Users\\Arotiana\\AppData\\Local\\Microsoft\\Edge\\User Data")  # Use your Edge profile
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    
    return webdriver.Edge(options=options)

def get_elements_text(driver, selector):
    """Get text from all matching elements"""
    try:
        elements = driver.find_elements(By.CSS_SELECTOR, selector)
        return [elem.text.strip() for elem in elements if elem.text.strip()]
    except:
        return []

def click_if_needed(driver, selector):
    """Click element if it exists"""
    try:
        element = driver.find_element(By.CSS_SELECTOR, selector)
        element.click()
        time.sleep(3)  # Wait for any animations/loading
    except:
        pass

def extract_page_number(url):
    """Extract page number from URL, similar to the JS logic"""
    try:
        parsed_url = urlparse(url)
        query_params = parse_qs(parsed_url.query)
        page = query_params.get('page', ['1'])[0]
        return int(page)
    except (ValueError, AttributeError, KeyError):
        return 1  # Default to page 1 if no page number found
    
def scrape_product_page(driver, url, position):
    try:
        driver.get(url)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "your-main-content-selector"))
        )
        
        # Extract page number from the URL
        page_number = extract_page_number(url)
        
        # Click any elements needed to reveal data (example)
        click_if_needed(driver, "show-more-button-selector")
        click_if_needed(driver, "size-chart-selector")
        
        # Scrape all fields, handling multiple elements
        product_data = {
            "A. UNIQUE_ID": time.now().strftime("%Y%m%d%H%M%S") + str(position),
            "B. APPAREL_TYPOLOGY": get_elements_text(" div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:last-of-type"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": get_elements_text("div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.brand-name"),
            "E. WEBSITE_URL": "https://www.ralphlauren.com/",
            "F. PRODUCT_URL": url,
            "G. CITY": "New York",
            "H. WEBSITE_COUNTRY": "USA",
            "I. GENDER": get_elements_text("#product-content > div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:nth-child(2)"),
            "J. SEASON": "unknown", #get_elements_text("season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text("#product-content > div.pdp-infos-cont.pdp-infos-cont-bottom > div.breadcrumb > a:last-of-type"),
            "L. PRODUCT_NAME": get_elements_text("name-selector"),
            "M. PRODUCT_REFERENCE": get_elements_text("reference-selector"),
            "N. PRODUCT_DESCRIPTION": get_elements_text("description-selector"),
            "O. PATTERN": get_elements_text("pattern-selector"),
            "P. STYLE": get_elements_text("style-selector"),
            "Q. FIT": get_elements_text("fit-selector"),
            "R. NON_IRON": get_elements_text("non-iron-selector"),
            "S. PERIPHERALS": get_elements_text("peripherals-selector"),
            "T. COUNTRY_OF_ORIGIN": get_elements_text("origin-selector"),
            "U. MAIN_BODY_COLOUR": get_elements_text("main-color-selector"),
            "V. LIST_OF_COLORS": get_elements_text("colors-selector"),
            "W. FABRIC_COMPOSITION": get_elements_text("fabric-selector"),
            "X. SUSTAINABLE": get_elements_text("sustainable-selector"),
            "Y. KNITTING_WOVEN": get_elements_text("knitting-selector"),
            "Z. RETAIL_CURRENCY": get_elements_text("currency-selector"),
            "ZA. PRICE": get_elements_text("price-selector"),
            "ZB. ON_SALE": get_elements_text("sale-selector"),
            "ZC. SALE_TYPE": get_elements_text("sale-type-selector"),
            "ZD. SALE_AMOUNT": get_elements_text("sale-amount-selector"),
            "ZE. CIEL_TEX_PRODUCT": get_elements_text("ciel-tex-selector"),
            "ZF. FEATURED_PRODUCT": get_elements_text("featured-selector"),
            "ZG. MAIN_IMAGE_URL": get_elements_text("image-selector"),
            "ZH. IMAGE_META_TAGS": get_elements_text("image-meta-selector"),
            "ZI. SIZE_SET": get_elements_text("size-set-selector"),
            "ZJ. SIZE_AVAILABILITY": get_elements_text(),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number, # Now using the extracted page number
            "Shirts and Trousers":position
        }
        
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