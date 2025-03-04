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
    options.add_argument("user-data-dir=C:\\Users\\Arotiana\\AppData\\Local\\Microsoft\\Edge\\User Data")
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

def click_description_button(driver):
    """Click the description button and wait for content to load"""
    try:
        description_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "description-button-selector"))
        )
        description_button.click()
        time.sleep(2)  # Wait for content to load
        return True
    except Exception as e:
        print(f"Error clicking description button: {str(e)}")
        return False

def extract_page_number(url):
    """Extract page number from URL"""
    try:
        parsed_url = urlparse(url)
        query_params = parse_qs(parsed_url.query)
        page = query_params.get('page', ['1'])[0]
        return int(page)
    except (ValueError, AttributeError, KeyError):
        return 1

def scrape_product_page(driver, url, position):
    try:
        driver.get(url)
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "your-main-content-selector"))
        )
        
        # Get page number
        page_number = extract_page_number(url)
        
        # First get all elements that don't require description button click
        initial_data = {
            "A. UNIQUE_ID": get_elements_text(driver, "id-selector"),
            "B. APPAREL_TYPOLOGY": get_elements_text(driver, "typology-selector"),
            "C. AGENT_NAME": get_elements_text(driver, "agent-selector"),
            "D. BRAND": get_elements_text(driver, "brand-selector"),
            "E. WEBSITE_URL": "https://add the current url here",
            "F. PRODUCT_URL": url,
            "G. CITY": get_elements_text(driver, "city-selector"),
            "H. WEBSITE_COUNTRY": get_elements_text(driver, "country-selector"),
            "I. GENDER": get_elements_text(driver, "gender-selector"),
            "J. SEASON": get_elements_text(driver, "season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text(driver, "category-selector"),
            "L. PRODUCT_NAME": get_elements_text(driver, "name-selector"),
            "N. PRODUCT_DESCRIPTION": get_elements_text(driver, "description-selector"),
            "R. NON_IRON": get_elements_text(driver, "non-iron-selector"),
            "U. MAIN_BODY_COLOUR": get_elements_text(driver, "main-color-selector"),
            "V. LIST_OF_COLORS": get_elements_text(driver, "colors-selector"),
            "X. SUSTAINABLE": get_elements_text(driver, "sustainable-selector"),
            "Z. RETAIL_CURRENCY": get_elements_text(driver, "currency-selector"),
            "ZA. PRICE": get_elements_text(driver, "price-selector"),
            "ZB. ON_SALE": get_elements_text(driver, "sale-selector"),
            "ZC. SALE_TYPE": get_elements_text(driver, "sale-type-selector"),
            "ZD. SALE_AMOUNT": get_elements_text(driver, "sale-amount-selector"),
            "ZE. CIEL_TEX_PRODUCT": get_elements_text(driver, "ciel-tex-selector"),
            "ZF. FEATURED_PRODUCT": get_elements_text(driver, "featured-selector"),
            "ZG. MAIN_IMAGE_URL": get_elements_text(driver, "image-selector"),
            "ZH. IMAGE_META_TAGS": get_elements_text(driver, "image-meta-selector"),
            "ZI. SIZE_SET": get_elements_text(driver, "size-set-selector"),
            "ZJ. SIZE_AVAILABILITY": get_elements_text(driver, "size-availability-selector"),
            "ZK. POSITION": position,
            "ZL. PAGE_NUMBER": page_number
        }
        
        # Click description button and get remaining elements
        description_data = {}
        if click_description_button(driver):
            description_data = {
                "M. PRODUCT_REFERENCE": get_elements_text(driver, "reference-selector"),
                "O. PATTERN": get_elements_text(driver, "pattern-selector"),
                "P. STYLE": get_elements_text(driver, "style-selector"),
                "Q. FIT": get_elements_text(driver, "fit-selector"),
                "S. PERIPHERALS": get_elements_text(driver, "peripherals-selector"),
                "T. COUNTRY_OF_ORIGIN": get_elements_text(driver, "origin-selector"),
                "W. FABRIC_COMPOSITION": get_elements_text(driver, "fabric-selector"),
                "Y. KNITTING_WOVEN": get_elements_text(driver, "knitting-selector")
            }
        
        # Merge both dictionaries while maintaining order
        product_data = {**initial_data}
        for key in sorted(description_data.keys()):
            product_data[key] = description_data[key]
        
        return product_data
        
    except TimeoutException:
        print(f"Timeout on {url} - possible CAPTCHA")
        input("Please solve CAPTCHA and press Enter to continue...")
        return scrape_product_page(driver, url, position)
    except Exception as e:
        print(f"Error on {url}: {str(e)}")
        return None

def main():
    driver = setup_edge_driver()
    all_products = []
    current_batch = []
    batch_size = 200
    
    try:
        driver.get("your_main_grid_url")
        
        product_elements = driver.find_elements(By.CSS_SELECTOR, "product-link-selector")
        product_links = [(elem.get_attribute("href"), idx + 1) for idx, elem in enumerate(product_elements)]
        
        for url, position in product_links:
            product_data = scrape_product_page(driver, url, position)
            if product_data:
                current_batch.append(product_data)
            
            if len(current_batch) >= batch_size:
                df = pd.DataFrame(current_batch)
                df.to_excel(f'products_batch_{position//batch_size}.xlsx', index=False)
                all_products.extend(current_batch)
                current_batch = []
                
            time.sleep(2)
        
        if current_batch:
            df = pd.DataFrame(current_batch)
            df.to_excel(f'products_batch_final.xlsx', index=False)
            all_products.extend(current_batch)
        
        df_all = pd.DataFrame(all_products)
        df_all.to_excel('all_products.xlsx', index=False)
            
    except Exception as e:
        print(f"Main error: {str(e)}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()