import  undetected_chromedriver as uc
#from selenium import webdriver
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
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")
    #options.add_experimental_option("excludeSwitches", ["enable-automation"])
    #options.add_experimental_option('useAutomationExtension', False)
    #options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/83.0.4103.53 Safari/537.36')

    # Initialize undetected Chrome driver
    driver = uc.Chrome(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

# def rotate_user_agent(driver):
#     user_agents = [
#         "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3",
#         "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:53.0) Gecko/20100101 Firefox/53.0",
#         "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_12_6) AppleWebKit/603.3.8 (KHTML, like Gecko) Version/10.1.2 Safari/603.3.8"
#     ]
#     user_agent = random.choice(user_agents)
#     driver.execute_cdp_cmd('Network.setUserAgentOverride', {"userAgent": user_agent})

def random_delay():
    time.sleep(random.uniform(1, 3))


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
                return [elem.text.strip() for elem in elements if elem.text.strip()]
        else:
            print(f"{letter}: No elements found with the selector: {selector}")
            raise NoSuchElementException(f"No elements found with the selector: {selector}")
          
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
            "B. APPAREL_TYPOLOGY": get_elements_text("B",driver," div.ProductHeroGridContainer__breadcrumbs.padt4.padb4 > ul > li > a"),
            "C. AGENT_NAME": "Aro",
            "D. BRAND": get_elements_text("D",driver,"div.ProductHeading > div.ProductHeading__brand--container > a"),
            "E. WEBSITE_URL": "https://huckberry.com/store",
            "F. PRODUCT_URL": url,
            "G. CITY": "220 Industrial Blvd, Austin, Texas 78745",
            "H. WEBSITE_COUNTRY": "USA",
            "I. GENDER":"Men", # get_elements_text("I",driver,"#main-block-wrapper > div.blocks.flex.flex-col.pb-\[--pdp-sticky-details-height\].lg\:pb-0.pt-\[var\(--header-height\)\].lg\:pt-0 > div.overflow-auto.lg\:grid.lg\:grid-cols-12.lg\:gap-x-4.lg\:overflow-visible.lg\:px-4 > div.lg\:min-h-auto.lg\:static.lg\:grid-cols-6.lg\:overflow-visible.lg\:bottom-auto.lg\:top-0.lg\:col-span-6.lg\:col-start-7.lg\:row-start-1 > div > div.mb-auto.flex.w-full.flex-col.justify-center.px-3\.75.lg\:grid.lg\:grid-cols-6.lg\:gap-4.lg\:gap-y-0.lg\:px-0.lg\:pt-9 > div.col-span-full.flex.flex-col.pt-\[1\.3125rem\].lg\:col-span-4.lg\:gap-8.lg\:pl-2.-mx-3\.75.px-3\.75.lg\:mx-0.lg\:pr-0 > div.relative.ml-\[-1\.125rem\].w-full.hidden.font_small_xs_regular.lg\:mb-\[max\(5rem\,5\.3vw\)\].lg\:block > div > a:nth-child(3)"),
            "J. SEASON": "unknown", #get_elements_text(driver,"season-selector"),
            "K. PRODUCT_CATEGORY": get_elements_text("K",driver," div.ProductHeroGridContainer__breadcrumbs.padt4.padb4 > ul > li:nth-child(5) > a"),
            "L. PRODUCT_NAME": get_elements_text("L",driver,"div.ProductHeading > h1 > span"),
            "M. PRODUCT_REFERENCE":get_last_path_element(url),
            "N. PRODUCT_DESCRIPTION": get_elements_text("N",driver,"div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.ProductHeroDescription.padt8.padb6 > div > div > div > p"),
            "O. PATTERN": get_elements_text("O",driver,' div.ProductHeroGridContainer__breadcrumbs.padt4.padb4 > ul > li:nth-child(5) > a'),
            "P. STYLE": get_elements_text("P",driver,'div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.ProductHeroDescription.padt8.padb6 > h3'),
            "S. PERIPHERALS": get_elements_text("S",driver,'#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(1) > div.AccordionItem__items > div > div > div > ul > li'),
            "U. MAIN_BODY_COLOUR": get_elements_text("U",driver,"div.ProductColors > p > span.fw--400"),
            "V. LIST_OF_COLORS":hover_and_extract_text("V",driver,"div.ProductColors > div > div> a", "div.ProductColors > p > span.fw--400"),
            "X. IDENTIFIED AS SUSTAINABLE":"unknown", #get_elements_text("X",driver,"#product-materials-suppliers > div > div.mb-3.grid.grid-cols-2.gap-x-7\.5.gap-y-4.font_s_regular > p"),
            "Z. RETAIL_CURRENCY":"$",# extract_currency_symbol(get_elements_text(driver," div.pdp-infos-cont.pdp-infos-cont-bottom > div.pdp-infos > div.product-price > span")),
            "ZA. PRICE": extract_number(get_elements_text("ZA",driver," div.ProductHeading > div.ProductHeading__price.fg--xsmall")),
            "ZB. ON_SALE": "Yes",#get_elements_text(driver,"sale-selector"),
            "ZC. SALE_TYPE":"Normal Sale",#get_elements_text(driver,"sale-type-selector"),
            "ZD. SALE_AMOUNT": extract_number(get_elements_text("ZD",driver," div.ProductHeading > div.ProductHeading__price.fg--xsmall")),
            "ZF. FEATURED_PRODUCT": get_elements_text("ZF.Not mendatory",driver,"div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__media.padr14--lg > div.ProductHeroMedia__desktop > div.ProductHeroMedia__image.pr.ProductHeroMedia__image--square > div.ProductDealTags.mb2 > div"),
            "ZG. MAIN_IMAGE_URL": get_elements_attribute("ZG",driver,"div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__media.padr14--lg > div.ProductHeroMedia__desktop > div.ProductHeroMedia__image.pr.ProductHeroMedia__image--square > div.ProductHeroMedia__image--wrapper > img", "src"),
            "ZH. IMAGE_META_TAGS": get_elements_attribute("ZH",driver,"div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__media.padr14--lg > div.ProductHeroMedia__desktop > div.ProductHeroMedia__image.pr.ProductHeroMedia__image--square > div.ProductHeroMedia__image--wrapper > img", "alt"),
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
        if click_btn(driver,'#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(2)'):
            materials_data = {
                "W. FABRIC_COMPOSITION": get_elements_text("W",driver,'#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(2) > div.AccordionItem__items > div > div > div > ul > li:nth-child(1)'),
                "T. COUNTRY_OF_ORIGIN": get_elements_text("T",driver,"#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(2) > div.AccordionItem__items > div > div > div > ul > li:nth-child(2)"),
            }

        size_data = {}
        if click_btn(driver,'#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(3)'):
            size_data = {
                "Q. FIT": get_elements_text("Q",driver,'#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(3) > div.AccordionItem__items > div > div > div > ul:nth-child(1) > li:nth-child(1)'),
                "ZI. SIZE_SET": "Size chart containing XS,S,M,L,XL,XXL",#get_elements_text("ZI",driver,"div > div > div:nth-child(2) > div > div > div > div.max-w-full.overflow-x-auto > table > tbody> tr > td.border-main-chart.px-1.py-4.text-center.font_mini_regular_uc.group-hover\:border-t.group-hover\:border-t-main-chart-hover"),
                "ZJ. SIZE_AVAILABILITY":"Size chart containing XS,S,M,L,XL,XXL",# get_elements_text("ZJ",driver,"div > div > div:nth-child(2) > div > div > div > div.max-w-full.overflow-x-auto > table > tbody> tr > td.border-main-chart.px-1.py-4.text-center.font_mini_regular_uc.group-hover\:border-t.group-hover\:border-t-main-chart-hover"),
            }
        care_data = {}
        if click_btn(driver,'#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(4)'):
            care_data = {
                "R. NON_IRON": get_elements_text("R",driver,"#ProductPageHero > div > div.ProductHeroGridContainer > div.ProductHeroGridContainer__details > div > div.Accordion > div:nth-child(4) > div.AccordionItem__items > div > div > div > ul > li"),
            }

        # Merge dictionaries while maintaining alphabetical order by prefix
        # Step 1: Combine all dictionaries into one
        all_data = {}
        all_data.update(initial_data)
       # all_data.update(features_data)
        all_data.update(materials_data)
        all_data.update(size_data)
        all_data.update(care_data)

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
    product_name= "Huckberry_Men_Shirts"
    driver = setup_chrome_driver()
    all_products = []
    current_batch = []
    batch_size = 50 # Save to Excel every n products
    
    try:
        # Get main product grid page
        main_url=check_redirection("https://huckberry.com/store/t/category/clothing/shirts")
        driver.get(main_url)
       # rotate_user_agent(driver)
        random_delay()
        #Close_Ad
        click_if_needed(driver,"#App > div.SoftGate__modal.SoftGate__modal--short > span > button")
        #Close_survey
        click_if_needed(driver,"#survicate-box > div > div.sv--background-main.sv-bg-bw.sv__micro-theme.sv__survey.sv__survey--button_next.sv__position-center > div.sv--color-question.sv__micro-top-bar.sv__micro-top-bar--flat.sv__button_next__micro-top-bar > div.sv__micro-top-bar__row > div > button")
        #Load_more
        click_if_needed(driver, "#body-block-0 > div > div > div.Pagination > div > div > div > div > button")
        
        # Get all product links with their positions
        product_elements = driver.find_elements(By.CSS_SELECTOR, "#productLink")#product-link-selector
        product_links = [(elem.get_attribute("href"), idx + 1) for idx, elem in enumerate(product_elements)]
        
        for url, position in product_links:
            click_if_needed(driver,"#App > div.SoftGate__modal.SoftGate__modal--short > span > button")
            random_delay()
            click_if_needed(driver,"#App > div.SoftGate__modal.SoftGate__modal--short > span > button")
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
        df_all.to_excel('{product_name}.xlsx', index=False)
            
    except Exception as e:
        print(f"Main error: {str(e)}")
        
    finally:
        driver.quit()

if __name__ == "__main__":
    main()