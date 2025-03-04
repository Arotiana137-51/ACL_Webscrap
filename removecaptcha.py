from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def press_and_hold_until_disappear(driver, element_xpath):
    """Press and hold the elements until they disappear."""
    try:
        # Wait for the page to load completely and find the p elements
        elements = WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.XPATH, element_xpath))
        )

        # Initialize ActionChains
        actions = ActionChains(driver)

        # Iterate over each element and press & hold
        for element in elements:
            if element.is_displayed():
                actions.click_and_hold(element).perform()
                time.sleep(2)  # Adjust the sleep time as needed
                actions.release(element).perform()

        # Wait until the p elements disappear
        WebDriverWait(driver, 20).until_not(
            EC.presence_of_all_elements_located((By.XPATH, element_xpath))
        )
    
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage
if __name__ == "__main__":
    driver = webdriver.Chrome()
    driver.get("YOUR_WEBPAGE_URL")

    try:
        # Check if the element is displayed before calling the function
        if WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.XPATH, "//p[contains(text(), 'Press & Hold')]"))
        ).is_displayed():
            press_and_hold_until_disappear(driver, "//p[contains(text(), 'Press & Hold')]")
    finally:
        driver.quit()
