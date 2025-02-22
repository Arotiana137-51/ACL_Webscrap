from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
import time
import os

def setup_edge_driver():
    edge_options = Options()
    
    # Profile settings
    user_data_dir = "C:\\Users\\Arotiana\\AppData\\Local\\Microsoft\\Edge\\User Data"
    
    # Verify if the profile directory exists
    if not os.path.exists(user_data_dir):
        print(f"Profile directory not found: {user_data_dir}")
        return None
        
    # Add necessary options
    edge_options.add_argument(f"user-data-dir={user_data_dir}")
    edge_options.add_argument("profile-directory=Default")
    
    # Add stability options
    edge_options.add_argument("--no-sandbox")
    edge_options.add_argument("--disable-dev-shm-usage")
    edge_options.add_argument("--start-maximized")
    edge_options.add_argument("--disable-extensions")  # Temporarily disable extensions
    edge_options.add_experimental_option("excludeSwitches", ["enable-automation"])
    edge_options.add_experimental_option('useAutomationExtension', False)
    
    try:
        # Close any existing Edge sessions
        #os.system("taskkill /f /im msedge.exe")
        #time.sleep(2)  # Wait for processes to close
        
        driver = webdriver.Edge(options=edge_options)
        return driver
    except Exception as e:
        print(f"Error creating driver: {e}")
        return None

def main():
    print("Starting Edge setup...")
    driver = setup_edge_driver()
    
    if not driver:
        print("Failed to create driver")
        return
    
    try:
        print("Opening Google...")
        driver.get("https://www.google.com")
        time.sleep(2)  # Give more time to load
        
        print("Page title:", driver.title)
        
        search_box = driver.find_element(By.NAME, "q")
        search_box.send_keys("Donald trump executive order for Malagasy people")
        search_box.submit()
        
        time.sleep(20)
        
    except Exception as e:
        print(f"Error during execution: {e}")
    
    finally:
        if driver:
            print("Closing browser...")
            driver.quit()

if __name__ == "__main__":
    main()