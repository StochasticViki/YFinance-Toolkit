from selenium import webdriver # This is Chrome
from selenium.webdriver.chrome.service import Service # Required along with chrome for background services which run in taskmanager

from selenium.webdriver.common.by import By # Tool for searching/querying similar to find_all in BS4 

from selenium.webdriver.common.keys import Keys # This enables to write (like a keyboard)

from selenium.webdriver.support.ui import WebDriverWait # This waits for JS contents to load
from selenium.webdriver.support import expected_conditions as EC # This is used with WebDriverWait to check if elements are loaded
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, NoSuchElementException
from webdriver_manager.chrome import ChromeDriverManager
import re
import os
import time
import sys
from bs4 import BeautifulSoup

chromedriver_path = r"C:\Users\vikig\OneDrive\Documents\Python\Beta\BetaGen2.0\chromedriver\chromedriver.exe"
download_path = r"C:\Users\vikig\OneDrive\Documents\Python\EDA\downloads"


def wait_for_downloads(download_path, company_dict, timeout=120):
    """
    Wait for Chrome downloads to complete with improved detection and timeout handling.
    """
    # Get initial count of files in the download directory
    initial_files = set(os.listdir(download_path))
    
    # Wait for .crdownload or .tmp files to appear (download starts)
    start_time = time.time()
    download_started = False
    
    print("Waiting for download to start...")
    while time.time() - start_time < 30:  # Wait up to 30 seconds for download to start
        current_files = set(os.listdir(download_path))
        new_files = current_files - initial_files
        
        # Check for temporary download files
        if any(fname.endswith('.crdownload') or fname.endswith('.tmp') for fname in new_files):
            download_started = True
            print(f"Download started...")
            break
        
        # Also check if download completed instantly
        if new_files:
            download_started = True
            print(f"Download may have completed instantly...")
            break
        
        time.sleep(1)
    
    if not download_started:
        print("Warning: No download appears to have started")
        return False
    
    # Now wait for all .crdownload and .tmp files to disappear
    print("Waiting for download to complete...")
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(download_path))
        temp_downloads = [f for f in current_files if f.endswith('.crdownload') or f.endswith('.tmp')]
        
        if not temp_downloads:
            # Download appears complete, get new files
            new_files = current_files - initial_files
            if new_files:
                # Find the most recently created file that's not a temp file
                valid_new_files = [f for f in new_files if not (f.endswith('.crdownload') or f.endswith('.tmp'))]
                
                if valid_new_files:
                    newest_file = max([os.path.join(download_path, f) for f in valid_new_files], 
                                      key=os.path.getctime)
                    file_size = os.path.getsize(newest_file)
                    
                    # Only rename if file size is valid
                    if file_size > 0:
                        new_name = f"{company_dict['name']}.csv"  # Assuming the file is a CSV
                        new_file_path = os.path.join(download_path, new_name)
                        
                        # Check if destination file already exists, remove it if it does
                        if os.path.exists(new_file_path):
                            try:
                                os.remove(new_file_path)
                                print(f"Removed existing file: {new_name}")
                            except Exception as e:
                                print(f"Error removing existing file: {e}")
                                return False
                        
                        try:
                            # Add a small delay to ensure file is fully available
                            time.sleep(0.5)
                            os.rename(newest_file, new_file_path)
                            print(f"Download completed and renamed: {os.path.basename(newest_file)} → {new_name} ({file_size} bytes)")
                            return True
                        except Exception as e:
                            print(f"Error renaming file: {e}")
                            print(f"File exists check: {os.path.exists(newest_file)}")
                            return False
                    else:
                        print("Warning: Downloaded file is empty")
                        return False
            else:
                print("Warning: Download seems complete but no new files found")
            break
        time.sleep(1)
    
    print(f"Download timed out after {timeout} seconds")
    return False

class search_screener:
    def __init__(self, dld_path):
        self.result_dict = []
        self.url = "https://finance.yahoo.com/"
        if getattr(sys, 'frozen', False):
            # Running as a bundled executable
            # The _MEI_ path is where PyInstaller unpacks the files
            base_path = sys._MEIPASS # Use _MEI_TEMP if you use --temp-dir
            # If not using --temp-dir, use sys._MEI_
            # base_path = sys._MEI_
        else:
            # Running as a regular Python script
            base_path = os.path.dirname(os.path.abspath(__file__))
           
        self.chromedriver_path = os.path.join(base_path, "drivers", "chromedriver.exe")
        
        # Verify if chromedriver exists
        if not os.path.exists(self.chromedriver_path):
            raise FileNotFoundError(f"Chromedriver not found at: {self.chromedriver_path}. Please ensure it's in the 'drivers' folder relative to your executable.")
        
        self.service = Service(ChromeDriverManager().install()) # the required chrome service to be run 
        self.options = Options()
        self.options.add_argument("--headless")  # Runs browser in background
        # self.options.add_argument("--disable-gpu")
        # self.options.add_argument("--window-size=1920,1080")
                
        self.download_path = dld_path
        prefs = {
        "download.default_directory": self.download_path,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
        }
        self.options.add_experimental_option("prefs", prefs)
        self.options.page_load_strategy = "eager"
        self.driver = webdriver.Chrome(service = self.service, options= self.options) # launch chrome uing the chrome service (different service is needed for different browsers)
        self.driver.get(self.url)
        wait_strategies = [
            (By.ID, "ybar-sbq"),
            (By.CSS_SELECTOR, "input[placeholder*='Search']"),
            (By.CSS_SELECTOR, "input[name='p']"),
            (By.CSS_SELECTOR, "#ybar-search input")]
        
        for strategy in wait_strategies:
            try:
                WebDriverWait(self.driver, 3).until(
                    EC.element_to_be_clickable(strategy)
                )
                return
            except TimeoutException:
                continue
        self.company = None
        self.topratios = []
        

    def search(self, string):
        self.result_dict = []
        search_box_selectors = [
            "ybar-sbq",
            "input[placeholder*='Search']",
            "input[name='p']",
            "#ybar-search input"
            ]

        searchbox = None
        for selector in search_box_selectors:
            try:
                if selector.startswith("#") or "[" in selector:
                    searchbox = self.driver.find_element(By.CSS_SELECTOR, selector)
                else:
                    searchbox = self.driver.find_element(By.ID, selector)
                break
            except NoSuchElementException:
                continue
        
        if not searchbox:
            raise Exception("Could not find search box")
        
        searchbox.clear()
        time.sleep(0.5)
        searchbox.send_keys(string)
        time.sleep(0.5)
        
        try_count = 0
        max_try = 3
        
        while try_count <= max_try:
            try:
                time.sleep(0.5)
                html = self.driver.page_source
                soup = BeautifulSoup(html, "html.parser")
                names = soup.select("ul[role='listbox'] li")
                if names:
                    break
                else:
                    try_count += 1
                    pass
            except Exception as e:
                try_count += 1
                continue

        
        
        for li in soup.find_all("li", title=True):  # All <li> tags that have a 'title' attribute
            company_name = li["title"]
            quote_div = li.find("div", class_=lambda x: x and "quoteSymbol" in x)
            if quote_div:
                ticker = quote_div.get_text(strip=True)

                self.result_dict.append({
                        "ticker": ticker,
                        "name": company_name,
                    })
            else:
                pass
        # for name in names:
        #     try:
        #         if name.text.strip() and "Search everywhere" not in name.text:
        #             results.append(name.text)
        #     except StaleElementReferenceException:
        #         continue
        
        # for item in results:
        #     parts = item.split('\n')
        #     if len(parts) == 4:
        #         self.result_dict.append({
        #             "ticker": parts[0],
        #             "name": parts[1],
        #             "instrument_type": parts[2],
        #             "exchange": parts[3]
        #         })
               
  
# instance = search_screener(download_path)

# rel = instance.search("hindal")
# print(instance.result_dict)
# rel = instance.search("reliance")
# print(instance.result_dict)


# print("d")
# instance.click_suggestion(rel[0])
# instance.co_page()    

# fromdate = "31/03/2020"
# todate = "31/03/2025"



# company = instance.company
# print(company)
# print(instance.topratios)
# time.sleep(5)
