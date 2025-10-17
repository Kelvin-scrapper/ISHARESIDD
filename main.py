import os
import time
import pandas as pd
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from bs4 import BeautifulSoup
import undetected_chromedriver as uc
import re
from typing import Dict, List, Optional, Tuple

# Configuration
HEADLESS_MODE = True  # Set to True to run in background, False to see browser
TAKE_SCREENSHOTS = False  # Enable/disable screenshot functionality
DEBUG_MODE = True  # Enable detailed debug logging

# Create screenshots directory if it doesn't exist
if TAKE_SCREENSHOTS:
    SCREENSHOT_DIR = "screenshots"
    os.makedirs(SCREENSHOT_DIR, exist_ok=True)

# ETF data from the document
etf_data = [
    ("JPM EM Bonds ETF", "https://www.ishares.com/us/products/239572/ishares-jp-morgan-usd-emerging-markets-bond-etf"),
    ("Iboxx High Yield Corp Bond ETF", "https://www.ishares.com/us/products/239565/ishares-iboxx-high-yield-corporate-bond-etf"),
    ("Iboxx IG Corp Bond ETF", "https://www.ishares.com/us/products/239566/ishares-iboxx-investment-grade-corporate-bond-etf"),
    ("TIPS Bond ETF", "https://www.ishares.com/us/products/239467/ishares-tips-bond-etf"),
    ("20 Yr Treasury Bond ETF", "https://www.ishares.com/us/products/239454/ishares-20-year-treasury-bond-etf"),
    ("7-10 Year Treasury Bond ETF", "https://www.ishares.com/us/products/239456/"),
    ("iShares J.P. Morgan $ EM Bond UCITS ETF", "https://www.ishares.com/uk/individual/en/products/251824/ishares-jp-morgan-emerging-markets-bond-ucits-etf"),
    ("iShares $ Treasury Bond 1-3yr UCITS ETF", "https://www.ishares.com/uk/individual/en/products/251715/ishares-treasury-bond-13yr-ucits-etf"),
    ("iShares $ High Yield Corp Bond UCITS ETF", "https://www.ishares.com/uk/individual/en/products/251833/ishares-high-yield-corporate-bond-ucits-etf"),
    ("iShares € Corp Bond 1-5yr UCITS ETF", "https://www.ishares.com/uk/individual/en/products/251728/ishares-euro-corporate-bond-15yr-ucits-etf"),
    ("iShares € High Yield Corp Bond UCITS ETF", "https://www.ishares.com/uk/individual/en/products/251843/ishares-euro-high-yield-corporate-bond-ucits-etf"),
    ("iShares Core € Corp Bond UCITS ETF", "https://www.ishares.com/uk/individual/en/products/251726/ishares-euro-corporate-bond-ucits-etf"),
    ("iShares MBS ETF", "https://www.ishares.com/us/products/239465/ishares-mbs-etf"),
    ("iShares Short-Term Corporate Bond ETF", "https://www.ishares.com/us/products/239451/ishares-13-year-credit-bond-etf"),
    ("iShares Intermediate-Term Corporate Bond ETF", "https://www.ishares.com/us/products/239463/ishares-intermediate-credit-bond-etf")
]

def log_debug(message: str, etf_name: str = ""):
    """Log debug messages if debug mode is enabled"""
    if DEBUG_MODE:
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = f"[{timestamp}]"
        if etf_name:
            prefix += f" [{etf_name}]"
        print(f"{prefix} {message}")

def take_screenshot(driver, etf_name: str, step_name: str):
    """Take a screenshot with descriptive filename"""
    if not TAKE_SCREENSHOTS:
        return
    
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    # Clean ETF name for filename
    clean_etf_name = etf_name.replace(" ", "_").replace("$", "USD").replace("€", "EUR").replace("/", "_").replace(".", "")
    filename = f"{SCREENSHOT_DIR}/{clean_etf_name}_{step_name}_{timestamp}.png"
    
    try:
        driver.save_screenshot(filename)
        log_debug(f"Screenshot saved: {filename}", etf_name)
    except Exception as e:
        log_debug(f"Failed to take screenshot: {str(e)}", etf_name)

def setup_driver():
    """Set up undetected Chrome WebDriver with options"""
    options = uc.ChromeOptions()
    
    if HEADLESS_MODE:
        options.add_argument("--headless")
    
    # Basic options that work with undetected-chromedriver
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-blink-features=AutomationControlled")
    
    try:
        # Use undetected-chromedriver
        driver = uc.Chrome(options=options, version_main=None)
        # Execute script after driver is created
        driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        return driver
    except Exception as e:
        log_debug(f"Error creating driver with undetected-chromedriver: {str(e)}")
        log_debug("Trying with standard Chrome driver...")
        
        # Fallback to standard selenium
        from selenium import webdriver
        from selenium.webdriver.chrome.service import Service
        
        # Try to find Chrome automatically
        service = Service()
        options = webdriver.ChromeOptions()
        
        if HEADLESS_MODE:
            options.add_argument("--headless")
        
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        
        try:
            driver = webdriver.Chrome(service=service, options=options)
            return driver
        except Exception as e2:
            log_debug(f"Error with standard Chrome driver: {str(e2)}")
            raise Exception("Could not initialize any Chrome driver")

def wait_and_click(driver, by: By, value: str, timeout: int = 2, etf_name: str = "") -> bool:
    """Wait for element to be clickable and click it"""
    try:
        element = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by, value))
        )
        element.click()
        log_debug(f"Successfully clicked element: {value}", etf_name)
        return True
    except TimeoutException:
        return False
    except Exception as e:
        log_debug(f"Error clicking element {value}: {str(e)}", etf_name)
        return False

def handle_interstitial_pages(driver, etf_name: str) -> bool:
    """Handle various interstitial pages that might appear"""
    log_debug("Checking for interstitial pages...", etf_name)
    
    # Handle cookie consent - reduce timeout to speed up process
    cookie_selectors = [
        "//button[contains(translate(text(), 'ACCEPT', 'accept'), 'accept')]",
        "//button[contains(text(), 'Accept All')]",
        "//button[contains(text(), 'I Agree')]",
        "//button[contains(text(), 'Accept Cookies')]",
        "//button[contains(@class, 'accept')]",
        "//a[contains(translate(text(), 'ACCEPT', 'accept'), 'accept')]",
        "//button[contains(@id, 'accept')]",
        "//button[contains(@aria-label, 'Accept')]",
        "//div[contains(@class, 'cookie')]//button[contains(translate(text(), 'ACCEPT', 'accept'), 'accept')]",
        "//div[contains(@class, 'consent')]//button[contains(translate(text(), 'ACCEPT', 'accept'), 'accept')]",
        "//button[contains(@class, 'cookie-consent')]",
        "//button[contains(text(), 'Agree')]",
        "//button[contains(text(), 'Continue')]",
        "//button[contains(translate(text(), 'CONTINUE', 'continue'), 'continue')]"
    ]
    
    for selector in cookie_selectors:
        if wait_and_click(driver, By.XPATH, selector, 2, etf_name):
            time.sleep(1)
            take_screenshot(driver, etf_name, "after_cookie_click")
            return True
    
    # Handle investor type selection - reduce timeout
    continue_selectors = [
        "//button[contains(translate(text(), 'CONTINUE', 'continue'), 'continue')]",
        "//a[contains(translate(text(), 'CONTINUE', 'continue'), 'continue')]",
        "//button[contains(@class, 'continue')]",
        "//a[contains(@class, 'continue')]",
        "//button[contains(text(), 'individual investor')]",
        "//a[contains(text(), 'individual investor')]",
        "//button[contains(text(), 'proceed')]",
        "//a[contains(text(), 'proceed')]"
    ]
    
    for selector in continue_selectors:
        if wait_and_click(driver, By.XPATH, selector, 2, etf_name):
            time.sleep(2)
            take_screenshot(driver, etf_name, "after_continue_click")
            return True
    
    log_debug("No interstitial pages found or handled", etf_name)
    return False

def scroll_to_content(driver, etf_name: str) -> bool:
    """Scroll the page to reveal dynamic content"""
    log_debug("Scrolling to reveal content...", etf_name)
    
    try:
        # Method 1: Scroll down incrementally
        for i in range(5):
            driver.execute_script(f"window.scrollBy(0, {500 + i * 100});")
            time.sleep(0.5)
        
        # Method 2: Scroll to bottom
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1)
        
        # Method 3: Scroll back up a bit
        driver.execute_script("window.scrollBy(0, -300);")
        time.sleep(0.5)
        
        # Method 4: Try to find and scroll to key sections
        potential_sections = [
            "//h2[contains(text(), 'Portfolio')]",
            "//h2[contains(text(), 'Characteristics')]",
            "//h3[contains(text(), 'Portfolio')]",
            "//h3[contains(text(), 'Characteristics')]",
            "//div[contains(@class, 'portfolio')]",
            "//div[contains(@class, 'characteristics')]",
            "//section[contains(@class, 'portfolio')]",
            "//section[contains(@class, 'characteristics')]"
        ]
        
        for selector in potential_sections:
            try:
                element = driver.find_element(By.XPATH, selector)
                driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element)
                time.sleep(1)
                log_debug(f"Scrolled to section: {selector}", etf_name)
                return True
            except NoSuchElementException:
                continue
        
        return True
    except Exception as e:
        log_debug(f"Error during scrolling: {str(e)}", etf_name)
        return False

def find_data_section(driver, etf_name: str) -> Optional[any]:
    """Find the section containing portfolio data using multiple methods"""
    log_debug("Searching for data section...", etf_name)
    
    # Method 1: Try known IDs
    section_ids = [
        "fundamentalsAndRisk",
        "portfolio-characteristics",
        "portfolio",
        "characteristics",
        "key-facts",
        "fund-data"
    ]
    
    for section_id in section_ids:
        try:
            section = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, section_id))
            )
            log_debug(f"Found section by ID: {section_id}", etf_name)
            return section
        except TimeoutException:
            continue
    
    # Method 2: Try to find by headings
    heading_selectors = [
        "//h2[contains(translate(text(), 'PORTFOLIO', 'portfolio'), 'portfolio') and contains(translate(text(), 'CHARACTERISTICS', 'characteristics'), 'characteristics')]",
        "//h3[contains(translate(text(), 'PORTFOLIO', 'portfolio'), 'portfolio') and contains(translate(text(), 'CHARACTERISTICS', 'characteristics'), 'characteristics')]",
        "//h2[contains(translate(text(), 'EFFECTIVE', 'effective'), 'effective')]",
        "//h3[contains(translate(text(), 'EFFECTIVE', 'effective'), 'effective')]",
        "//h2[contains(translate(text(), 'DURATION', 'duration'), 'duration')]",
        "//h3[contains(translate(text(), 'DURATION', 'duration'), 'duration')]"
    ]
    
    for selector in heading_selectors:
        try:
            heading = driver.find_element(By.XPATH, selector)
            # Try to find the parent section
            section = heading.find_element(By.XPATH, "./ancestor::div[contains(@class, 'section') or contains(@class, 'component') or contains(@class, 'container')]")
            log_debug(f"Found section by heading: {selector}", etf_name)
            return section
        except NoSuchElementException:
            continue
    
    # Method 3: Look for any section containing duration-related text
    try:
        duration_elements = driver.find_elements(By.XPATH, "//*[contains(translate(text(), 'EFFECTIVE DURATION', 'effective duration'), 'effective duration')]")
        if duration_elements:
            for element in duration_elements[:3]:  # Check first 3 elements
                section = element.find_element(By.XPATH, "./ancestor::div[contains(@class, 'section') or contains(@class, 'component') or contains(@class, 'container')]")
                log_debug("Found section containing duration text", etf_name)
                return section
    except NoSuchElementException:
        pass
    
    log_debug("Could not find specific data section", etf_name)
    return None

def parse_date_flexible(date_str: str) -> str:
    """Parse various date formats and return YYYY-MM-DD"""
    date_formats = [
        "%B %d, %Y",      # October 15, 2025
        "%b %d, %Y",      # Oct 15, 2025
        "%d/%b/%Y",       # 15/Oct/2025
        "%d/%B/%Y",       # 15/October/2025
        "%d-%b-%Y",       # 15-Oct-2025
        "%d-%B-%Y",       # 15-October-2025
        "%m/%d/%Y",       # 10/15/2025
        "%m-%d-%Y",       # 10-15-2025
        "%d/%m/%Y",       # 15/10/2025
        "%d-%m-%Y",       # 15-10-2025
        "%Y-%m-%d",       # 2025-10-15
        "%Y/%m/%d",       # 2025/10/15
        "%b %d %Y",       # Oct 15 2025
        "%B %d %Y",       # October 15 2025
    ]
    
    # Clean the date string
    date_str = date_str.strip().replace("as of", "").replace("As of", "").strip()
    
    for fmt in date_formats:
        try:
            date_obj = datetime.strptime(date_str, fmt)
            return date_obj.strftime("%Y-%m-%d")
        except ValueError:
            continue
    
    # Try to extract date using regex as last resort
    date_patterns = [
        r'(\d{1,2})/(\d{1,2})/(\d{4})',  # MM/DD/YYYY or DD/MM/YYYY
        r'(\d{1,2})-(\d{1,2})-(\d{4})',  # MM-DD-YYYY or DD-MM-YYYY
        r'(\d{4})-(\d{1,2})-(\d{1,2})',  # YYYY-MM-DD
        r'([A-Za-z]{3,9}) (\d{1,2}),? (\d{4})',  # Month DD, YYYY
        r'(\d{1,2}) ([A-Za-z]{3,9}) (\d{4})',  # DD Month YYYY
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, date_str)
        if match:
            try:
                groups = match.groups()
                if len(groups) == 3:
                    # Try different interpretations
                    if pattern.startswith(r'(\d{4})'):
                        # YYYY-MM-DD format
                        year, month, day = groups
                    elif '/' in pattern or '-' in pattern:
                        # Try MM/DD/YYYY first, then DD/MM/YYYY
                        month, day, year = groups
                        if int(month) > 12:  # Likely DD/MM/YYYY
                            day, month = month, day
                    else:
                        # Text month format
                        if groups[0].isalpha():
                            # Month DD, YYYY
                            month_str, day, year = groups
                            month = datetime.strptime(month_str[:3], "%b").month
                        else:
                            # DD Month YYYY
                            day, month_str, year = groups
                            month = datetime.strptime(month_str[:3], "%b").month
                    
                    date_obj = datetime(int(year), int(month), int(day))
                    return date_obj.strftime("%Y-%m-%d")
            except (ValueError, TypeError):
                continue
    
    log_debug(f"Could not parse date: {date_str}")
    return "N/A"

def extract_duration_and_date(soup: BeautifulSoup, etf_name: str) -> Tuple[str, str]:
    """Extract both duration and date using multiple methods"""
    log_debug("Extracting duration and date...", etf_name)
    
    duration_value = "N/A"
    date_value = "N/A"
    
    # Method 1: Look for structured data using BeautifulSoup find methods
    try:
        # Find all elements containing "Effective Duration"
        duration_elements = soup.find_all(string=lambda text: text and "effective duration" in text.lower())
        log_debug(f"Found {len(duration_elements)} elements with 'effective duration' text", etf_name)
        
        for element in duration_elements:
            # Get the parent element
            parent = element.parent
            if parent:
                # Look for the value in siblings
                next_sibling = parent.next_sibling
                while next_sibling:
                    if hasattr(next_sibling, 'get_text'):
                        sibling_text = next_sibling.get_text(strip=True)
                        if re.search(r'\d+\.?\d*', sibling_text):
                            duration_value = sibling_text
                            log_debug(f"Found duration via sibling: {duration_value}", etf_name)
                            
                            # Look for date in the same container
                            container = parent.find_parent()
                            if container:
                                container_text = container.get_text()
                                date_match = re.search(r'as of ([^\n]+)', container_text, re.IGNORECASE)
                                if date_match:
                                    date_value = parse_date_flexible(date_match.group(1))
                                    log_debug(f"Found date via container: {date_value}", etf_name)
                                    return duration_value, date_value
                    next_sibling = next_sibling.next_sibling
                
                # Look for the value in parent's children
                for child in parent.find_all():
                    child_text = child.get_text(strip=True)
                    if re.search(r'\d+\.?\d*', child_text) and "effective duration" not in child_text.lower():
                        duration_value = child_text
                        log_debug(f"Found duration via child: {duration_value}", etf_name)
                        
                        # Look for date in the same container
                        container = parent.find_parent()
                        if container:
                            container_text = container.get_text()
                            date_match = re.search(r'as of ([^\n]+)', container_text, re.IGNORECASE)
                            if date_match:
                                date_value = parse_date_flexible(date_match.group(1))
                                log_debug(f"Found date via container: {date_value}", etf_name)
                                return duration_value, date_value
    except Exception as e:
        log_debug(f"Structured extraction failed: {str(e)}", etf_name)
    
    # Method 2: Text-based extraction
    full_text = soup.get_text()
    
    # Extract duration with more specific patterns
    duration_patterns = [
        r'Effective Duration\s*([0-9]+\.?[0-9]*)',
        r'Effective Duration\s*([0-9]+\.?[0-9]*\s*(?:years|yrs|year))',
        r'Effective Duration[:\s]+([0-9]+\.?[0-9]*)',
        r'Effective Duration[:\s]+([0-9]+\.?[0-9]*\s*(?:years|yrs|year))',
        r'Duration\s*([0-9]+\.?[0-9]*)',
        r'Duration\s*([0-9]+\.?[0-9]*\s*(?:years|yrs|year))',
        r'Duration[:\s]+([0-9]+\.?[0-9]*)',
        r'Duration[:\s]+([0-9]+\.?[0-9]*\s*(?:years|yrs|year))'
    ]
    
    for pattern in duration_patterns:
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            duration_value = match.group(1).strip()
            log_debug(f"Found duration via text pattern: {duration_value}", etf_name)
            break
    
    # Extract date
    date_patterns = [
        r'as of ([A-Za-z0-9\s/,-]+)',
        r'as of: ([A-Za-z0-9\s/,-]+)',
        r'date: ([A-Za-z0-9\s/,-]+)',
        r'([A-Za-z]{3,9} \d{1,2},? \d{4})',
        r'(\d{1,2}/[A-Za-z]{3,9}/\d{4})',
        r'(\d{1,2}-[A-Za-z]{3,9}-\d{4})',
        r'(\d{1,2}/\d{1,2}/\d{4})',
        r'(\d{4}-\d{1,2}-\d{1,2})'
    ]
    
    for pattern in date_patterns:
        matches = re.findall(pattern, full_text, re.IGNORECASE)
        if matches:
            date_value = parse_date_flexible(matches[0])
            if date_value != "N/A":
                log_debug(f"Found date via text pattern: {date_value}", etf_name)
                break
    
    # Method 3: Context-based extraction
    # Look for duration and date that appear close together
    lines = full_text.split('\n')
    for i, line in enumerate(lines):
        if 'effective duration' in line.lower():
            # Check current line for date
            date_match = re.search(r'as of ([A-Za-z0-9\s/,-]+)', line, re.IGNORECASE)
            if date_match:
                date_value = parse_date_flexible(date_match.group(1))
            
            # Check surrounding lines
            for j in range(max(0, i-2), min(len(lines), i+3)):
                date_match = re.search(r'as of ([A-Za-z0-9\s/,-]+)', lines[j], re.IGNORECASE)
                if date_match:
                    date_value = parse_date_flexible(date_match.group(1))
                    break
            
            # Extract duration from current line
            duration_match = re.search(r'([0-9]+\.?\d*)', line)
            if duration_match:
                duration_value = duration_match.group(1)
            
            if duration_value != "N/A" or date_value != "N/A":
                log_debug(f"Found via context extraction: Duration={duration_value}, Date={date_value}", etf_name)
                break
    
    return duration_value, date_value

def extract_effective_duration(driver, etf_name: str, url: str) -> Optional[Dict[str, str]]:
    """Navigate to ETF page and extract Effective Duration data"""
    try:
        log_debug(f"Starting extraction for {etf_name}", etf_name)
        log_debug(f"URL: {url}", etf_name)
        
        # Navigate to page
        driver.get(url)
        time.sleep(3)
        take_screenshot(driver, etf_name, "page_loaded")
        
        # Handle interstitial pages
        handle_interstitial_pages(driver, etf_name)
        
        # Wait for page to fully load
        time.sleep(2)
        
        # Scroll to reveal content
        scroll_to_content(driver, etf_name)
        take_screenshot(driver, etf_name, "after_scroll")
        
        # Try to find data section
        data_section = find_data_section(driver, etf_name)
        if data_section:
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", data_section)
            time.sleep(1)
            take_screenshot(driver, etf_name, "section_found")
        
        # Get page source and extract data
        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')
        
        # Save HTML for debugging
        if TAKE_SCREENSHOTS:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            clean_etf_name = etf_name.replace(" ", "_").replace("$", "USD").replace("€", "EUR").replace("/", "_").replace(".", "")
            html_file = f"{SCREENSHOT_DIR}/{clean_etf_name}_page_source_{timestamp}.html"
            with open(html_file, 'w', encoding='utf-8') as f:
                f.write(page_source)
            log_debug(f"Page source saved: {html_file}", etf_name)
        
        # Extract duration and date
        duration_value, date_value = extract_duration_and_date(soup, etf_name)
        
        result = {
            "ETF Name": etf_name,
            "Effective Duration": duration_value,
            "As of Date": date_value
        }
        
        if duration_value != "N/A" or date_value != "N/A":
            take_screenshot(driver, etf_name, "extraction_success")
            log_debug(f"Successfully extracted: Duration={duration_value}, Date={date_value}", etf_name)
        else:
            take_screenshot(driver, etf_name, "extraction_failed")
            log_debug("Extraction failed - no data found", etf_name)
        
        return result
        
    except Exception as e:
        log_debug(f"Error processing {etf_name}: {str(e)}", etf_name)
        take_screenshot(driver, etf_name, "processing_error")
        return None

def main():
    """Main function to process all ETFs and save to Excel"""
    log_debug("Starting ETF data extraction...")
    
    # Set up WebDriver
    try:
        driver = setup_driver()
        log_debug("WebDriver successfully initialized")
    except Exception as e:
        log_debug(f"Failed to initialize WebDriver: {str(e)}")
        return
    
    # Initialize list to store all data
    all_data = []
    
    try:
        # Process each ETF
        for etf_name, url in etf_data:
            result = extract_effective_duration(driver, etf_name, url)
            if result:
                all_data.append(result)
                log_debug(f"Successfully processed {etf_name}")
            else:
                log_debug(f"Failed to process {etf_name}")
            
            # Add delay between requests
            time.sleep(2)
            
    finally:
        # Close the driver
        try:
            driver.quit()
            log_debug("WebDriver closed")
        except:
            log_debug("Error closing WebDriver (may already be closed)")
    
    # Create DataFrame and save to Excel
    if all_data:
        df = pd.DataFrame(all_data)
        
        # Save to Excel
        output_file = "ETF_Effective_Duration.xlsx"
        df.to_excel(output_file, index=False)
        log_debug(f"Data saved to {output_file}")
        
        # Print summary
        success_count = len([d for d in all_data if d['Effective Duration'] != 'N/A'])
        log_debug(f"Extraction complete: {success_count}/{len(all_data)} ETFs with duration data")
    else:
        log_debug("No data was extracted. Excel file not created.")

if __name__ == "__main__":
    main()