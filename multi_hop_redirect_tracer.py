import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time

CHROME_DRIVER_PATH = "chromedriver-win64/chromedriver.exe"
input_file = "websites.xlsx"
df = pd.read_excel(input_file)
websites = df["Website URL"].dropna().tolist()

options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--ignore-certificate-errors")
options.add_argument("--incognito")
options.add_argument("--disable-popup-blocking")
options.add_argument("--log-level=3")

service = Service(CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)

def get_http_status(url):
    try:
        response = requests.get(url, timeout=10)
        print(f"Checking {url} - HTTP Status: {response.status_code}")
        return response.status_code
    except requests.exceptions.SSLError:
        print(f"{url} - SSL Certificate Issue")
        return "SSL Error"
    except requests.exceptions.ConnectionError:
        print(f"{url} - Connection Error")
        return "Connection Error"
    except requests.exceptions.Timeout:
        print(f"{url} - Timeout Error")
        return "Timeout Error"
    except requests.exceptions.RequestException as e:
        print(f"{url} - HTTP Error: {str(e)}")
        return f"HTTP Error: {str(e)}"

def check_website_status(url):
    http_status = get_http_status(url)
    status_mapping = {
        200: "Working",
        403: "Forbidden (403) - Blocked",
        404: "Not Found (404)",
        500: "Internal Server Error (500)",
        503: "Service Unavailable (503)",
        "SSL Error": "SSL Certificate Issue",
        "Connection Error": "Connection Failed",
        "Timeout Error": "Server Timeout"
    }
    error_code = status_mapping.get(http_status, f"Error ({http_status})")
    status = "Working" if http_status == 200 else "Not Working"
    print(f"{url} - Initial Check: {error_code}, Status: {status}")
    return url, error_code, status

def retry_with_selenium(url):
    try:
        print(f"Retrying with Selenium: {url}")
        driver.get(url)
        time.sleep(3)
        start_time = time.time()
        scrolled = False
        
        while time.time() - start_time < 10:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.PAGE_DOWN)
            time.sleep(0.5)
            scrolled = True
        
        page_text = driver.page_source.lower()
        keywords = ["about us", "contact", "privacy policy", "policy"]
        
        keyword_found = any(keyword in page_text for keyword in keywords)
        
        if scrolled and keyword_found:
            error_code = "Scrolled & Keyword Found - Working"
            status = "Working"
        else:
            error_code = "Unknown Issue"
            status = "Not Working"
    except Exception:
        error_code = "Unknown Error"
        status = "Not Working"
    print(f"{url} - Selenium Check: {error_code}, Status: {status}")
    return error_code, status

def deep_retry_with_selenium(url):
    try:
        print(f"Deep Retrying with Selenium: {url}")
        driver.get(url)
        time.sleep(3)
        start_time = time.time()
        while time.time() - start_time < 10:
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.PAGE_DOWN)
            time.sleep(0.5)
        
        page_text = driver.page_source.lower()
        keywords = ["about us", "contact", "privacy policy", "policy"]
        
        for keyword in keywords:
            if keyword in page_text:
                print(f"{url} - Found '{keyword}' - Working")
                return "Keyword Found - Working", "Working"
        
        return "Keyword Not Found - Not Working", "Not Working"
    except Exception:
        return "Deep Retry Failed", "Not Working"

results = []

for website in websites:
    website = website.strip()
    if not website.startswith(("http://", "https://")):
        website = "https://" + website
    
    print(f"Processing: {website}")
    url, first_error_code, first_status = check_website_status(website)
    retry_error_code, retry_status = "-", "-"
    deep_retry_error_code, deep_retry_status = "-", "-"
    
    if first_status == "Not Working" and "403" not in first_error_code:
        retry_error_code, retry_status = retry_with_selenium(url)
    
    if retry_status == "Not Working":
        deep_retry_error_code, deep_retry_status = deep_retry_with_selenium(url)
    
    final_status = "Working" if "Working" in [first_status, retry_status, deep_retry_status] else "Not Working"
    print(f"Final Status for {website}: {final_status}\n")
    results.append((url, first_error_code, first_status, retry_error_code, retry_status, deep_retry_error_code, deep_retry_status, final_status))

detailed_output = "website_status_detailed_c1.xlsx"
df_detailed = pd.DataFrame(results, columns=["Website URL", "First Check Error Code", "First Check Status", "Retry Error Code", "Retry Status", "Deep Retry Error Code", "Deep Retry Status", "Final Status"])
df_detailed.to_excel(detailed_output, index=False)

summary_output = "website_status_summary_c1.xlsx"
df_summary = df_detailed[["Website URL", "Final Status"]]
df_summary.to_excel(summary_output, index=False)

driver.quit()

print("Website status check completed. Results saved in Excel files.")
