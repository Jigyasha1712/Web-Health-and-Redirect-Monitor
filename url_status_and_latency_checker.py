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

# 
options = webdriver.ChromeOptions()
options.add_argument("--headless")
options.add_argument("--ignore-certificate-errors")
options.add_argument("--incognito")
options.add_argument("--disable-popup-blocking")
options.add_argument("--log-level=3")

service = Service(CHROME_DRIVER_PATH)
driver = webdriver.Chrome(service=service, options=options)

def get_http_status(url):
    """Returns actual HTTP status code of the website using requests."""
    try:
        response = requests.get(url, timeout=10)
        return response.status_code
    except requests.exceptions.SSLError:
        return "SSL Error"
    except requests.exceptions.ConnectionError:
        return "Connection Error"
    except requests.exceptions.Timeout:
        return "Timeout Error"
    except requests.exceptions.RequestException as e:
        return f"HTTP Error: {str(e)}"

def check_website_status(url):
    """Checks website status using requests."""
    print(f"\n🔄 Checking: {url}")

 
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

    print(f"🔍 HTTP STATUS: {url} → {error_code}")

    return url, error_code, status

def retry_with_selenium(url):
    """Retries opening website in Selenium, scrolling, and checking status."""
    print(f"\n🔄 Retrying with Selenium: {url}")

    try:
        driver.get(url)
        time.sleep(3)  # Let it load
        print(f"✅ OPENED: {url}")

        # ✅ Scroll page 10 times
        for _ in range(10):
            driver.find_element(By.TAG_NAME, "body").send_keys(Keys.PAGE_DOWN)
            time.sleep(0.5)

        # ✅ Check if website is in Maintenance Mode
        page_text = driver.page_source.lower()
        maintenance_keywords = ["under maintenance", "coming soon", "site is down", "maintenance mode"]
        if any(keyword in page_text for keyword in maintenance_keywords):
            error_code = "Maintenance Mode"
            status = "Not Working"
        else:
            error_code = "Working after Selenium"
            status = "Working"

    except Exception as e:
        print(f"❌ ERROR: {url} → {str(e)}")
        error_code = "Unknown Error"
        status = "Not Working"

    return error_code, status

# ✅ Step 1: Check All Websites (First Attempt)
results = []

for website in websites:
    website = website.strip()
    if not website.startswith(("http://", "https://")):
        website = "https://" + website
    
    url, first_error_code, first_status = check_website_status(website)

    retry_error_code, retry_status = "-", "-"  # Default if no retry is needed

    # ✅ Only retry if NOT 403 (Forbidden)
    if first_status == "Not Working" and "403" not in first_error_code:
        retry_error_code, retry_status = retry_with_selenium(url)

    # ✅ Determine Final Status
    if first_status == "Working":
        final_status = "Working"
    elif retry_status == "Working":
        final_status = "Working"
    else:
        final_status = "Not Working"

    results.append((url, first_error_code, first_status, retry_error_code, retry_status, final_status))

# ✅ Save Results to Excel
output_path = "website_status_final_Checker1.xlsx"
df_results = pd.DataFrame(results, columns=["Website URL", "First Check Error Code", "First Check Status", "Retry Error Code", "Retry Status", "Final Status"])
df_results.to_excel(output_path, index=False)

print(f"\n✅ Results saved to {output_path}")

# ✅ Close WebDriver
driver.quit()
