import time
import openpyxl
from seleniumbase import Driver
from selenium.webdriver import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException

# ------------ Config ------------
EXCEL_PATH = "prices.xlsx"          # your workbook
HTML_OUTPUT = "page_source.html"    # saved full HTML
URL = ("https://www.upwork.com/hire/landing/?utm_campaign=SEMBrand_Google_INTL_Marketplace_Core"
       "&utm_medium=PaidSearch&utm_content=150606034558&utm_term=upwork&campaignid=20227594544"
       "&matchtype=e&device=c&utm_source=google")
# Use uc=True as you prefer for undetected mode
# Set headless=False so you can see what's happening (set True for headless)
# --------------------------------

wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
sheet = wb.active

driver = Driver(uc=True, headless=False)  # undetected chromium, visible window

try:
    # Open the page (with auto-reconnect attempts)
    driver.uc_open_with_reconnect(URL, 4)

    # wait until body exists
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.TAG_NAME, "body"))
    )

    # ⚡ grab HTML immediately before redirect
    time.sleep(1)
    html = driver.page_source
    with open("upwork_page.html", "w", encoding="utf-8") as f:
        f.write(html)

    print("✅ Saved Upwork HTML before redirect.")

    # optional: stop further redirects (helps block the error page)
    driver.execute_cdp_cmd("Network.enable", {})
    driver.execute_cdp_cmd("Network.setBlockedURLs", {"urls": ["*not-found*", "*Looking-for-something*"]})

finally:
    try:
        wb.close()
    except Exception:
        pass
    driver.quit()
    print("Browser closed.")
