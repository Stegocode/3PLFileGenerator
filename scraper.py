# -*- coding: utf-8 -*-
"""
Homesource Scraper
==================
Automates four exports from the Homesource dealer management system, to feed
the flat file generator:

  1. Batch invoice export   (filtered by next business day + carrier truck)
  2. Model inventory export (full catalog)
  3. Serial inventory export (all locations, for unit-level cost)
  4. Orders detail export    (open orders for next business day)

Supported environments
----------------------
  Google Colab   — reads credentials from Colab Secrets; writes to /content
  Local Python   — reads credentials from local_config.py (which reads .env)

HOW TO USE (Google Colab)
  1. Add HS_USERNAME and HS_PASSWORD to Colab Secrets
  2. Edit the CONFIG block below to match your Homesource instance
  3. Run this cell BEFORE the flat file generator cell

HOW TO USE (local)
  1. Copy .env.example to .env and fill in credentials
  2. Edit the CONFIG block below to match your Homesource instance
  3. python scraper.py
"""

import time
import os
from datetime import datetime, timedelta

try:
    from IPython.display import Image, display
except ImportError:
    def display(x): pass
    class Image:
        def __init__(self, path): pass

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC


# ══════════════════════════════════════════════════════════════════
# CONFIGURATION — edit these values to adapt to a different instance
# ══════════════════════════════════════════════════════════════════
CONFIG = {
    # Homesource subdomain — each dealer has their own instance.
    # Example: if your Homesource URL is https://acme1.homesourcesystems.com,
    # set base_url to "https://acme1.homesourcesystems.com".
    "base_url": "https://your-subdomain.homesourcesystems.com",

    # Truck / route filter to apply on the batch invoice page.
    # This is the dropdown label shown in the Homesource Truck picker.
    "truck_filter": "HUB #1",

    # Business-day offset. The scraper always targets the NEXT business day
    # (Friday runs target Monday, weekend runs target Monday/Tuesday).
    # Leave True unless you specifically want same-day or multi-day targeting.
    "use_next_business_day": True,

    # Selenium wait timeout in seconds for page/element loads.
    "wait_timeout": 40,
}


# ══════════════════════════════════════════════════════════════════
# Credentials & download directory — Colab vs local
# ══════════════════════════════════════════════════════════════════
try:
    from google.colab import userdata
    HS_USERNAME  = userdata.get('HS_USERNAME')
    HS_PASSWORD  = userdata.get('HS_PASSWORD')
    DOWNLOAD_DIR = "/content"
    IS_COLAB     = True
except ImportError:
    from local_config import HS_USERNAME, HS_PASSWORD, DOWNLOAD_DIR
    IS_COLAB = False

BASE_URL = CONFIG["base_url"]

# Clear inbox before each run (local only) — keeps only today's files
if not IS_COLAB:
    import glob as _g, os as _o
    cleared = 0
    for _f in _g.glob(f'{DOWNLOAD_DIR}/*'):
        try:
            _o.remove(_f)
            cleared += 1
        except Exception:
            pass
    print(f'  Cleared inbox: {cleared} old files removed')
    print(f'  Inbox: {DOWNLOAD_DIR}')


# ══════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════
def get_next_business_day():
    """Return next business day as 'Month DD, YYYY' string."""
    today = datetime.today()
    if today.weekday() == 4:      # Friday → Monday
        delta = 3
    elif today.weekday() == 5:    # Saturday → Monday
        delta = 2
    else:
        delta = 1
    return (today + timedelta(days=delta)).strftime('%B %d, %Y')


DELIVERY_DATE = get_next_business_day() if CONFIG["use_next_business_day"] else datetime.today().strftime('%B %d, %Y')
print(f"  Delivery date: {DELIVERY_DATE}")


# Chrome setup
options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')
options.add_experimental_option('prefs', {
    'download.default_directory':   DOWNLOAD_DIR,
    'download.prompt_for_download': False,
    'download.directory_upgrade':   True,
    'safebrowsing.enabled':         True,
})

driver = webdriver.Chrome(options=options)
wait   = WebDriverWait(driver, CONFIG["wait_timeout"])


def screenshot(label):
    path = f'/tmp/{label}.png'
    driver.save_screenshot(path)
    print(f"  Screenshot: {label}")
    if IS_COLAB:
        display(Image(path))


def js_click(el):
    driver.execute_script("arguments[0].click();", el)


def wait_for_download(keyword, timeout=30):
    """Poll the download directory for a file containing *keyword*."""
    print(f"  Waiting for: {keyword}...")
    for _ in range(timeout):
        matches = [f for f in os.listdir(DOWNLOAD_DIR)
                   if keyword in f.lower() and not f.endswith('.crdownload')]
        if matches:
            print(f"  Downloaded: {matches[0]}")
            return os.path.join(DOWNLOAD_DIR, matches[0])
        time.sleep(1)
    print(f"  Timed out: {keyword}")
    return None


# ══════════════════════════════════════════════════════════════════
# STEP 1 — LOGIN
# ══════════════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  HOMESOURCE SCRAPER")
print("=" * 55)
print("\nLogging in...")

driver.get(f"{BASE_URL}/login")
time.sleep(8)
screenshot('00_login_page')
print(f"  Page title: {driver.title}")
print(f"  Current URL: {driver.current_url}")

wait.until(EC.presence_of_element_located((By.NAME, 'email')))
driver.find_element(By.NAME, 'email').send_keys(HS_USERNAME)
driver.find_element(By.NAME, 'password').send_keys(HS_PASSWORD)
driver.find_element(By.XPATH, "//button[@type='submit']").click()
time.sleep(4)
screenshot("01_after_login")
print(f"  URL: {driver.current_url}")
print("  Logged in")


# ══════════════════════════════════════════════════════════════════
# STEP 2 — SELECT WAREHOUSE
# ══════════════════════════════════════════════════════════════════
print("\nSelecting warehouse...")
try:
    save_btn = wait.until(EC.presence_of_element_located((By.ID, "save-current-location")))
    time.sleep(1)
    screenshot("02_location_popup")

    # Warehouse is already selected by default — just click Save
    print("  Warehouse already set — clicking Save")
    time.sleep(1)
    js_click(save_btn)
    time.sleep(3)
    screenshot("03_after_location")
    print(f"  URL: {driver.current_url}")
    print("  Location saved")
except Exception as e:
    print(f"  Location issue: {e}")
    screenshot("03_location_error")


# ══════════════════════════════════════════════════════════════════
# STEP 3 — BATCH INVOICE
# ══════════════════════════════════════════════════════════════════
print("\nNavigating to batch invoice...")
driver.get(f"{BASE_URL}/sales/batch-invoice")
time.sleep(5)
screenshot("04_batch_page")
print(f"  URL: {driver.current_url}")

print(f"  Setting date to {DELIVERY_DATE}...")
try:
    date_field = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "input.form-control.input[type='text']")
    ))
    driver.execute_script("arguments[0].value = '';", date_field)
    date_field.click()
    time.sleep(0.5)
    date_field.send_keys(DELIVERY_DATE)
    driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles:true}));", date_field)
    time.sleep(1)
    print("  Date set")
except Exception as e:
    print(f"  Date issue: {e}")
    screenshot("04b_date_error")

truck_filter = CONFIG["truck_filter"]
print(f"  Selecting {truck_filter}...")
try:
    truck_input = driver.find_element(By.CSS_SELECTOR, "input[aria-controls='Truck_listbox']")
    js_click(truck_input)
    time.sleep(2)
    screenshot("05_truck_dropdown")
    truck_option = wait.until(EC.presence_of_element_located(
        (By.XPATH, f"//ul[@id='Truck_listbox']//span[@class='k-list-item-text' and contains(text(),'{truck_filter}')]")
    ))
    js_click(truck_option)
    time.sleep(1)
    print(f"  {truck_filter} selected")
except Exception as e:
    print(f"  Truck issue: {e}")
    screenshot("05_truck_error")

print("  Clicking export...")
try:
    export_btn = driver.find_element(By.XPATH, "//button[@onclick='batchPrintExcel()']")
    js_click(export_btn)
    batch_file = wait_for_download('bulk-invoice', timeout=30)
    if not batch_file:
        batch_file = wait_for_download('.xlsx', timeout=15)
except Exception as e:
    print(f"  Export issue: {e}")
    screenshot("07_export_error")


# ══════════════════════════════════════════════════════════════════
# STEP 4 — MODEL INVENTORY
# ══════════════════════════════════════════════════════════════════
print("\nNavigating to inventory...")
driver.get(f"{BASE_URL}/inventory/model")
time.sleep(5)
screenshot("08_inventory_page")

print("  Clicking inventory export...")
try:
    inv_btn = driver.find_element(By.XPATH, "//button[@onclick=\"exportInventory('/inventory/model/export')\"]")
    js_click(inv_btn)
    inv_file = wait_for_download('model-inventory', timeout=30)
    if not inv_file:
        inv_file = wait_for_download('.csv', timeout=15)
except Exception as e:
    print(f"  Inventory issue: {e}")
    screenshot("08_inventory_error")


# ══════════════════════════════════════════════════════════════════
# STEP 4b — SERIAL INVENTORY
# ══════════════════════════════════════════════════════════════════
print("\nExporting serial inventory...")
driver.get(f"{BASE_URL}/inventory/serial")
time.sleep(4)
screenshot("08b_serial_page")

try:
    # No location filter — export all locations to catch allocated units anywhere
    export_btn = driver.find_element(By.XPATH, "//i[contains(@class,'fa-file-excel-o')]/..")
    js_click(export_btn)
    serial_file = wait_for_download('serial-number-inventory', timeout=45)
    if serial_file:
        print(f"  Downloaded: {serial_file}")
    else:
        print("  Serial inventory download timed out")
except Exception as e:
    print(f"  Serial inventory issue: {e}")
    screenshot("08b_serial_error")


# ══════════════════════════════════════════════════════════════════
# STEP 5 — ORDERS DETAIL
# ══════════════════════════════════════════════════════════════════
print("\nNavigating to orders detail export...")
driver.get(f"{BASE_URL}/sales/orders")
time.sleep(4)
screenshot("09_orders_page")

print("  Clicking Open Orders tab...")
try:
    open_orders_tab = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//a[@data-toggle='tab' and contains(text(),'Open Orders')]")
    ))
    js_click(open_orders_tab)
    time.sleep(2)
    print("  Open Orders tab selected")
except Exception as e:
    print(f"  Tab issue: {e}")

print("  Setting Date Field to Estimated Delivery...")
try:
    date_type = driver.find_element(By.NAME, "date-type")
    Select(date_type).select_by_value("EstimatedDeliveryDate")
    time.sleep(1)
    print("  Date field set")
except Exception as e:
    print(f"  Date field issue: {e}")

print(f"  Setting date range to {DELIVERY_DATE}...")
try:
    d = datetime.strptime(DELIVERY_DATE, "%B %d, %Y")
    month = d.month - 1
    day   = d.day
    year  = d.year

    date_input = driver.find_element(By.NAME, "dates")
    js_click(date_input)
    time.sleep(2)

    driver.execute_script(f"""
        var el = $('input[name="dates"]');
        if (el.data('daterangepicker')) {{
            var d = new Date({year}, {month}, {day});
            el.data('daterangepicker').setStartDate(d);
            el.data('daterangepicker').setEndDate(d);
            el.data('daterangepicker').updateElement();
        }}
    """)
    time.sleep(1)

    apply_btn = driver.find_element(By.XPATH, "//button[contains(@class,'applyBtn')]")
    js_click(apply_btn)
    time.sleep(2)

    val = driver.execute_script("return $('input[name=\"dates\"]').val()")
    print(f"  Date range set: {val}")

except Exception as e:
    print(f"  Date range issue: {e}")

print("  Clicking download icon...")
try:
    download_icon = driver.find_element(
        By.XPATH, "//span[contains(@class,'fa-download')]/.."
    )
    js_click(download_icon)
    time.sleep(1)
    screenshot("10_download_menu")

    # Click Detail button
    detail_btn = wait.until(EC.presence_of_element_located(
        (By.ID, "detail-btn")
    ))
    js_click(detail_btn)
    time.sleep(2)
    screenshot("11_export_popup")

    # Click Export button in popup
    export_btn = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//button[contains(@class,'k-button-solid-primary') and contains(text(),'Export')]")
    ))
    js_click(export_btn)
    orders_file = wait_for_download('orders-detail', timeout=30)
    if not orders_file:
        orders_file = wait_for_download('orders', timeout=15)
except Exception as e:
    print(f"  Orders detail export issue: {e}")
    screenshot("10_orders_error")

driver.quit()


# ══════════════════════════════════════════════════════════════════
# SUMMARY
# ══════════════════════════════════════════════════════════════════
print("\n" + "=" * 55)
print("  EXPORT SUMMARY")
print("=" * 55)

all_files    = os.listdir(DOWNLOAD_DIR)
xlsx_files   = [f for f in all_files if f.endswith('.xlsx') and 'FlatFile' not in f]
csv_files    = [f for f in all_files if f.endswith('.csv') and 'orders-detail' not in f and 'serial' not in f]
serial_files = [f for f in all_files if 'serial-number-inventory' in f]
orders_files = [f for f in all_files if 'orders-detail' in f]

print(f"\n  Batch invoice    : {xlsx_files}")
print(f"  Model inventory  : {csv_files}")
print(f"  Serial inventory : {serial_files}")
print(f"  Orders detail    : {orders_files}")

if xlsx_files and csv_files and orders_files and serial_files:
    print("\n  ✅ All 4 files ready — run flat file generator next")
elif xlsx_files and csv_files:
    print("\n  ⚠️  Some files missing — check output above")
else:
    print("\n  Missing files — check screenshots above")
