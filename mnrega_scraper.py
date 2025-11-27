import time
import os
import pandas as pd

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException


def switch_to_new_window(driver, old_handles, timeout=30):
    WebDriverWait(driver, timeout).until(
        lambda d: len(d.window_handles) > len(old_handles)
    )
    new_handle = [h for h in driver.window_handles if h not in old_handles][0]
    driver.switch_to.window(new_handle)


def safe_select(select_id, by_text=True, value=None, driver=None, wait=None):
    for _ in range(5):
        try:
            elem = wait.until(EC.presence_of_element_located((By.ID, select_id)))
            sel = Select(elem)
            if by_text:
                sel.select_by_visible_text(value)
            else:
                sel.select_by_value(value)
            return
        except StaleElementReferenceException:
            time.sleep(1)


def extract_muster_table(driver, work_code, muster_no):
    wait = WebDriverWait(driver, 30)

    heading = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//*[contains(text(),'Mustroll Detail')]")
        )
    )
    table = heading.find_element(By.XPATH, "./following::table[1]")

    rows = table.find_elements(By.TAG_NAME, "tr")
    if not rows:
        return []

    headers = [c.text.strip() for c in rows[0].find_elements(By.XPATH, ".//th|.//td")]
    data = []

    for r in rows[1:]:
        cells = [c.text.strip() for c in r.find_elements(By.TAG_NAME, "td")]
        if not any(cells):
            continue
        while len(cells) < len(headers):
            cells.append("")
        row_dict = dict(zip(headers, cells))
        row_dict["Work_Code"] = work_code
        row_dict["Muster_No"] = muster_no
        data.append(row_dict)

    return data


def run_mnrega_scraper(work_code: str, output_dir: str = ".") -> str:
    """
    Given a WORK_CODE, run MNREGA automation and return path of saved Excel file.
    """

    # Sanitize work_code for filename
    safe_code = work_code.replace("/", "_").replace("\\", "_").replace(" ", "_")
    output_path = os.path.join(output_dir, f"mnrega_{safe_code}.xlsx")

    # Headless Chrome options
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--log-level=3")

    options.binary_location = "/usr/bin/chromium"

    driver = webdriver.Chrome(
    service=Service("/usr/bin/chromedriver"),
    options=options
    )
    wait = WebDriverWait(driver, 40)
    driver.set_window_size(1400, 900)

    all_records = []

    try:
        # STEP 1: HOME SEARCH
        driver.get("https://mnregaweb4.nic.in/netnrega/homesearch.htm")

        iframe = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//iframe[contains(@src,'nregasearch1.aspx')]")
            )
        )
        driver.switch_to.frame(iframe)

        safe_select("ddl_search", by_text=True, value="Work", driver=driver, wait=wait)
        safe_select("ddl_state", by_text=True, value="GUJARAT", driver=driver, wait=wait)
        time.sleep(2)
        safe_select("ddl_district", by_text=False, value="1109", driver=driver, wait=wait)
        time.sleep(2)

        keyword = wait.until(EC.presence_of_element_located((By.ID, "txt_keyword2")))
        driver.execute_script("arguments[0].value = arguments[1];", keyword, work_code)

        before = driver.window_handles[:]
        driver.find_element(By.ID, "btn_go").click()

        # STEP 2: RESULTS WINDOW
        switch_to_new_window(driver, before)

        work_link = wait.until(
            EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, work_code))
        )
        driver.execute_script("arguments[0].click();", work_link)

        # STEP 3: ASSET REGISTER
        wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//td[contains(.,'Distinct Number of Muster Rolls used')]")
            )
        )

        asset_window = driver.current_window_handle

        row = driver.find_element(
            By.XPATH,
            "//td[contains(.,'Distinct Number of Muster Rolls used')]/parent::tr",
        )
        links = row.find_elements(By.TAG_NAME, "a")
        muster_numbers = [a.text.strip() for a in links if a.text.strip()]

        # STEP 4: EACH MUSTER
        for muster_no in muster_numbers:
            driver.switch_to.window(asset_window)

            row = wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//td[contains(.,'Distinct Number of Muster Rolls used')]/parent::tr",
                    )
                )
            )
            link = row.find_element(
                By.XPATH, f".//a[contains(text(), '{muster_no}')]"
            )

            old_handles = driver.window_handles[:]
            driver.execute_script("arguments[0].click();", link)
            time.sleep(2)

            new_window_opened = len(driver.window_handles) > len(old_handles)
            if new_window_opened:
                switch_to_new_window(driver, old_handles)
            # else: opened in same window, stay

            rows = extract_muster_table(driver, work_code, muster_no)
            all_records.extend(rows)

            if new_window_opened:
                driver.close()
                driver.switch_to.window(asset_window)
            else:
                driver.back()
                wait.until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            "//td[contains(.,'Distinct Number of Muster Rolls used')]",
                        )
                    )
                )

        if all_records:
            df = pd.DataFrame(all_records)
            df.to_excel(output_path, index=False)
        else:
            # Still create empty file so user gets something
            pd.DataFrame([]).to_excel(output_path, index=False)

        return output_path

    finally:
        driver.quit()
