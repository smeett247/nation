import logging
import os
import time
from datetime import datetime
from typing import Tuple, List
import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path

from webdriver_manager.chrome import ChromeDriverManager


logger = logging.getLogger(__name__)

COMPANIES = [
    "Booz Allen Hamilton - BAH",
    "CACI - CACI",
    "Leidos - LDOS",
    "PARSONS - PSN",
    "Tetra Tech - TTEK"
]


# =========================================================
# Utility function to fetch a Selenium web element
# =========================================================
def get_element(
        driver: webdriver.Chrome,
        identifier: Tuple[By, str],
        time: int = 10,
        multiple: bool = False,
        presence_of_element_located: bool = False
):
    if presence_of_element_located:
        if multiple:
            element = WebDriverWait(driver, time).until(
                EC.presence_of_all_elements_located(identifier)
            )
        else:
            element = WebDriverWait(driver, time).until(
                EC.presence_of_element_located(identifier)
            )
    else:
        if multiple:
            element = WebDriverWait(driver, time).until(
                EC.visibility_of_all_elements_located(identifier)
            )
        else:
            element = WebDriverWait(driver, time).until(
                EC.visibility_of_element_located(identifier)
            )

    return element


def get_clickable_element(driver: webdriver.Chrome, identifier: Tuple[By, str], time: int = 20):
    """
    Tableau fix: element_to_be_clickable is unreliable due to overlays.
    We wait for presence first, then attempt clickable quickly.
    """
    el = WebDriverWait(driver, time).until(EC.presence_of_element_located(identifier))
    try:
        WebDriverWait(driver, 3).until(EC.element_to_be_clickable(identifier))
    except Exception:
        pass
    return el


def close_any_open_dropdowns(driver: webdriver.Chrome) -> None:
    """
    Closes any open dropdowns/overlays (Supplier/Funding/Search).
    """
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(1)
    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(1)


# =========================================================
# Selenium Driver Path Function
# =========================================================
def get_selenium_driver_path():
    sleep_time_sec = 30
    retry_count = 5

    while retry_count > 0:
        try:
            chrome_driver_path = os.path.join(
                os.path.dirname(ChromeDriverManager().install()),
                "chromedriver.exe"
            )
        except Exception as e:
            logger.error(f"Error installing ChromeDriver: {e}")
            retry_count -= 1

            if retry_count == 0:
                raise RuntimeError("Failed to install ChromeDriver after multiple attempts")

            time.sleep(sleep_time_sec)
        else:
            return chrome_driver_path


def get_driver() -> webdriver.Chrome:
    chrome_driver_path = get_selenium_driver_path()
    service = Service(chrome_driver_path)

    driver = webdriver.Chrome(service=service)
    driver.get("https://nationanalytics.com/")
    driver.maximize_window()

    logger.info("Driver initialized.")
    return driver


# =========================================================
# Login
# =========================================================
def login(driver: webdriver.Chrome) -> None:
    logger.info("Logging in to Nation Analytics...")

    get_element(driver, (By.XPATH, "//a[@href='/login']")).click()

    get_element(driver, (By.XPATH, "//input[@id='exampleInputEmail1']")) \
        .send_keys("ecantor@juntocap.com")

    get_element(driver, (By.XPATH, "//input[@id='exampleInputPassword1']")) \
        .send_keys("Brolake1912!")

    get_element(driver, (By.XPATH, "//button[@type='submit']")).click()

    logger.info("Login successful.")


def checkbox_input(driver: webdriver.Chrome, word: str) -> None:
    time.sleep(1)

    logger.info(f"Selecting item: {word}")

    checkbox = get_element(
        driver,
        (By.XPATH, f"//a[@title='{word}']/preceding-sibling::input"),
        presence_of_element_located=True
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
    time.sleep(0.5)

    try:
        ActionChains(driver).move_to_element(checkbox).click().perform()
    except Exception:
        driver.execute_script("arguments[0].click();", checkbox)

    logger.info("Checkbox clicked.")


def checkbox_select(driver: webdriver.Chrome, word: str) -> None:
    time.sleep(1)

    logger.info(f"Ensuring item is selected: {word}")

    checkbox = get_element(
        driver,
        (By.XPATH, f"//a[@title='{word}']/preceding-sibling::input"),
        presence_of_element_located=True
    )

    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", checkbox)
    time.sleep(0.5)

    if not checkbox.is_selected():
        try:
            ActionChains(driver).move_to_element(checkbox).click().perform()
        except Exception:
            driver.execute_script("arguments[0].click();", checkbox)

        logger.info("Checkbox selected.")
    else:
        logger.info("Already selected, skipping.")


def download_excel_file(driver: webdriver.Chrome) -> None:
    logger.info("Downloading Excel file...")

    button = get_element(
        driver,
        (By.XPATH, "//button[@id='download']"),
        presence_of_element_located=True
    )
    driver.execute_script("arguments[0].click();", button)

    time.sleep(5)

    cross_button = get_element(driver, (By.XPATH, "//span[contains(text(),'Crosstab')]"))
    driver.execute_script("arguments[0].click();", cross_button)

    time.sleep(2)

    download_button = get_element(driver, (By.XPATH, "//button[contains(text(),'Download')]"))
    driver.execute_script("arguments[0].click();", download_button)

    logger.info("Download initiated.")


# =========================================================
# Read Excel
# =========================================================
def get_excel_df(file_path: str) -> pd.DataFrame:
    logger.info("Transforming Excel file...")

    df = pd.read_excel(file_path)
    df.columns = df.iloc[0]
    df.drop(0, inplace=True)

    dfs: List[pd.DataFrame] = []

    for _, row in df.iterrows():
        temp_df = pd.DataFrame([row])

        temp_df.columns = [
            f"{col} {temp_df.iloc[0, 0]}" for col in temp_df.columns
        ]

        temp_df = temp_df[temp_df.columns[1:]]
        temp_df.reset_index(drop=True, inplace=True)
        dfs.append(temp_df)

    df = pd.concat(dfs, axis=1)
    df = pd.melt(df)
    df.rename(columns={"variable": "date"}, inplace=True)
    df["date"] = pd.to_datetime(df["date"]) + pd.offsets.MonthEnd(1)

    return df


def wait_for_file_download(file_path: str, timeout: int = 60) -> bool:
    start_time = time.time()
    file_dir = os.path.dirname(file_path)
    file_name = os.path.basename(file_path)
    partial_name = file_name + ".crdownload"

    while time.time() - start_time < timeout:
        if os.path.exists(file_path):
            partial_path = os.path.join(file_dir, partial_name)
            if not os.path.exists(partial_path):
                time.sleep(2)
                return True
        time.sleep(2)
    return False


def clear_stale_downloads(directory: str, pattern: str = "contracts-flow"):
    import glob
    files = glob.glob(os.path.join(directory, f"{pattern}*.xlsx"))
    for f in files:
        try:
            os.remove(f)
            logger.info(f"Removed stale file: {f}")
        except Exception as e:
            logger.warning(f"Could not remove stale file {f}: {e}")


# =========================================================
# Navigation
# =========================================================
def navigate_to_yoy_comparisons(driver: WebDriver) -> None:
    logger.info("Navigating to YOY Comparisons page...")

    get_element(driver, (By.XPATH, "//button[@id='dropdownMenuButton1']")).click()
    get_element(driver, (By.XPATH, "//a[@href='/reports']")).click()

    get_element(driver, (By.XPATH, "//a[contains(text(),'Market Research')]")).click()
    get_element(driver, (By.XPATH, "//a[contains(text(),'YOY Comparisons')]")).click()

    logger.info("Navigated to YOY Comparisons page.")


# =========================================================
# FUNDING AGENCY LOGIC (UI ONE-BY-ONE)
# =========================================================
def process_funding_agency(driver: webdriver.Chrome, processed_agencies: set) -> str:
    """
    Select funding agency one-by-one from UI dropdown (NOT hardcoded).
    Handles long list using search-based selection.
    """

    funding_dropdown_xpath = (
        "//div[@class='TitleAndControls CF2Button HideControls'][contains(normalize-space(.),'Funding Agency')]"
        "/following-sibling::div[2]/span/div[@class='tabComboBoxNameContainer tab-ctrl-formatted-fixedsize']"
    )

    retry = 8
    while retry > 0:
        try:
            close_any_open_dropdowns(driver)

            dropdown = get_clickable_element(driver, (By.XPATH, funding_dropdown_xpath), time=30)
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", dropdown)
            time.sleep(1)

            try:
                dropdown.click()
            except Exception:
                driver.execute_script("arguments[0].click();", dropdown)

            time.sleep(2)

            options = get_element(
                driver,
                (By.XPATH, "//a[@title]"),
                multiple=True,
                presence_of_element_located=True,
                time=25
            )

            available_agencies = []
            for opt in options:
                title = opt.get_attribute("title")
                if title:
                    cleaned = title.strip()
                    if cleaned and cleaned not in ["(All)", "Null"] and cleaned not in processed_agencies:
                        available_agencies.append(cleaned)

            if not available_agencies:
                close_any_open_dropdowns(driver)
                return ""

            next_agency = available_agencies[0]
            logger.info(f"Next agency from UI: {next_agency}")

            input_area = get_element(
                driver,
                (By.XPATH, "//textarea[@class='QueryBox']"),
                presence_of_element_located=True,
                time=25
            )
            input_area.clear()
            input_area.send_keys(next_agency)
            time.sleep(2)

            checkbox_select(driver, next_agency)
            time.sleep(1)

            try:
                checkbox_input(driver, "(All)")
            except Exception:
                pass

            time.sleep(2)

            close_any_open_dropdowns(driver)

            processed_agencies.add(next_agency)
            logger.info(f"Funding agency '{next_agency}' selected successfully")
            return next_agency

        except Exception as e:
            retry -= 1
            logger.warning(f"Funding agency selection failed. retries left={retry}. error={e}")

            try:
                driver.switch_to.default_content()
                time.sleep(1)
                iframe = WebDriverWait(driver, 25).until(
                    EC.presence_of_element_located((By.XPATH, "//iframe[@title='Data Visualization']"))
                )
                driver.switch_to.frame(iframe)
                time.sleep(2)
            except Exception:
                pass

            time.sleep(2)

    raise Exception("Funding agency selection failed permanently.")


# =========================================================
# Main DF Function
# =========================================================
def get_nation_analytic_df() -> pd.DataFrame:
    driver = get_driver()

    login(driver)
    navigate_to_yoy_comparisons(driver)

    iframe = get_element(driver, (By.XPATH, "//iframe[@title='Data Visualization']"))
    driver.switch_to.frame(iframe)
    logger.info("Switched to iframe.")

    chart_view_dropdown = get_element(
        driver,
        (By.XPATH, "//div/span[contains(text(),'Quarter')]/parent::*"),
        presence_of_element_located=True
    )
    chart_view_dropdown.click()

    month = get_element(
        driver,
        (By.XPATH, "//span[@class='tabMenuItemName'][contains(text(),'Month')]"),
        presence_of_element_located=True
    )
    month.click()

    logger.info("selected month view button")
    time.sleep(5)

    all_agency_dfs = []
    processed_agencies = set()

    supplier_dropdown_xpath = (
        "//div[@class='TitleAndControls CF2Button HideControls'][contains(normalize-space(.),'Supplier')]"
        "/following-sibling::div[2]/span/div[@class='tabComboBoxNameContainer tab-ctrl-formatted-fixedsize']"
    )

    while True:
        agency = process_funding_agency(driver, processed_agencies)
        if not agency:
            logger.info("All funding agencies processed.")
            break

        logger.info("opening supplier dropdown")
        close_any_open_dropdowns(driver)

        supplier_dropdown = get_clickable_element(driver, (By.XPATH, supplier_dropdown_xpath), time=30)
        driver.execute_script("arguments[0].scrollIntoView({block:'center'});", supplier_dropdown)
        time.sleep(1)

        try:
            supplier_dropdown.click()
        except Exception:
            driver.execute_script("arguments[0].click();", supplier_dropdown)

        time.sleep(2)

        checkbox_input(driver, "(All)")

        for company in COMPANIES:
            logger.info(f"Processing company: {company} for agency: {agency}")

            company_name, ticker = company.split("-")
            company_name = company_name.strip()

            input_area = get_element(driver, (By.XPATH, "//textarea[@class='QueryBox']"))
            input_area.clear()
            input_area.send_keys(company_name)
            time.sleep(1)

            checkbox_input(driver, company_name)

            driver.execute_script("window.scrollBy(0, 2000);")
            time.sleep(1)

            downloads_dir = os.path.join(str(Path.home()), "Downloads")
            downloaded_file_path = os.path.join(downloads_dir, "contracts-flow.xlsx")
            clear_stale_downloads(downloads_dir, "contracts-flow")

            download_excel_file(driver)

            if wait_for_file_download(downloaded_file_path, timeout=60):
                df = get_excel_df(downloaded_file_path)
                df["Ticker"] = ticker.strip()
                df["Company"] = company_name
                df["Funding Agency"] = agency
                all_agency_dfs.append(df)

                os.remove(downloaded_file_path)
            else:
                logger.warning(f"No data / download timeout for company: {company_name}. Skipping...")
                close_any_open_dropdowns(driver)

            checkbox_input(driver, company_name)
            time.sleep(2)
            input_area.clear()

        close_any_open_dropdowns(driver)

    driver.close()
    driver.quit()
    logger.info("closing driver")

    if not all_agency_dfs:
        raise Exception("No data collected")

    final_df = pd.concat(all_agency_dfs, ignore_index=True)
    final_df["field"] = "Federal Obligations PIT"
    return final_df


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    final_df = get_nation_analytic_df()

    if not final_df.empty:
        output_file = "Nation_Analytics_Funding_Data.csv"
        final_df.to_csv(output_file, index=False)
        print(f"Data saved to {output_file}")
    else:
        print("No data collected to save.")

    print(final_df.head())
    print("Total Rows:", len(final_df))
