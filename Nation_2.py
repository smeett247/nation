import logging
import os
import time
from datetime import datetime
from typing import Tuple, List

import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
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

DOWNLOAD_FILENAME = "contracts-flow.xlsx"


def get_element(
        driver: webdriver.Chrome,
        identifier: Tuple[By, str],
        time: int = 60,
        multiple: bool = False,
        presence_of_element_located: bool = False
):
    if not presence_of_element_located:
        if multiple:
            return WebDriverWait(driver, time).until(
                EC.visibility_of_all_elements_located(identifier)
            )
        return WebDriverWait(driver, time).until(
            EC.visibility_of_element_located(identifier)
        )

    return WebDriverWait(driver, time).until(
        EC.presence_of_element_located(identifier)
    )


def get_selenium_driver_path():
    sleep_time_sec = 5
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


def login(driver: webdriver.Chrome) -> None:
    logger.info("Logging in...")

    get_element(driver, (By.XPATH, "//a[@href='/login']")).click()
    get_element(driver, (By.XPATH, "//input[@id='exampleInputEmail1']")).send_keys("ecantor@juntocap.com")
    get_element(driver, (By.XPATH, "//input[@id='exampleInputPassword1']")).send_keys("Brolake1912!")
    get_element(driver, (By.XPATH, "//button[@type='submit']")).click()

    logger.info("Login successful.")


def navigate_to_yoy_comparisons(driver: webdriver.Chrome) -> None:
    logger.info("Navigating to YOY Comparisons page...")

    get_element(driver, (By.XPATH, "//button[@id='dropdownMenuButton1']")).click()
    get_element(driver, (By.XPATH, "//a[@href='/reports']")).click()

    get_element(driver, (By.XPATH, "//a[contains(text(),'Market Research')]")).click()
    get_element(driver, (By.XPATH, "//a[contains(text(),'YOY Comparisons')]")).click()

    logger.info("Reached YOY Comparisons page.")


def checkbox_input(driver: webdriver.Chrome, word: str) -> None:
    time.sleep(1)

    checkbox = get_element(
        driver,
        (By.XPATH, f"//a[@title='{word}']/preceding-sibling::input"),
        presence_of_element_located=True
    )

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", checkbox)
    time.sleep(0.3)

    ActionChains(driver).move_to_element(checkbox).click().perform()
    time.sleep(1)


def download_excel_file(driver: webdriver.Chrome) -> None:
    button = get_element(driver, (By.XPATH, "//button[@id='download']"), presence_of_element_located=True)
    driver.execute_script("arguments[0].click();", button)
    time.sleep(3)

    cross_button = get_element(driver, (By.XPATH, "//span[contains(text(),'Crosstab')]"))
    driver.execute_script("arguments[0].click();", cross_button)

    get_element(driver, (By.XPATH, "//button[contains(text(),'Download')]")).click()
    time.sleep(2)


def wait_for_download(download_dir: str, filename: str, timeout: int = 120) -> str:
    file_path = os.path.join(download_dir, filename)
    temp_path = file_path + ".crdownload"

    start = time.time()
    while time.time() - start < timeout:
        if os.path.exists(temp_path):
            time.sleep(1)
            continue

        if os.path.exists(file_path):
            return file_path

        time.sleep(1)

    raise Exception(f"Download not completed within {timeout} seconds: {filename}")


def get_excel_df(file_path: str) -> pd.DataFrame:
    df = pd.read_excel(file_path)
    df.columns = df.iloc[0]
    df.drop(0, inplace=True)

    dfs: List[pd.DataFrame] = []

    for _, row in df.iterrows():
        temp_df = pd.DataFrame([row])
        temp_df.columns = [f"{col} {temp_df.iloc[0, 0]}" for col in temp_df.columns]
        temp_df = temp_df[temp_df.columns[1:]]
        temp_df.reset_index(drop=True, inplace=True)
        dfs.append(temp_df)

    df = pd.concat(dfs, axis=1)
    df = pd.melt(df)
    df.rename(columns={"variable": "DATE", "value": "VALUE"}, inplace=True)
    df["DATE"] = pd.to_datetime(df["DATE"]) + pd.offsets.MonthEnd(1)

    return df


def open_funding_dropdown(driver: webdriver.Chrome):
    funding_dropdown_xpath = (
        "//div[@class='TitleAndControls CF2Button HideControls'][contains(normalize-space(.),'Funding Agency')]"
        "/following-sibling::div[2]/span/div[@class='tabComboBoxNameContainer tab-ctrl-formatted-fixedsize']"
    )
    funding_dropdown = get_element(driver, (By.XPATH, funding_dropdown_xpath), presence_of_element_located=True)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", funding_dropdown)
    time.sleep(1)
    funding_dropdown.click()
    time.sleep(2)


def open_supplier_dropdown(driver: webdriver.Chrome):
    supplier_dropdown_xpath = (
        "//div[@class='TitleAndControls CF2Button HideControls'][contains(normalize-space(.),'Supplier')]"
        "/following-sibling::div[2]/span/div[@class='tabComboBoxNameContainer tab-ctrl-formatted-fixedsize']"
    )
    supplier_dropdown = get_element(driver, (By.XPATH, supplier_dropdown_xpath), presence_of_element_located=True)
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", supplier_dropdown)
    time.sleep(1)
    supplier_dropdown.click()
    time.sleep(2)


def get_all_funding_agencies(driver: webdriver.Chrome) -> List[str]:
    open_funding_dropdown(driver)

    # uncheck (All)
    checkbox_input(driver, "(All)")
    time.sleep(1)

    # fetch agency list
    agency_links = driver.find_elements(By.XPATH, "//div[contains(@class,'facetOverflow')]//a[@title]")
    agencies = [a.get_attribute("title") for a in agency_links if a.get_attribute("title") and a.get_attribute("title") != "(All)"]

    ActionChains(driver).send_keys(Keys.ESCAPE).perform()
    time.sleep(1)

    return agencies


def get_nation_analytic_df() -> pd.DataFrame:
    driver = get_driver()
    login(driver)
    navigate_to_yoy_comparisons(driver)

    iframe = get_element(driver, (By.XPATH, "//iframe[@title='Data Visualization']"))
    driver.switch_to.frame(iframe)
    logger.info("Switched to iframe.")

    # Month selection
    chart_view_dropdown = get_element(driver, (By.XPATH, "//div/span[contains(text(),'Quarter')]/parent::*"),
                                      presence_of_element_located=True)
    chart_view_dropdown.click()
    time.sleep(1)

    month = get_element(driver, (By.XPATH, "//span[@class='tabMenuItemName'][contains(text(),'Month')]"),
                        presence_of_element_located=True)
    month.click()
    time.sleep(5)

    download_dir = os.path.join(str(Path.home()), "Downloads")

    agencies = get_all_funding_agencies(driver)
    logger.info(f"Total agencies found: {len(agencies)}")

    final_dfs = []

    # ===============================
    # FUNDING AGENCY × SUPPLIER LOOP
    # ===============================
    for agency in agencies:
        logger.info(f"====== Funding Agency: {agency} ======")

        open_funding_dropdown(driver)
        checkbox_input(driver, agency)
        time.sleep(5)

        # Close funding dropdown
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(1)

        # Supplier dropdown open once per agency
        open_supplier_dropdown(driver)
        checkbox_input(driver, "(All)")
        time.sleep(2)

        for company in COMPANIES:
            company_name, ticker = company.split("-")
            company_name = company_name.strip()
            ticker = ticker.strip()

            logger.info(f"Processing: {company_name} under {agency}")

            input_area = get_element(driver, (By.XPATH, "//textarea[@class='QueryBox']"))
            input_area.clear()
            input_area.send_keys(company_name)
            time.sleep(1)

            checkbox_input(driver, company_name)
            time.sleep(5)

            driver.execute_script("window.scrollBy(0, 2000);")
            time.sleep(2)

            download_excel_file(driver)

            downloaded_file_path = wait_for_download(download_dir, DOWNLOAD_FILENAME, timeout=120)

            df = get_excel_df(downloaded_file_path)
            df["TICKER"] = ticker
            df["COMPANY"] = company_name
            df["FUNDING_AGENCY"] = agency
            df["FIELD"] = "Federal Obligations PIT"

            final_dfs.append(df)

            os.remove(downloaded_file_path)

            # uncheck company
            driver.execute_script("window.scrollTo(0, 0);")
            time.sleep(2)
            open_supplier_dropdown(driver)
            checkbox_input(driver, company_name)
            time.sleep(2)

        # Uncheck funding agency
        open_funding_dropdown(driver)
        checkbox_input(driver, agency)
        time.sleep(2)
        ActionChains(driver).send_keys(Keys.ESCAPE).perform()
        time.sleep(2)

    driver.quit()

    final_df = pd.concat(final_dfs, ignore_index=True)

    # format date like your example
    final_df["DATE"] = pd.to_datetime(final_df["DATE"]).dt.strftime("%m/%d/%Y")

    final_df = final_df[["DATE", "TICKER", "COMPANY", "FUNDING_AGENCY", "FIELD", "VALUE"]]
    return final_df


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    df = get_nation_analytic_df()
    print(df.head())
    print("Total rows:", len(df))
