import logging
import os
import time
from datetime import datetime
from typing import Union, Tuple, List
import pandas as pd
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
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
        time: int = 60,
        multiple: bool = False,
        presence_of_element_located: bool = False
):
    """
    Waits for and returns a Selenium web element.

    :param driver: Selenium WebDriver instance
    :param identifier: Tuple of (By, locator)
    :param time: Maximum wait time in seconds
    :param multiple: Whether to return multiple elements
    :param presence_of_element_located: Wait only for presence (not visibility)
    :return: WebElement or list of WebElements
    """

    if not presence_of_element_located:
        if multiple:
            element = WebDriverWait(driver, time).until(
                EC.visibility_of_all_elements_located(identifier)
            )
        else:
            element = WebDriverWait(driver, time).until(
                EC.visibility_of_element_located(identifier)
            )
    else:
        element = WebDriverWait(driver, time).until(
            EC.presence_of_element_located(identifier)
        )

    return element


# =========================================================
# NEW: Selenium Driver Path Function
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
    """
    Initialize and return a Chrome WebDriver instance.
    """
    chrome_driver_path = get_selenium_driver_path()
    service = Service(chrome_driver_path)

    driver = webdriver.Chrome(service=service)

    # Open target website
    driver.get("https://nationanalytics.com/")

    # Maximize browser window
    driver.maximize_window()

    logger.info("Driver initialized.")
    return driver


# =========================================================
# Perform login to Nation Analytics
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
    time.sleep(3)

    logger.info(f"Selecting company: {word}")

    checkbox = get_element(
        driver,
        (By.XPATH, f"//a[@title='{word}']/preceding-sibling::input"),
        presence_of_element_located=True
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block: 'center'});",
        checkbox
    )

    time.sleep(0.5)

    ActionChains(driver).move_to_element(checkbox).click().perform()

    logger.info("Checkbox selected.")


def download_excel_file(driver: webdriver.Chrome) -> None:
    logger.info("Downloading Excel file...")

    button = get_element(
        driver,
        (By.XPATH, "//button[@id='download']"),
        presence_of_element_located=True
    )

    driver.execute_script("arguments[0].click();", button)

    time.sleep(5)

    logger.info("Selecting download option...")

    cross_button = get_element(
        driver,
        (By.XPATH, "//span[contains(text(),'Crosstab')]")
    )
    driver.execute_script("arguments[0].click();", cross_button)

    get_element(
        driver,
        (By.XPATH, "//button[contains(text(),'Download')]")
    ).click()

    logger.info("Download initiated.")


# =========================================================
# Read and transform Excel file into DataFrame
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


def get_file_mode_date(file_path: str) -> str:
    mod_time = os.path.getmtime(file_path)
    mod_date = datetime.fromtimestamp(mod_time).strftime('%Y-%m-%d')
    return mod_date


def navigate_to_yoy_comparisons(driver: WebDriver) -> None:
    logger.info("Navigating to YOY Comparisons page...")

    get_element(
        driver,
        (By.XPATH, "//button[@id='dropdownMenuButton1']")
    ).click()

    get_element(
        driver,
        (By.XPATH, "//a[@href='/reports']")
    ).click()

    logger.info("Reports menu opened.")

    get_element(
        driver,
        (By.XPATH, "//a[contains(text(),'Market Research')]")
    ).click()

    get_element(
        driver,
        (By.XPATH, "//a[contains(text(),'YOY Comparisons')]")
    ).click()

    logger.info("Navigated to YOY Comparisons page.")


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

    logger.info("Clicked on chart view dropdown.")

    month = get_element(
        driver,
        (By.XPATH, "//span[@class='tabMenuItemName'][contains(text(),'Month')]"),
        presence_of_element_located=True
    )
    month.click()

    logger.info("selected month view button")
    time.sleep(5)

    logger.info("opening supplier dropdown")

    get_element(
        driver,
        (By.XPATH,
         "//div[@class='TitleAndControls CF2Button HideControls']/following-sibling::*[2]//div[@class='tabComboBoxNameContainer tab-ctrl-formatted-fixedsize']"),
        presence_of_element_located=True
    ).click()

    checkbox_input(driver, "(All)")

    dfs = []
    for company in COMPANIES:
        logger.info(f"Processing company: {company}")

        company, ticker = company.split("-")
        company = company.strip()

        logger.info(f"searching company name {company}")

        input_area = get_element(driver, (By.XPATH, "//textarea[@class='QueryBox']"))
        input_area.send_keys(company)

        checkbox_input(driver, company)
        driver.execute_script("window.scrollBy(0, 2000);")

        download_excel_file(driver)
        time.sleep(10)

        downloaded_file_path = os.path.join(str(Path.home()) + "\Downloads", "contracts-flow.xlsx")
        today_date = datetime.today().date()

        if os.path.exists(downloaded_file_path) and str(today_date) == get_file_mode_date(downloaded_file_path):
            logger.info("File downloaded successfully.")

            df = get_excel_df(downloaded_file_path)
            df["Ticker"] = ticker.strip()
            df["Company"] = company
            dfs.append(df)

            os.remove(downloaded_file_path)
        else:
            raise Exception("Expected File not found")

        checkbox_input(driver, company)
        time.sleep(3)

        input_area.clear()

    driver.close()
    driver.quit()
    logger.info("closing driver")

    df = pd.concat(dfs, ignore_index=True)
    df["field"] = "Federal Obligations PIT"
    return df


if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s - %(levelname)s - %(message)s"
    )

    final_df = get_nation_analytic_df()
    print(final_df.head())
    print("Total Rows:", len(final_df))
