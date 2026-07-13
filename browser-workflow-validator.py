"""
Refactored automation script.

Improvements:
- main() entry point
- Type hints
- Docstrings
- Logging
- Constants
- try/finally cleanup
"""

from __future__ import annotations

import logging
import re
import time
from pathlib import Path
from typing import Optional, Tuple

import openpyxl
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait

# ==========================
# Configuration
# ==========================

WORKBOOK_PATH = Path(r"C:\SEU_CAMINHO\arquivo.xlsx")
SYSTEM_URL = "https://example.com"

SAVE_EVERY = 10
MAX_RETRY = 3
WAIT_TIMEOUT = 20

# Status constants
STATUS_OK = "ok"
STATUS_CHECK = "check"
STATUS_DOCUMENT_MISMATCH = "document mismatch"
STATUS_SEARCH_ERROR = "search_error"
STATUS_ERROR = "error"

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
)


def clean_document(document: Optional[str]) -> str:
    """Remove non-numeric characters from a document."""
    if not document:
        return ""
    return re.sub(r"\D", "", str(document)).strip()


def normalize_chassis(chassis: Optional[str]) -> str:
    """Normalize chassis value."""
    return "" if not chassis else str(chassis).strip()


def compare_documents(doc1: Optional[str], doc2: Optional[str]) -> bool:
    """Compare documents ignoring formatting and leading zeros."""
    first = clean_document(doc1)
    second = clean_document(doc2)
    return first.endswith(second) or second.endswith(first)


def start_driver() -> Tuple[webdriver.Chrome, WebDriverWait]:
    """Connect to an existing Chrome debugging session."""
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

    driver = webdriver.Chrome(options=options)
    return driver, WebDriverWait(driver, WAIT_TIMEOUT)


def reset_search(driver: webdriver.Chrome, wait: WebDriverWait) -> None:
    """
    Return the application to the advanced search screen.

    NOTE:
    The element identifiers below are placeholders. Replace them with the
    identifiers used by your target application.
    """
    wait.until(
        EC.element_to_be_clickable((By.ID, "SEARCH_MENU_BUTTON"))
    ).click()

    wait.until(
        EC.element_to_be_clickable((By.ID, "ADVANCED_SEARCH_BUTTON"))
    ).click()

    wait.until(
        EC.presence_of_element_located((By.ID, "SEARCH_INPUT"))
    )

    Select(
        wait.until(
            EC.presence_of_element_located((By.ID, "SEARCH_CRITERIA"))
        )
    ).select_by_value("CH")


def search_chassis(
    driver: webdriver.Chrome,
    wait: WebDriverWait,
    chassis: str,
) -> Tuple[str, Optional[str], Optional[str]]:
    """
    Search a chassis and return (status, situation, document).

    NOTE:
    Element IDs and XPaths are generic placeholders for portfolio purposes.
    Replace them with the identifiers from your own application.
    """

    for _ in range(3):
        try:
            reset_search(driver, wait)

            field = wait.until(
                EC.element_to_be_clickable((By.ID, "SEARCH_INPUT"))
            )
            button = wait.until(
                EC.element_to_be_clickable((By.ID, "SEARCH_BUTTON"))
            )

            field.clear()
            field.send_keys(chassis)
            button.click()

            try:
                error = driver.find_element(
                    By.ID,
                    "ERROR_LABEL",
                ).text

                if "No records found" in error:
                    time.sleep(1)
                    continue

            except NoSuchElementException:
                pass

            wait.until(
                EC.presence_of_element_located((By.ID, "RESULT_LABEL"))
            )

            result = driver.find_element(
                By.ID,
                "RESULT_LABEL",
            ).text

            if "1 record" not in result:
                return STATUS_CHECK, None, None

            row = wait.until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//table[@id='RESULT_TABLE']//tr[2]",
                    )
                )
            )

            situation = row.find_element(By.XPATH, ".//td[6]").text.strip()
            site_document = clean_document(
                row.find_element(By.XPATH, ".//td[4]").text
            )

            return STATUS_OK, situation, site_document

        except Exception:
            time.sleep(1)

    return STATUS_SEARCH_ERROR, None, None


def main() -> None:
    """Program entry point."""

    workbook = openpyxl.load_workbook(WORKBOOK_PATH)
    sheet = workbook.active

    driver, wait = start_driver()

    try:
        driver.get(SYSTEM_URL)

        for row in range(2, sheet.max_row + 1):

            current_status = sheet[f"H{row}"].value

            if current_status in {
                STATUS_OK,
                STATUS_CHECK,
                STATUS_DOCUMENT_MISMATCH,
                STATUS_SEARCH_ERROR,
            }:
                continue

            chassis = normalize_chassis(sheet[f"C{row}"].value)
            spreadsheet_doc = clean_document(sheet[f"G{row}"].value)

            if not chassis:
                continue

            for attempt in range(1, MAX_RETRY + 1):

                try:
                    status, situation, site_doc = search_chassis(
                        driver,
                        wait,
                        chassis,
                    )

                    if status == STATUS_SEARCH_ERROR:
                        sheet[f"H{row}"] = STATUS_SEARCH_ERROR
                        break

                    if (
                        status == STATUS_CHECK
                        or situation != "Q00"
                    ):
                        sheet[f"H{row}"] = STATUS_CHECK

                    elif not compare_documents(site_doc, spreadsheet_doc):
                        sheet[f"H{row}"] = STATUS_DOCUMENT_MISMATCH

                    else:
                        sheet[f"H{row}"] = STATUS_OK

                    break

                except Exception as exc:
                    logging.warning(
                        "Row %s | Attempt %s | %s",
                        row,
                        attempt,
                        exc,
                    )
                    time.sleep(2)

            else:
                sheet[f"H{row}"] = STATUS_ERROR

            if row % SAVE_EVERY == 0:
                workbook.save(WORKBOOK_PATH)
                logging.info("Workbook saved (row %s).", row)

    finally:
        workbook.save(WORKBOOK_PATH)
        driver.quit()

    logging.info("Process finished successfully.")


if __name__ == "__main__":
    main()
