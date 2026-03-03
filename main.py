import argparse
import mimetypes
import re
import time
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple
from urllib.parse import quote

from openpyxl import load_workbook
from selenium import webdriver
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

WHATSAPP_WEB_URL = "https://web.whatsapp.com"
DEFAULT_SHEET_NAME = "Data"
STATUS_COL = 5  # E
TIMESTAMP_COL = 6  # F


class WhatsAppAutomationError(Exception):
    """Raised for invalid configuration or runtime WhatsApp automation conditions."""


def now_str() -> str:
    return datetime.now().isoformat(sep=" ", timespec="seconds")


def sanitize_phone_number(raw: object, default_country_code: Optional[str] = None) -> Optional[str]:
    """Return WhatsApp-compatible phone number (digits only) or None if unusable."""
    if raw is None:
        return None

    digits = re.sub(r"\D", "", str(raw))
    if not digits:
        return None

    if default_country_code:
        cc = re.sub(r"\D", "", str(default_country_code))
        if cc and len(digits) <= 10 and not digits.startswith(cc):
            digits = f"{cc}{digits}"

    return digits if len(digits) >= 10 else None


def wait_for_login(driver: webdriver.Chrome, timeout: int = 240) -> None:
    """Wait until WhatsApp Web is ready after QR scan/login."""
    print("Open WhatsApp Web and scan QR if required...")
    WebDriverWait(driver, timeout).until(
        EC.any_of(
            EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Chat list']")),
            EC.presence_of_element_located((By.XPATH, "//div[@role='grid']")),
            EC.presence_of_element_located((By.XPATH, "//div[@id='pane-side']")),
        )
    )
    print("WhatsApp login detected.")


def open_chat(driver: webdriver.Chrome, phone: str, message: str, timeout: int = 60) -> Tuple[bool, str]:
    encoded = quote(message or "")
    driver.get(f"{WHATSAPP_WEB_URL}/send?phone={phone}&text={encoded}")

    invalid_xpaths = [
        "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'not on whatsapp')]",
        "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'phone number shared via url is invalid')]",
        "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'number is invalid')]",
    ]

    try:
        WebDriverWait(driver, timeout).until(
            EC.any_of(
                EC.presence_of_element_located((By.XPATH, "//footer//div[@contenteditable='true']")),
                *[EC.presence_of_element_located((By.XPATH, xp)) for xp in invalid_xpaths],
            )
        )
    except TimeoutException:
        return False, "Timeout opening chat (number may be invalid/non-WhatsApp or network is slow)"

    for xp in invalid_xpaths:
        if driver.find_elements(By.XPATH, xp):
            return False, "Number not on WhatsApp or invalid"

    return True, "Chat opened"


def send_text(driver: webdriver.Chrome, timeout: int = 20) -> Tuple[bool, str]:
    """Send prefilled text by pressing Enter in composer."""
    try:
        input_box = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((By.XPATH, "//footer//div[@contenteditable='true']"))
        )
        input_box.click()
        input_box.send_keys("\n")
        return True, "Text sent"
    except TimeoutException:
        return False, "Text send failed: composer not ready"


def _set_attachment_caption(driver: webdriver.Chrome, caption: str, timeout: int) -> None:
    if not caption:
        return

    caption_box = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located(
            (
                By.XPATH,
                "//div[@role='dialog']//div[@contenteditable='true'][@data-tab or @spellcheck='true']"
                " | //div[@role='dialog']//*[contains(@aria-label,'caption') and @contenteditable='true']",
            )
        )
    )
    caption_box.click()
    caption_box.send_keys(caption)


def send_attachment(driver: webdriver.Chrome, file_path: str, caption: str = "", timeout: int = 45) -> Tuple[bool, str]:
    """Attach and send a local file (image/video/document) with optional caption in one message."""
    if not file_path:
        return True, "No attachment"

    resolved = Path(file_path).expanduser().resolve()
    if not resolved.exists() or not resolved.is_file():
        return False, f"Attachment not found: {resolved}"

    mime_type, _ = mimetypes.guess_type(str(resolved))
    media_file = bool(mime_type and (mime_type.startswith("image/") or mime_type.startswith("video/")))

    try:
        attach_button = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//button[@title='Attach' or @aria-label='Attach' or .//span[@data-icon='plus'] or .//*[contains(@data-testid,'attach')]]",
                )
            )
        )
        attach_button.click()

        input_xpath = (
            "//input[@type='file' and contains(@accept,'image/*')]"
            if media_file
            else "//input[@type='file' and not(contains(@accept,'image/*'))]"
        )

        file_input = WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((By.XPATH, input_xpath))
        )

        try:
            driver.execute_script(
                "arguments[0].style.display='block'; arguments[0].style.visibility='visible'; arguments[0].style.opacity=1;",
                file_input,
            )
            file_input.send_keys(str(resolved))
        except Exception as exc:  # noqa: BLE001
            return False, f"Attachment failed: could not upload file ({exc})"

        WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    "//div[@role='dialog'] | //button[@aria-label='Send'] | //span[@data-icon='send']",
                )
            )
        )

        _set_attachment_caption(driver, caption, timeout)

        send_button = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable(
                (By.XPATH, "//button[@aria-label='Send'] | //span[@data-icon='send']/ancestor::button")
            )
        )
        send_button.click()
        return True, "Attachment sent with caption"
    except TimeoutException:
        return False, "Attachment failed: timeout waiting for attachment UI"


def write_result(sheet, row: int, status: str) -> None:
    sheet.cell(row=row, column=STATUS_COL).value = status
    sheet.cell(row=row, column=TIMESTAMP_COL).value = now_str()


def process_rows(
    excel_path: Path,
    sheet_name: str,
    default_country_code: Optional[str],
    pause_seconds: float,
    cooldown_every: int,
    cooldown_seconds: float,
) -> None:
    workbook = load_workbook(excel_path)
    if sheet_name not in workbook.sheetnames:
        raise WhatsAppAutomationError(f"Sheet '{sheet_name}' not found in {excel_path}")

    sheet = workbook[sheet_name]

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)

    try:
        driver.get(WHATSAPP_WEB_URL)
        wait_for_login(driver)

        for row in range(2, sheet.max_row + 1):
            raw_phone = sheet.cell(row=row, column=1).value
            message = str(sheet.cell(row=row, column=2).value or "")
            attachment = str(sheet.cell(row=row, column=4).value or "").strip()

            phone = sanitize_phone_number(raw_phone, default_country_code)
            if not phone:
                status = "Skipped: invalid phone"
                print(f"Row {row}: {status}")
                write_result(sheet, row, status)
                workbook.save(excel_path)
                continue

            # Important behavior: if attachment exists, message will be sent as attachment caption
            # so text + file go together in one WhatsApp message.
            chat_text = "" if attachment else message
            ok, reason = open_chat(driver, phone, chat_text)
            if not ok:
                status = f"Failed: {reason}"
                print(f"Row {row} ({phone}): {status}")
                write_result(sheet, row, status)
                workbook.save(excel_path)
                continue

            if attachment:
                media_ok, media_reason = send_attachment(driver, attachment, caption=message)
                final_status = "Sent attachment + caption" if media_ok else f"Failed: {media_reason}"
            else:
                txt_ok, txt_reason = send_text(driver)
                final_status = "Sent text" if txt_ok else f"Failed: {txt_reason}"

            print(f"Row {row} ({phone}): {final_status}")
            write_result(sheet, row, final_status)
            workbook.save(excel_path)

            time.sleep(max(0.0, pause_seconds))
            if cooldown_every > 0 and (row - 1) % cooldown_every == 0:
                print(f"Cooldown for {cooldown_seconds} seconds...")
                time.sleep(max(0.0, cooldown_seconds))

        print(f"Finished. Status written to {excel_path} columns E/F.")
    finally:
        workbook.save(excel_path)
        driver.quit()


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Send WhatsApp messages from Excel via WhatsApp Web.")
    parser.add_argument("excel", help="Path to Excel file (.xlsx or .xlsm)")
    parser.add_argument("--sheet", default=DEFAULT_SHEET_NAME, help=f"Sheet name (default: {DEFAULT_SHEET_NAME})")
    parser.add_argument("--default-country-code", default=None, help="Prefix for 10-digit local numbers, e.g. 91")
    parser.add_argument("--pause", type=float, default=3.0, help="Pause between rows in seconds")
    parser.add_argument("--cooldown-every", type=int, default=9, help="Run cooldown every N processed rows (0 to disable)")
    parser.add_argument("--cooldown-seconds", type=float, default=42.0, help="Cooldown duration in seconds")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    excel_path = Path(args.excel).expanduser().resolve()
    if not excel_path.exists():
        raise WhatsAppAutomationError(f"Excel file not found: {excel_path}")

    if excel_path.suffix.lower() not in {".xlsx", ".xlsm"}:
        raise WhatsAppAutomationError("Use .xlsx or .xlsm file")

    process_rows(
        excel_path=excel_path,
        sheet_name=args.sheet,
        default_country_code=args.default_country_code,
        pause_seconds=args.pause,
        cooldown_every=args.cooldown_every,
        cooldown_seconds=args.cooldown_seconds,
    )


if __name__ == "__main__":
    main()
