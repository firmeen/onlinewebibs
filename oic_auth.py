from __future__ import annotations

import base64
import json
import os
import re
import tempfile
import time
from datetime import datetime
from pathlib import Path
from urllib.parse import urlparse

from dotenv import dotenv_values

from selenium import webdriver
from selenium.common.exceptions import (
    NoAlertPresentException,
    StaleElementReferenceException,
    TimeoutException,
    UnexpectedAlertPresentException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


# ============================================================
# INSTALLATION
# ============================================================
# python -m pip install selenium python-dotenv
#
# FILE: oic_auth.py
# RUN : python oic_auth.py
#
# Live OIC execution has not been tested.
# Use only an authorized OIC account.
# ============================================================


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"

BASE_URL = "https://onlinewebibs.oic.or.th"

LOGIN_URL = (
    f"{BASE_URL}/WebForm/Account/IBSG_S_Login.aspx"
)

GENTOKEN_PATH = (
    "/WebForm/Dashboard/Pivot/PivotCube.asmx/GenToken"
)

LOGIN_WAIT = 4
LOGIN_TIMEOUT = 30
ELEMENT_TIMEOUT = 20
NETWORK_TIMEOUT = 60

POLL_INTERVAL = 0.20

MAX_LOGIN_ATTEMPTS = 2
MAX_SESSION_CONFIRMATIONS = 1

HEADLESS = False
CLOSE_BROWSER = False


# ============================================================
# ELEMENT SELECTORS
# ============================================================

USERNAME_ID = "cphBody_txtUserName"
PASSWORD_ID = "cphBody_txtUserPassword"
LOGIN_BUTTON_ID = "cphBody_btnLogin"

PIVOT_TOGGLE_ID = "ucUserMenu_treeViewn8"
PIVOT_MENU_ID = "ucUserMenu_treeViewt8"
PIVOT_CONTAINER_ID = "ucUserMenu_treeViewn8Nodes"

TARGET_CUBE = "MR"

CUBE_OPTIONS = {
    "MR": {
        "id": "ucUserMenu_treeViewt2",
        "name": "ประกันภัยทางทะเลและการขนส่ง",
    },
    "MI": {
        "id": "ucUserMenu_treeViewt3",
        "name": "ประกันภัยเบ็ดเตล็ด",
    },
    "FI": {
        "id": "ucUserMenu_treeViewt4",
        "name": "ประกันภัยทรัพย์สิน",
    },
}


SESSION_ALERT_KEYWORDS = (
    "ล็อคอินจากเครื่องอื่น",
    "ล็อกอินจากเครื่องอื่น",
    "เข้าสู่ระบบจากเครื่องอื่น",
    "ต้องออกจากระบบก่อน",
    "already logged in",
    "already logged on",
    "active session",
    "another computer",
)


JWT_PATTERN = re.compile(
    r"(?<![A-Za-z0-9_-])"
    r"[A-Za-z0-9_-]+\."
    r"[A-Za-z0-9_-]+\."
    r"[A-Za-z0-9_-]+"
    r"(?![A-Za-z0-9_-])"
)


# ============================================================
# TERMINAL
# ============================================================

def log(status: str, message: str) -> None:

    timestamp = datetime.now().strftime("%H:%M:%S")

    print(
        f"{timestamp} [{status:^7}] {message}",
        flush=True,
    )


def print_banner() -> None:

    print()
    print("=" * 76)
    print("                       OIC AUTO AUTHENTICATION")
    print("                 SELENIUM / CDP / SESSION MANAGEMENT")
    print("=" * 76)
    print()


# ============================================================
# ENVIRONMENT
# ============================================================

def load_credentials() -> tuple[str, str]:

    if not ENV_PATH.is_file():

        raise FileNotFoundError(
            ".env file was not found."
        )

    config = dotenv_values(
        ENV_PATH,
        interpolate=False,
        encoding="utf-8-sig",
    )

    username = config.get("USER")
    password = config.get("PASSWORD")

    if not username:

        raise ValueError(
            "USER is missing from .env."
        )

    if not password:

        raise ValueError(
            "PASSWORD is missing from .env."
        )

    log("OK", "Environment loaded")
    log("OK", "Login credentials available")

    return username, password


def quote_env(value: str) -> str:

    value = (
        value
        .replace("\\", "\\\\")
        .replace('"', '\\"')
        .replace("\r", "\\r")
        .replace("\n", "\\n")
    )

    return f'"{value}"'


def update_env(cookie: str, token: str) -> None:

    if not cookie or not token:

        raise ValueError(
            "Cannot update .env with empty authentication values."
        )

    updates = {
        "OIC_COOKIE": cookie,
        "OIC_BEARER_TOKEN": token,
    }

    original_bytes = ENV_PATH.read_bytes()

    has_bom = original_bytes.startswith(
        b"\xef\xbb\xbf"
    )

    original_text = original_bytes.decode(
        "utf-8-sig"
    )

    newline = (
        "\r\n"
        if "\r\n" in original_text
        else "\n"
    )

    lines = original_text.splitlines()

    output = []
    updated = set()

    key_pattern = re.compile(
        r"^\s*(?:#\s*)?(?:export\s+)?"
        r"(OIC_COOKIE|OIC_BEARER_TOKEN)\s*="
    )

    for line in lines:

        match = key_pattern.match(line)

        if not match:

            output.append(line)
            continue

        key = match.group(1)

        if key in updated:
            continue

        output.append(
            f"{key}={quote_env(updates[key])}"
        )

        updated.add(key)

    missing = [
        key
        for key in updates
        if key not in updated
    ]

    if missing and output and output[-1].strip():

        output.append("")

    for key in missing:

        output.append(
            f"{key}={quote_env(updates[key])}"
        )

    final_content = (
        newline.join(output) + newline
    )

    encoding = (
        "utf-8-sig"
        if has_bom
        else "utf-8"
    )

    temp_path = None

    try:

        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding=encoding,
            newline="",
            prefix=".oic_env_",
            suffix=".tmp",
            dir=BASE_DIR,
            delete=False,
        ) as temp_file:

            temp_path = Path(
                temp_file.name
            )

            temp_file.write(final_content)
            temp_file.flush()

            os.fsync(
                temp_file.fileno()
            )

        os.replace(
            temp_path,
            ENV_PATH,
        )

    finally:

        if temp_path and temp_path.exists():

            temp_path.unlink()

    log("OK", "OIC_COOKIE updated")
    log("OK", "OIC_BEARER_TOKEN updated")
    log("OK", ".env saved successfully")


# ============================================================
# SELENIUM DRIVER
# ============================================================

def create_driver() -> webdriver.Chrome:

    options = webdriver.ChromeOptions()

    options.add_experimental_option("detach", True)

    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-extensions")
    options.add_argument("--log-level=3")

    options.page_load_strategy = "normal"

    options.set_capability(
        "unhandledPromptBehavior",
        "ignore",
    )

    options.set_capability(
        "goog:loggingPrefs",
        {
            "performance": "ALL",
        },
    )

    options.add_experimental_option(
        "excludeSwitches",
        ["enable-logging"],
    )

    options.add_experimental_option(
        "prefs",
        {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
        },
    )

    if HEADLESS:

        options.add_argument("--headless=new")
        options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(
        options=options
    )

    driver.set_page_load_timeout(60)

    try:

        driver.execute_cdp_cmd(
            "Network.enable",
            {
                "maxTotalBufferSize": 50 * 1024 * 1024,
                "maxResourceBufferSize": 5 * 1024 * 1024,
            },
        )

    except Exception:

        driver.quit()
        raise

    log("OK", "Chrome launched")
    log("OK", "JavaScript dialog handling enabled")
    log("OK", "CDP network monitoring enabled")

    return driver


# ============================================================
# COOKIE CONSENT
# ============================================================

def accept_cookie_consent(
    driver: webdriver.Chrome,
) -> None:

    selectors = (
        "//button[normalize-space(.)='รับทราบ']",
        "//a[normalize-space(.)='รับทราบ']",
        "//input[@value='รับทราบ']",
        "//*[normalize-space(text())='รับทราบ' and @onclick]",
    )

    for selector in selectors:

        elements = driver.find_elements(
            By.XPATH,
            selector,
        )

        for element in elements:

            try:

                if (
                    element.is_displayed()
                    and element.is_enabled()
                ):

                    element.click()

                    log(
                        "OK",
                        "Cookie consent accepted",
                    )

                    return

            except WebDriverException:

                continue


# ============================================================
# JAVASCRIPT ALERT
# ============================================================

def get_current_alert(
    driver: webdriver.Chrome,
):

    try:

        return driver.switch_to.alert

    except NoAlertPresentException:

        return None


def is_session_conflict(
    message: str,
) -> bool:

    normalized = message.strip().lower()

    return any(
        keyword.lower() in normalized
        for keyword in SESSION_ALERT_KEYWORDS
    )


def handle_login_alert(
    driver: webdriver.Chrome,
    confirmation_count: int,
) -> tuple[bool, int]:

    alert = get_current_alert(driver)

    if alert is None:

        return False, confirmation_count

    message = alert.text.strip()

    log(
        "WARN",
        "JavaScript dialog detected",
    )

    if not is_session_conflict(message):

        alert.dismiss()

        raise RuntimeError(
            "Unexpected login dialog detected. "
            "Authentication stopped."
        )

    log(
        "WARN",
        "Existing active OIC session detected",
    )

    if confirmation_count >= MAX_SESSION_CONFIRMATIONS:

        alert.dismiss()

        raise RuntimeError(
            "Session conflict confirmation limit reached."
        )

    print()
    print("-" * 76)
    print("  OIC SESSION CONFLICT")
    print("-" * 76)
    print(f"  {message}")
    print()
    print("  Continuing may terminate another active session.")
    print("  Confirm only if authorized to terminate that session.")
    print("-" * 76)

    choice = "y"

    log("INFO", "Session conflict automatically confirmed")

    if choice not in ("y", "yes"):

        alert.dismiss()

        raise RuntimeError(
            "Session conflict confirmation cancelled."
        )

    alert = get_current_alert(driver)

    if alert is None:

        raise RuntimeError(
            "Session confirmation dialog is no longer available."
        )

    if not is_session_conflict(alert.text):

        alert.dismiss()

        raise RuntimeError(
            "Session confirmation dialog has changed."
        )

    alert.accept()

    confirmation_count += 1

    print()

    log(
        "OK",
        "Session conflict confirmation accepted",
    )

    log(
        "WAIT",
        f"Waiting {LOGIN_WAIT} seconds after confirmation",
    )

    return True, confirmation_count


# ============================================================
# PAGE STATUS
# ============================================================

def dashboard_is_ready(
    driver: webdriver.Chrome,
) -> bool:

    try:
        if "ibsg_s_home.aspx" not in driver.current_url.lower():
            return False

        if driver.execute_script(
            "return document.readyState;"
        ) != "complete":
            return False

        elements = driver.find_elements(
            By.ID,
            PIVOT_MENU_ID,
        )

        return any(
            element.is_displayed()
            for element in elements
        )

    except StaleElementReferenceException:
        return False

    except UnexpectedAlertPresentException:
        raise


def login_page_is_ready(
    driver: webdriver.Chrome,
) -> bool:

    required_ids = (
        USERNAME_ID,
        PASSWORD_ID,
        LOGIN_BUTTON_ID,
    )

    for element_id in required_ids:

        elements = driver.find_elements(
            By.ID,
            element_id,
        )

        visible = False

        for element in elements:

            try:

                if (
                    element.is_displayed()
                    and element.is_enabled()
                ):

                    visible = True
                    break

            except StaleElementReferenceException:

                continue

        if not visible:

            return False

    return True


def switch_to_dashboard_window(
    driver: webdriver.Chrome,
) -> bool:

    original_window = driver.current_window_handle

    handles = driver.window_handles

    if len(handles) <= 1:

        return dashboard_is_ready(driver)

    for handle in handles:

        try:

            driver.switch_to.window(handle)

            if get_current_alert(driver) is not None:

                return False

            if dashboard_is_ready(driver):

                return True

        except UnexpectedAlertPresentException:

            return False

        except WebDriverException:

            continue

    if original_window in driver.window_handles:

        driver.switch_to.window(
            original_window
        )

    return False


# ============================================================
# LOGIN FORM
# ============================================================

def submit_login(
    driver: webdriver.Chrome,
    username: str,
    password: str,
    attempt: int,
) -> None:

    wait = WebDriverWait(
        driver,
        ELEMENT_TIMEOUT,
    )

    username_element = wait.until(
        EC.visibility_of_element_located(
            (By.ID, USERNAME_ID)
        )
    )

    password_element = wait.until(
        EC.visibility_of_element_located(
            (By.ID, PASSWORD_ID)
        )
    )

    login_button = wait.until(
        EC.element_to_be_clickable(
            (By.ID, LOGIN_BUTTON_ID)
        )
    )

    username_element.clear()
    username_element.send_keys(username)

    password_element.clear()
    password_element.send_keys(password)

    log(
        "OK",
        f"Credentials entered | Attempt {attempt}",
    )

    try:

        login_button.click()

    except UnexpectedAlertPresentException:

        pass

    log(
        "OK",
        f"Login submitted | Attempt {attempt}",
    )


# ============================================================
# WAIT FOR LOGIN RESULT
# ============================================================

def wait_for_login_result(
    driver: webdriver.Chrome,
    confirmation_count: int,
) -> tuple[str, int]:

    started_at = time.monotonic()

    deadline = (
        started_at + LOGIN_TIMEOUT
    )

    minimum_wait_until = (
        started_at + LOGIN_WAIT
    )

    session_confirmed = False

    log(
        "WAIT",
        f"Waiting {LOGIN_WAIT} seconds after login",
    )

    while time.monotonic() < deadline:

        handled, confirmation_count = handle_login_alert(
            driver,
            confirmation_count,
        )

        if handled:

            session_confirmed = True

            now = time.monotonic()

            minimum_wait_until = (
                now + LOGIN_WAIT
            )

            deadline = (
                now + LOGIN_TIMEOUT
            )

            continue

        if time.monotonic() < minimum_wait_until:

            time.sleep(POLL_INTERVAL)
            continue

        try:

            if dashboard_is_ready(driver):

                return (
                    "SUCCESS",
                    confirmation_count,
                )

            if switch_to_dashboard_window(driver):

                return (
                    "SUCCESS",
                    confirmation_count,
                )

            if session_confirmed:

                if login_page_is_ready(driver):

                    current_url = driver.current_url.lower()

                    ready_state = driver.execute_script(
                        "return document.readyState;"
                    )

                    if (
                        "ibsg_s_login.aspx" in current_url
                        and ready_state == "complete"
                    ):

                        return (
                            "RETRY",
                            confirmation_count,
                        )

        except UnexpectedAlertPresentException:

            continue

        except StaleElementReferenceException:

            continue

        time.sleep(POLL_INTERVAL)

    raise TimeoutError(
        "Login failed or OIC dashboard was not loaded "
        f"within {LOGIN_TIMEOUT} seconds."
    )


# ============================================================
# LOGIN CONTROLLER
# ============================================================

def login(
    driver: webdriver.Chrome,
    username: str,
    password: str,
) -> None:

    log(
        "INFO",
        "Opening OIC login page",
    )

    driver.get(LOGIN_URL)

    WebDriverWait(
        driver,
        ELEMENT_TIMEOUT,
    ).until(
        EC.visibility_of_element_located(
            (By.ID, USERNAME_ID)
        )
    )

    accept_cookie_consent(driver)

    confirmation_count = 0

    for attempt in range(
        1,
        MAX_LOGIN_ATTEMPTS + 1,
    ):

        submit_login(
            driver,
            username,
            password,
            attempt,
        )

        result, confirmation_count = wait_for_login_result(
            driver,
            confirmation_count,
        )

        if result == "SUCCESS":

            log("OK", "Login successful")
            log("OK", "OIC dashboard loaded")

            return

        if result == "RETRY":

            if attempt >= MAX_LOGIN_ATTEMPTS:

                break

            log(
                "INFO",
                "OIC returned to the login page",
            )

            log(
                "RETRY",
                "Submitting login credentials again",
            )

            continue

    raise RuntimeError(
        "Maximum login attempts reached."
    )


# ============================================================
# PIVOT REPORT MENU STATUS
# ============================================================

def pivot_menu_is_expanded(
    driver: webdriver.Chrome,
) -> bool:

    containers = driver.find_elements(
        By.ID,
        PIVOT_CONTAINER_ID,
    )

    for container in containers:

        try:

            if container.is_displayed():

                return True

        except StaleElementReferenceException:

            continue

    return False


# ============================================================
# EXPAND PIVOT REPORT MENU
# ============================================================

def expand_pivot_menu(
    driver: webdriver.Chrome,
) -> None:

    wait = WebDriverWait(
        driver,
        ELEMENT_TIMEOUT,
        poll_frequency=POLL_INTERVAL,
    )

    log("INFO", "Waiting for dashboard to complete loading")

    wait.until(dashboard_is_ready)

    log("OK", "Dashboard fully loaded")

    time.sleep(2)

    if pivot_menu_is_expanded(driver):

        log("OK", "Pivot Report menu already expanded")
        return

    for attempt in range(1, 4):

        wait.until(dashboard_is_ready)

        if pivot_menu_is_expanded(driver):
            log("OK", "Pivot Report menu expanded")
            return

        toggle = wait.until(
            EC.element_to_be_clickable(
                (By.ID, PIVOT_TOGGLE_ID)
            )
        )

        log(
            "INFO",
            f"Clicking Pivot Report (+) | Attempt {attempt}",
        )

        try:
            toggle.click()

            WebDriverWait(
                driver,
                5,
                poll_frequency=POLL_INTERVAL,
            ).until(pivot_menu_is_expanded)

            log("OK", "Pivot Report menu expanded")
            return

        except (
            TimeoutException,
            StaleElementReferenceException,
        ):

            log("WARN", "Menu not expanded yet")

            time.sleep(1)

    raise TimeoutError(
        "Unable to expand Pivot Report(Cube)"
    )


# ============================================================
# FIND INSURANCE CATEGORY
# ============================================================

def find_visible_cube(
    driver: webdriver.Chrome,
    cube_code: str = TARGET_CUBE,
):

    cube = CUBE_OPTIONS.get(cube_code)

    if not cube:

        raise ValueError(
            f"Unsupported insurance cube: {cube_code}"
        )

    containers = driver.find_elements(
        By.ID,
        PIVOT_CONTAINER_ID,
    )

    for container in containers:

        try:

            if not container.is_displayed():

                continue

            # Search only inside Pivot Report(Cube).
            # This avoids accidentally selecting a similarly
            # named category under Dashboard Pivot(Cube).

            selectors = (
                (
                    By.XPATH,
                    (
                        ".//a["
                        f"acronym[@title='{cube['name']}']"
                        "]"
                    ),
                ),
                (
                    By.XPATH,
                    (
                        ".//a[normalize-space(.)="
                        f"'{cube['name']}']"
                    ),
                ),
                (
                    By.ID,
                    cube["id"],
                ),
            )

            for by, selector in selectors:

                elements = container.find_elements(
                    by,
                    selector,
                )

                for element in elements:

                    if not element.is_displayed():

                        continue

                    if not element.is_enabled():

                        continue

                    element_text = element.text.strip()

                    if element_text != cube["name"]:

                        continue

                    return (
                        cube_code,
                        cube["name"],
                        element,
                    )

        except StaleElementReferenceException:

            continue

    return False


# ============================================================
# OPEN INSURANCE CUBE
# ============================================================

def open_cube(
    driver: webdriver.Chrome,
) -> str:

    log(
        "INFO",
        "Opening Pivot Report(Cube)",
    )

    # STEP 01: Expand Pivot Report using the + button.

    expand_pivot_menu(driver)

    # STEP 02: Wait until MR becomes visible and clickable.

    log(
        "INFO",
        f"Waiting for insurance category: {TARGET_CUBE}",
    )

    selected = WebDriverWait(
        driver,
        ELEMENT_TIMEOUT,
        poll_frequency=POLL_INTERVAL,
    ).until(
        lambda current_driver: find_visible_cube(
            current_driver,
            TARGET_CUBE,
        )
    )

    cube_code, cube_name, element = selected

    actual_element_id = element.get_attribute("id")

    log(
        "OK",
        f"Insurance category located: {cube_code}",
    )

    log(
        "INFO",
        f"Category: {cube_name}",
    )

    log(
        "INFO",
        f"Element ID: {actual_element_id}",
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: "center",
            inline: "nearest"
        });
        """,
        element,
    )

    # STEP 03: Click the insurance category itself.
    # Let ASP.NET execute its own __doPostBack normally.

    log(
        "INFO",
        f"Clicking insurance category: {cube_code}",
    )

    element.click()

    log(
        "OK",
        f"Insurance category clicked: {cube_code}",
    )

    log(
        "OK",
        f"Cube selected: {cube_code}",
    )

    return cube_code


# ============================================================
# JWT VALIDATION
# ============================================================

def decode_jwt_part(
    part: str,
) -> dict:

    padding = "=" * (-len(part) % 4)

    decoded = base64.urlsafe_b64decode(
        part + padding
    )

    result = json.loads(
        decoded.decode("utf-8")
    )

    if not isinstance(result, dict):

        raise ValueError(
            "Invalid JWT section."
        )

    return result


def validate_token(
    token: str,
) -> dict:

    parts = token.split(".")

    if len(parts) != 3:

        raise ValueError(
            "Invalid JWT format."
        )

    header = decode_jwt_part(
        parts[0]
    )

    payload = decode_jwt_part(
        parts[1]
    )

    algorithm = header.get("alg")

    if not algorithm:

        raise ValueError(
            "JWT algorithm is missing."
        )

    if str(algorithm).lower() == "none":

        raise ValueError(
            "Unsigned JWT is not accepted."
        )

    expiration = payload.get("exp")

    if expiration is None:

        raise ValueError(
            "JWT expiration is missing."
        )

    if float(expiration) <= time.time():

        raise ValueError(
            "JWT has expired."
        )

    not_before = payload.get("nbf")

    if not_before is not None:

        if float(not_before) > time.time() + 5:

            raise ValueError(
                "JWT is not valid yet."
            )

    # Local format and timestamp validation only.
    # Cryptographic signature verification is not performed.

    return payload


def extract_jwt(
    data,
    depth: int = 0,
) -> str | None:

    if depth > 15:

        return None

    if isinstance(data, dict):

        preferred_keys = (
            "token",
            "Token",
            "access_token",
            "accessToken",
            "bearerToken",
            "jwt",
            "JWT",
            "d",
            "data",
            "result",
        )

        visited = set()

        for key in preferred_keys:

            if key not in data:

                continue

            visited.add(key)

            token = extract_jwt(
                data[key],
                depth + 1,
            )

            if token:

                return token

        for key, value in data.items():

            if key in visited:

                continue

            token = extract_jwt(
                value,
                depth + 1,
            )

            if token:

                return token

    elif isinstance(data, (list, tuple)):

        for item in data:

            token = extract_jwt(
                item,
                depth + 1,
            )

            if token:

                return token

    elif isinstance(data, str):

        text = data.strip()

        try:

            parsed = json.loads(text)

            if parsed != data:

                token = extract_jwt(
                    parsed,
                    depth + 1,
                )

                if token:

                    return token

        except (ValueError, TypeError):

            pass

        for match in JWT_PATTERN.finditer(text):

            candidate = match.group(0)

            try:

                validate_token(candidate)

                return candidate

            except (
                ValueError,
                TypeError,
                OverflowError,
            ):

                continue

    return None


# ============================================================
# REQUEST HEADER UTILITIES
# ============================================================

def get_header(
    headers: dict,
    header_name: str,
) -> str:

    for key, value in headers.items():

        if key.lower() == header_name.lower():

            return str(value).strip()

    return ""


def extract_request_cookie(
    request: dict,
) -> str:

    extra_headers = request.get(
        "extra_headers",
        {},
    )

    normal_headers = request.get(
        "headers",
        {},
    )

    cookie = (
        get_header(extra_headers, "Cookie")
        or get_header(normal_headers, "Cookie")
    )

    if cookie:

        return cookie

    associated_cookies = request.get(
        "associated_cookies",
        [],
    )

    cookie_parts = []

    for item in associated_cookies:

        if item.get("blockedReasons"):

            continue

        browser_cookie = item.get(
            "cookie",
            {},
        )

        name = browser_cookie.get("name")
        value = browser_cookie.get("value")

        if name and value is not None:

            cookie_parts.append(
                f"{name}={value}"
            )

    return "; ".join(cookie_parts)


def is_gentoken_request(
    url: str,
    method: str,
) -> bool:

    parsed = urlparse(url)

    return (
        parsed.scheme.lower() == "https"
        and parsed.hostname == "onlinewebibs.oic.or.th"
        and parsed.path.lower() == GENTOKEN_PATH.lower()
        and method.upper() == "POST"
    )


# ============================================================
# BROWSER WINDOW MANAGEMENT
# ============================================================

def switch_to_request_window(
    driver: webdriver.Chrome,
    webview: str,
) -> None:

    if not webview:

        return

    target = webview.removeprefix(
        "CDwindow-"
    )

    for handle in driver.window_handles:

        handle_id = handle.removeprefix(
            "CDwindow-"
        )

        if handle_id == target:

            if driver.current_window_handle != handle:

                driver.switch_to.window(
                    handle
                )

            return


# ============================================================
# RESPONSE BODY
# ============================================================

def get_response_body(
    driver: webdriver.Chrome,
    request_id: str,
    webview: str,
) -> str | None:

    try:

        switch_to_request_window(
            driver,
            webview,
        )

        result = driver.execute_cdp_cmd(
            "Network.getResponseBody",
            {
                "requestId": request_id,
            },
        )

        body = result.get(
            "body",
            "",
        )

        if result.get("base64Encoded"):

            body = base64.b64decode(
                body
            ).decode(
                "utf-8",
                errors="replace",
            )

        return body

    except (
        WebDriverException,
        ValueError,
    ):

        return None


# ============================================================
# CAPTURE GENTOKEN
# ============================================================

def capture_gentoken(
    driver: webdriver.Chrome,
) -> tuple[str, str]:

    requests = {}
    monitored_requests = set()

    deadline = (
        time.monotonic()
        + NETWORK_TIMEOUT
    )

    log(
        "INFO",
        "Monitoring GenToken network events",
    )

    while time.monotonic() < deadline:

        entries = driver.get_log(
            "performance"
        )

        for entry in entries:

            try:

                envelope = json.loads(
                    entry["message"]
                )

                message = envelope["message"]

                event = message.get(
                    "method",
                    "",
                )

                params = message.get(
                    "params",
                    {},
                )

                request_id = params.get(
                    "requestId"
                )

                if not request_id:

                    continue

                webview = envelope.get(
                    "webview",
                    "",
                )

                request_key = (
                    webview,
                    request_id,
                )

                request = requests.setdefault(
                    request_key,
                    {
                        "request_id": request_id,
                        "webview": webview,
                    },
                )

                # --------------------------------------------
                # REQUEST
                # --------------------------------------------

                if event == "Network.requestWillBeSent":

                    request_data = params.get(
                        "request",
                        {},
                    )

                    url = request_data.get(
                        "url",
                        "",
                    )

                    method = request_data.get(
                        "method",
                        "",
                    )

                    request["url"] = url
                    request["method"] = method

                    request["headers"] = request_data.get(
                        "headers",
                        {},
                    )

                    if is_gentoken_request(
                        url,
                        method,
                    ):

                        if request_key not in monitored_requests:

                            monitored_requests.add(
                                request_key
                            )

                            log(
                                "OK",
                                "GenToken request detected",
                            )

                    elif request_key in monitored_requests:

                        monitored_requests.discard(
                            request_key
                        )

                # --------------------------------------------
                # REQUEST EXTRA INFO
                # --------------------------------------------

                elif event == (
                    "Network.requestWillBeSentExtraInfo"
                ):

                    request["extra_headers"] = params.get(
                        "headers",
                        {},
                    )

                    request["associated_cookies"] = params.get(
                        "associatedCookies",
                        [],
                    )

                # --------------------------------------------
                # RESPONSE
                # --------------------------------------------

                elif event == "Network.responseReceived":

                    response = params.get(
                        "response",
                        {},
                    )

                    request["status"] = response.get(
                        "status"
                    )

                    request["response_url"] = response.get(
                        "url",
                        "",
                    )

                    if request_key in monitored_requests:

                        log(
                            "HTTP",
                            f"GenToken HTTP {request['status']}",
                        )

                # --------------------------------------------
                # RESPONSE COMPLETED
                # --------------------------------------------

                elif event == "Network.loadingFinished":

                    request["finished"] = True

                # --------------------------------------------
                # REQUEST FAILED
                # --------------------------------------------

                elif event == "Network.loadingFailed":

                    request["failed"] = True

            except (
                KeyError,
                ValueError,
                TypeError,
            ):

                continue

        # ----------------------------------------------------
        # PROCESS COMPLETED GENTOKEN REQUESTS
        # ----------------------------------------------------

        for request_key in tuple(monitored_requests):

            request = requests[request_key]

            if request.get("failed"):

                monitored_requests.discard(
                    request_key
                )

                log(
                    "WARN",
                    "GenToken request failed",
                )

                continue

            if not request.get("finished"):

                continue

            status = request.get(
                "status"
            )

            if status is None:

                continue

            if status != 200:

                monitored_requests.discard(
                    request_key
                )

                log(
                    "WARN",
                    "GenToken returned non-200 response",
                )

                continue

            response_url = request.get(
                "response_url",
                "",
            )

            if not is_gentoken_request(
                response_url,
                "POST",
            ):

                monitored_requests.discard(
                    request_key
                )

                continue

            # ------------------------------------------------
            # GET RESPONSE BODY
            # ------------------------------------------------

            if "response_body" not in request:

                response_body = get_response_body(
                    driver,
                    request["request_id"],
                    request["webview"],
                )

                if response_body is None:

                    continue

                request["response_body"] = response_body

                log(
                    "OK",
                    "GenToken response captured",
                )

            # ------------------------------------------------
            # EXTRACT JWT
            # ------------------------------------------------

            if "token" not in request:

                token = extract_jwt(
                    request["response_body"]
                )

                if not token:

                    monitored_requests.discard(
                        request_key
                    )

                    log(
                        "WARN",
                        "Response does not contain a valid JWT",
                    )

                    continue

                request["token"] = token

                log(
                    "OK",
                    "Bearer token captured",
                )

            # ------------------------------------------------
            # EXTRACT COOKIE
            # ------------------------------------------------

            cookie = extract_request_cookie(
                request
            )

            if not cookie:

                continue

            if "\r" in cookie or "\n" in cookie:

                monitored_requests.discard(
                    request_key
                )

                log(
                    "WARN",
                    "Invalid Cookie header format",
                )

                continue

            token = request["token"]

            payload = validate_token(
                token
            )

            log(
                "OK",
                "Request cookie captured",
            )

            log(
                "OK",
                f"Cookie length: {len(cookie)}",
            )

            log(
                "OK",
                f"Token length: {len(token)}",
            )

            log(
                "OK",
                "JWT structure and expiry validated",
            )

            # ------------------------------------------------
            # TOKEN EXPIRATION
            # ------------------------------------------------

            expiration = datetime.fromtimestamp(
                float(payload["exp"])
            ).astimezone()

            log(
                "INFO",
                (
                    "Token expires: "
                    f"{expiration:%Y-%m-%d %H:%M:%S %z}"
                ),
            )

            remaining = max(
                0,
                int(
                    float(payload["exp"])
                    - time.time()
                ),
            )

            hours, remainder = divmod(
                remaining,
                3600,
            )

            minutes, seconds = divmod(
                remainder,
                60,
            )

            log(
                "INFO",
                (
                    "Token remaining: "
                    f"{hours:02d}:"
                    f"{minutes:02d}:"
                    f"{seconds:02d}"
                ),
            )

            return cookie, token

        time.sleep(
            POLL_INTERVAL
        )

    raise TimeoutError(
        "Unable to capture a valid GenToken Cookie and JWT "
        f"within {NETWORK_TIMEOUT} seconds."
    )


# ============================================================
# SUMMARY
# ============================================================

def print_summary(
    cube_code: str,
    started_at: float,
) -> None:

    elapsed = (
        time.monotonic()
        - started_at
    )

    print()
    print("=" * 76)
    print("                     AUTHENTICATION COMPLETED")
    print("=" * 76)
    print("  Login            : SUCCESS")
    print("  Session Handling : COMPLETED")
    print("  Pivot Report     : EXPANDED")
    print(f"  Insurance Cube   : {cube_code}")
    print("  Category Click   : SUCCESS")
    print("  GenToken         : SUCCESS")
    print("  HTTP Status      : 200 OK")
    print("  Cookie           : CAPTURED")
    print("  Bearer Token     : CAPTURED")
    print("  Environment      : UPDATED")
    print(f"  Output           : {ENV_PATH}")
    print(f"  Elapsed          : {elapsed:.2f} seconds")
    print("=" * 76)
    print()


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    print_banner()

    driver = None

    started_at = time.monotonic()

    try:

        # STEP 01: LOAD ENVIRONMENT

        username, password = load_credentials()

        # STEP 02: LAUNCH CHROME

        driver = create_driver()

        # STEP 03: LOGIN AND SESSION CONFLICT HANDLING

        login(
            driver,
            username,
            password,
        )

        # STEP 04: CLEAR PREVIOUS NETWORK EVENTS

        driver.get_log(
            "performance"
        )

        log(
            "OK",
            "Previous network events cleared",
        )

        # STEP 05:
        # Click Pivot Report (+)
        # Wait for category container
        # Click insurance category MR

        cube_code = open_cube(
            driver
        )

        # STEP 06: CAPTURE GENTOKEN

        cookie, token = capture_gentoken(
            driver
        )

        # STEP 07: UPDATE ENVIRONMENT

        update_env(
            cookie,
            token,
        )

        # STEP 08: SUMMARY

        print_summary(
            cube_code,
            started_at,
        )

        return 0

    except KeyboardInterrupt:

        log(
            "STOP",
            "Authentication cancelled",
        )

        return 130

    except (
        FileNotFoundError,
        TimeoutError,
        ValueError,
        RuntimeError,
    ) as error:

        log(
            "ERROR",
            str(error),
        )

        log(
            "INFO",
            "Incomplete authentication values were not saved",
        )

        return 1

    except TimeoutException:

        log(
            "ERROR",
            "Selenium element or page timeout",
        )

        log(
            "INFO",
            "Check Pivot Report menu and insurance category selectors",
        )

        log(
            "INFO",
            "Existing .env authentication values preserved",
        )

        return 1

    except UnexpectedAlertPresentException:

        log(
            "ERROR",
            "Unhandled JavaScript dialog detected",
        )

        return 1

    except WebDriverException:

        log(
            "ERROR",
            "Chrome or Selenium operation failed",
        )

        return 1

    except Exception as error:

        log(
            "ERROR",
            f"Unexpected error: {type(error).__name__}",
        )

        return 1

    finally:

        if driver is not None and CLOSE_BROWSER:

            try:

                alert = get_current_alert(
                    driver
                )

                if alert is not None:

                    alert.dismiss()

            except WebDriverException:

                pass

            try:

                driver.quit()

                log(
                    "OK",
                    "Chrome closed",
                )

            except WebDriverException:

                pass


# ============================================================
# ENTRY POINT
# ============================================================

if __name__ == "__main__":

    raise SystemExit(main())