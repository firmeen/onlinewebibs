
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC XMLA COLLECTOR
Single Python File | Auto Bearer Token | Excel Only

Requirements:
    python -m pip install requests

Commands:
    python oic_collector.py --test-only
    python oic_collector.py
    python oic_collector.py --codes AH FI
    python oic_collector.py --timeout 300

Environment:
    .env in the same directory

    OIC_BEARER_TOKEN=...
    OIC_COOKIE=...

Authentication:
    1. Load the existing OIC_COOKIE from .env.
    2. Request a fresh JWT from PivotCube.asmx/GenToken.
    3. Use the fresh JWT to call the XMLA endpoint.
    4. Refresh the JWT when necessary.
    5. Export the collected data to Excel.

Output:
    ./output/oic_insurance_YYYYMMDD_HHMMSS.xlsx

No SQL Server.
No pytds.
No pyodbc.
No additional Python files required.
"""

import argparse
import base64
import hashlib
import json
import os
import re
import sys
import time
import zipfile

from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

import requests


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

ENV_PATH = BASE_DIR / ".env"
OUTPUT_DIR = BASE_DIR / "output"

BASE_URL = "https://onlinewebibs.oic.or.th"

TOKEN_URL = (
    BASE_URL
    + "/WebForm/Dashboard/Pivot/PivotCube.asmx/GenToken"
)

XMLA_URL = (
    BASE_URL
    + "/xmla-proxy/api/xmla-proxy"
)

PAGE_PATH = (
    "/WebForm/Dashboard/Pivot/"
    "DynamicPivotCube.aspx"
)

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"
MD_NS = "urn:schemas-microsoft-com:xml-analysis:mddataset"

DEFAULT_TIMEOUT = 180
DEFAULT_RETRIES = 3

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)


# ============================================================
# INSURANCE PROFILES
# ============================================================

PROFILES = {
    "AH": {
        "name": "อุบัติเหตุและสุขภาพ",
        "class": "AH",
        "catalog": "PIVOT_202609_AH",
        "cube": "AH_AGGREGATE_FY",
        "column": "[VW PV AH AGGREGATE FY].[FY YEAR]",
        "row1": "[DIM MAINSUB CLASS AH].[MAIN NAME TH]",
        "row2": "[DIM MAINSUB CLASS AH].[SUB NAME TH]",
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "EARNED",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "INMTH CLM PAID",
            "ATMTH CLM LI",
        ],
    },

    "FI": {
        "name": "ทรัพย์สิน",
        "class": "FI",
        "catalog": "PIVOT_202609_FI",
        "cube": "FI_AGGREGATE_FY",
        "column": "[VW PV FI AGGREGATE FY].[YEAR]",
        "row1": "[DIM CT SUBCLASS FI].[GROUP DESC TH]",
        "row2": None,
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "EARNED PREMIUM",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "SUM CLM PAID",
            "SUM CLM OS",
            "SUM CLM LI",
        ],
    },

    "MR": {
        "name": "ทางทะเลและขนส่ง",
        "class": "MR",
        "catalog": "PIVOT_202609_MR",
        "cube": "MR_AGGREGATE_FY",
        "column": "[VW PV MR AGGREGATE FY].[STATUS COMPANY PASS]",
        "row1": "[DIM CT SUBCLASS MR].[SUBCLASS DESC TH]",
        "row2": None,
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "CLAIM PAID",
            "CLAIM OS",
            "CLAIM LI",
        ],
    },

    "MI": {
        "name": "เบ็ดเตล็ด",
        "class": "MI",
        "catalog": "PIVOT_202609_MI",
        "cube": "MI_AGGREGATE_FY",
        "column": "[VW PV MI AGGREGATE FY].[FY YEAR]",
        "row1": "[DIM MAINSUB CLASS].[MAIN NAME TH]",
        "row2": "[DIM MAINSUB CLASS].[SUB NAME TH]",
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "CLM PAID",
            "CLM OS",
            "CLM LI",
        ],
    },

    "MV": {
        "name": "รถยนต์",
        "class": "MO",
        "catalog": "PIVOT_202609_MV",
        "cube": "MV_AGGREGATE_FY",
        "column": "[VW PV MV AGGREGATE FY].[YEAR]",
        "row1": "[DIM CT SUBCLASS MV].[SUBCLASS DESC TH]",
        "row2": None,
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "EARNED PREMIUM",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "CLAIM PAID",
            "CLAIM OS",
            "CLAIM LI",
        ],
    },
}


# ============================================================
# EXCEL COLUMNS
# ============================================================

EXCEL_COLUMNS = [
    "record_key",
    "insurance_code",
    "insurance_type",
    "catalog",
    "cube_name",
    "column_dimension",
    "column_member",
    "column_unique_name",
    "row_dimension_1",
    "row_member_1",
    "row_unique_name_1",
    "row_dimension_2",
    "row_member_2",
    "row_unique_name_2",
    "aggregation_level",
    "measure",
    "metric_value",
    "metric_raw",
    "last_data_update",
    "retrieved_at",
]


# ============================================================
# LOGGING
# ============================================================

def log(level, message):
    timestamp = datetime.now().strftime("%H:%M:%S")
    print(f"[{timestamp}] [{level}] {message}", flush=True)


def banner():
    print("=" * 72)
    print("OIC XMLA COLLECTOR")
    print("AUTO TOKEN | LIVE API | EXCEL ONLY")
    print("=" * 72)


# ============================================================
# ENVIRONMENT
# ============================================================

def load_env():
    if not ENV_PATH.exists():
        raise RuntimeError(
            f"ไม่พบไฟล์ .env ที่ {ENV_PATH}"
        )

    loaded = {}

    for raw_line in ENV_PATH.read_text(
        encoding="utf-8-sig"
    ).splitlines():

        line = raw_line.strip()

        if not line:
            continue

        if line.startswith("#"):
            continue

        if "=" not in line:
            continue

        key, value = line.split("=", 1)

        key = key.strip()
        value = value.strip()

        if not re.fullmatch(
            r"[A-Za-z_][A-Za-z0-9_]*",
            key,
        ):
            continue

        if (
            len(value) >= 2
            and value[0] == value[-1]
            and value[0] in ('"', "'")
        ):
            value = value[1:-1]

        loaded[key] = value

    for key, value in loaded.items():
        os.environ.setdefault(key, value)

    cookie = os.getenv("OIC_COOKIE", "").strip()
    token = os.getenv("OIC_BEARER_TOKEN", "").strip()

    if token.lower().startswith("bearer "):
        token = token[7:].strip()

    if token.startswith("REPLACE_WITH"):
        token = ""

    return {
        "cookie": cookie,
        "token": token,
    }


# ============================================================
# JWT UTILITIES
# ============================================================

def decode_jwt_payload(token):
    try:
        parts = token.split(".")

        if len(parts) != 3:
            return {}

        payload = parts[1]
        payload += "=" * (-len(payload) % 4)

        decoded = base64.urlsafe_b64decode(
            payload.encode("ascii")
        )

        return json.loads(decoded.decode("utf-8"))

    except Exception:
        return {}


def jwt_expiration(token):
    payload = decode_jwt_payload(token)

    try:
        return int(payload.get("exp", 0))
    except (TypeError, ValueError):
        return 0


def jwt_is_valid_time(token, margin_seconds=60):
    if not token:
        return False

    expiry = jwt_expiration(token)

    if not expiry:
        return False

    return time.time() + margin_seconds < expiry


def format_expiration(token):
    expiry = jwt_expiration(token)

    if not expiry:
        return "UNKNOWN"

    return datetime.fromtimestamp(
        expiry
    ).strftime("%Y-%m-%d %H:%M:%S")


# ============================================================
# SOAP REQUEST BUILDING
# ============================================================

def build_mdx(profile):
    measures = ",".join(
        f"[Measures].[{name}]"
        for name in profile["measures"]
    )

    column = profile["column"]
    row1 = profile["row1"]
    row2 = profile["row2"]

    dx_columns = (
        f"NonEmpty({{{column}.allMembers}}, "
        f"{{{measures}}})"
    )

    if row2:
        subtotal = (
            "NonEmpty("
            f"CrossJoin({{{row1}.allMembers}},"
            f"{{({row2}.[All])}}), "
            f"{{{measures}}})"
        )

        details = (
            "NonEmpty("
            f"CrossJoin({{{row1}.allMembers}},"
            f"{{{row2}.allMembers}}), "
            f"{{{measures}}})"
        )

        dx_rows = f"Union({subtotal},{details})"

    else:
        dx_rows = (
            f"NonEmpty({{{row1}.allMembers}}, "
            f"{{{measures}}})"
        )

    return (
        f"with set [DX_columns] as {dx_columns} "
        f"set [DX_rows] as {dx_rows} "
        f"SELECT CrossJoin([DX_columns],{{{measures}}}) "
        "DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,"
        "HIERARCHY_UNIQUE_NAME,"
        "MEMBER_VALUE ON columns,"
        "[DX_rows] DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,"
        "HIERARCHY_UNIQUE_NAME,"
        "MEMBER_VALUE ON rows "
        f"FROM [{profile['cube']}] "
        "CELL PROPERTIES "
        "VALUE, FORMAT_STRING, LANGUAGE, "
        "BACK_COLOR, FORE_COLOR, FONT_FLAGS"
    )


def build_soap(profile):
    envelope = ET.Element(
        f"{{{SOAP_NS}}}Envelope"
    )

    body = ET.SubElement(
        envelope,
        f"{{{SOAP_NS}}}Body",
    )

    execute = ET.SubElement(
        body,
        f"{{{XMLA_NS}}}Execute",
    )

    command = ET.SubElement(
        execute,
        f"{{{XMLA_NS}}}Command",
    )

    ET.SubElement(
        command,
        f"{{{XMLA_NS}}}Statement",
    ).text = build_mdx(profile)

    properties = ET.SubElement(
        execute,
        f"{{{XMLA_NS}}}Properties",
    )

    property_list = ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}PropertyList",
    )

    values = {
        "Catalog": profile["catalog"],
        "ShowHiddenCubes": "True",
        "SspropInitAppName": (
            "Microsoft SQL Server Management Studio"
        ),
        "Timeout": "3600",
        "LocaleIdentifier": "9",
    }

    for key, value in values.items():
        ET.SubElement(
            property_list,
            f"{{{XMLA_NS}}}{key}",
        ).text = value

    return ET.tostring(
        envelope,
        encoding="utf-8",
        xml_declaration=False,
    )


# ============================================================
# XML UTILITIES
# ============================================================

def local_name(tag):
    return tag.rsplit("}", 1)[-1]


def child_text(element, name):
    for child in element:
        if local_name(child.tag) == name:
            return child.text or ""

    return ""


def parse_xmla_response(content, code):
    try:
        root = ET.fromstring(content)

    except ET.ParseError as exc:
        raise RuntimeError(
            f"{code}: Response ไม่ใช่ XML ที่ถูกต้อง: {exc}"
        ) from exc

    for element in root.iter():
        if local_name(element.tag) == "Fault":
            message = " ".join(
                (node.text or "").strip()
                for node in element.iter()
                if node.text
            )

            raise RuntimeError(
                f"{code}: SOAP Fault: {message[:500]}"
            )

    mdroot = next(
        (
            element
            for element in root.iter()
            if element.tag == f"{{{MD_NS}}}root"
        ),
        None,
    )

    if mdroot is None:
        raise RuntimeError(
            f"{code}: ไม่พบ XMLA multidimensional dataset"
        )

    return mdroot


# ============================================================
# AUTHENTICATION
# ============================================================

class OICClient:

    def __init__(self, cookie="", token="", timeout=180, retries=3):
        self.session = requests.Session()

        self.timeout = timeout
        self.retries = max(1, retries)

        self.cookie = cookie
        self.token = token

        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Accept-Language": "th-TH,th;q=0.9",
        })

    def close(self):
        self.session.close()

    def page_url(self, code):
        profile = PROFILES[code]

        return (
            f"{PAGE_PATH}"
            f"?Class={profile['class']}"
        )

    def request_headers(self, code):
        return {
            "Accept": (
                "application/json, text/javascript, "
                "*/*; q=0.01"
            ),
            "Content-Type": (
                "application/json; charset=UTF-8"
            ),
            "Origin": BASE_URL,
            "Referer": (
                BASE_URL + self.page_url(code)
            ),
            "X-Requested-With": "XMLHttpRequest",
        }

    def set_cookie(self, headers):
        if self.cookie:
            headers["Cookie"] = self.cookie

        return headers

    def request_token(self, code):
        """
        Request a JWT using the site's observed GenToken endpoint.

        The existing authorized session cookie is required when
        the server associates token generation with a session.

        This method does not bypass login, CAPTCHA, WAF,
        or any other server-side access control.
        """

        page = self.page_url(code)

        payload = {
            "_tk": "Init",
            "_pageUrl": page,
        }

        headers = self.set_cookie(
            self.request_headers(code)
        )

        log(
            "AUTH",
            f"Requesting fresh token for {code}",
        )

        try:
            response = self.session.post(
                TOKEN_URL,
                json=payload,
                headers=headers,
                timeout=(15, 45),
                allow_redirects=False,
            )

        except requests.RequestException as exc:
            raise RuntimeError(
                "GenToken request failed: "
                f"{type(exc).__name__}: {exc}"
            ) from exc

        if response.status_code in (401, 403):
            raise RuntimeError(
                f"GenToken HTTP {response.status_code}: "
                "Session หรือ Cookie ไม่ได้รับอนุญาต "
                "หรือหมดอายุแล้ว ต้องเข้าสู่เว็บไซต์ "
                "ตามขั้นตอนปกติและใช้ Session ที่ถูกต้อง"
            )

        if response.status_code != 200:
            raise RuntimeError(
                "GenToken HTTP "
                f"{response.status_code}: "
                "ไม่สามารถสร้าง Token ได้"
            )

        try:
            data = response.json()

        except ValueError as exc:
            raise RuntimeError(
                "GenToken ไม่ได้ตอบกลับเป็น JSON "
                "อาจเป็นหน้า Login หรือ Session หมดอายุ"
            ) from exc

        token = data.get("d")

        if not isinstance(token, str):
            raise RuntimeError(
                "GenToken response ไม่มี JWT ใน field 'd'"
            )

        token = token.strip()

        if token.lower().startswith("bearer "):
            token = token[7:].strip()

        if len(token.split(".")) != 3:
            raise RuntimeError(
                "GenToken ไม่ได้คืน JWT\n"
                f"HTTP Status: {response.status_code}\n"
                f"Response Type: {type(data.get('d')).__name__}\n"
                f"Response Preview: {token[:200] if len(token) < 200 else '[REDACTED]'}"
            )

        expiry = jwt_expiration(token)

        if expiry and expiry <= time.time():
            raise RuntimeError(
                "GenToken ส่ง Token ที่หมดอายุแล้ว"
            )

        self.token = token

        log("AUTH", "GenToken successful")

        log(
            "AUTH",
            "JWT expires: "
            + format_expiration(token),
        )

        return token

    def ensure_token(self, code, force_refresh=False):
        if (
            not force_refresh
            and jwt_is_valid_time(
                self.token,
                margin_seconds=120,
            )
        ):
            return self.token

        return self.request_token(code)

    def xmla_headers(self, code):
        headers = {
            "Accept": "text/plain, */*; q=0.01",
            "Content-Type": "text/xml",
            "Origin": BASE_URL,
            "Referer": (
                BASE_URL + self.page_url(code)
            ),
            "X-Requested-With": "XMLHttpRequest",
            "Authorization": f"Bearer {self.token}",
        }

        return self.set_cookie(headers)

    def fetch(self, code):
        profile = PROFILES[code]

        self.ensure_token(code)

        payload = build_soap(profile)

        for attempt in range(1, self.retries + 1):

            headers = self.xmla_headers(code)

            try:
                response = self.session.post(
                    XMLA_URL,
                    data=payload,
                    headers=headers,
                    timeout=(15, self.timeout),
                    allow_redirects=False,
                )

            except requests.RequestException as exc:
                if attempt >= self.retries:
                    raise RuntimeError(
                        f"{code}: Network error: {exc}"
                    ) from exc

                delay = min(2 ** attempt, 15)

                log(
                    "RETRY",
                    f"{code}: Connection error; "
                    f"retrying in {delay}s",
                )

                time.sleep(delay)
                continue

            status = response.status_code

            if status == 200:
                mdroot = parse_xmla_response(
                    response.content,
                    code,
                )

                return mdroot, len(response.content)

            if status in (401, 403):
                if attempt == 1:
                    log(
                        "AUTH",
                        f"{code}: HTTP {status}; "
                        "requesting a fresh JWT",
                    )

                    self.ensure_token(
                        code,
                        force_refresh=True,
                    )

                    continue

                raise RuntimeError(
                    f"{code}: HTTP {status} หลังขอ "
                    "Token ใหม่แล้ว เซิร์ฟเวอร์ยังปฏิเสธ "
                    "Request โปรดตรวจสอบสิทธิ์และ Session"
                )

            if 300 <= status < 400:
                raise RuntimeError(
                    f"{code}: HTTP {status}; "
                    "Redirect: "
                    + response.headers.get(
                        "Location",
                        "unknown",
                    )
                )

            if status in (
                429,
                500,
                502,
                503,
                504,
            ):
                if attempt < self.retries:
                    retry_after = response.headers.get(
                        "Retry-After",
                        "",
                    )

                    if retry_after.isdigit():
                        delay = min(
                            int(retry_after),
                            60,
                        )
                    else:
                        delay = min(
                            2 ** attempt,
                            15,
                        )

                    log(
                        "RETRY",
                        f"{code}: HTTP {status}; "
                        f"retrying in {delay}s",
                    )

                    time.sleep(delay)
                    continue

            raise RuntimeError(
                f"{code}: HTTP {status}; "
                "XMLA request rejected"
            )

        raise RuntimeError(
            f"{code}: XMLA request failed"
        )


# ============================================================
# XMLA DATA EXTRACTION
# ============================================================

def members_from_tuple(tuple_element):
    result = []

    for member in tuple_element.findall(
        f"{{{MD_NS}}}Member"
    ):
        unique_name = child_text(
            member,
            "UName",
        )

        result.append({
            "dimension": member.attrib.get(
                "Hierarchy",
                "",
            ),
            "caption": child_text(
                member,
                "Caption",
            ),
            "unique_name": unique_name,
            "is_all": unique_name.endswith("[All]"),
        })

    return result


def empty_member():
    return {
        "dimension": "",
        "caption": "",
        "unique_name": "",
        "is_all": False,
    }


def parse_dataset(mdroot, code):
    profile = PROFILES[code]

    axes_element = mdroot.find(
        f"{{{MD_NS}}}Axes"
    )

    if axes_element is None:
        raise RuntimeError(
            f"{code}: XMLA response has no Axes"
        )

    axes = {
        axis.attrib.get("name"): axis
        for axis in axes_element.findall(
            f"{{{MD_NS}}}Axis"
        )
    }

    axis0_element = axes.get("Axis0")
    axis1_element = axes.get("Axis1")

    if axis0_element is None or axis1_element is None:
        raise RuntimeError(
            f"{code}: Missing Axis0 or Axis1"
        )

    axis0 = [
        members_from_tuple(element)
        for element in axis0_element.findall(
            f".//{{{MD_NS}}}Tuples/"
            f"{{{MD_NS}}}Tuple"
        )
    ]

    axis1 = [
        members_from_tuple(element)
        for element in axis1_element.findall(
            f".//{{{MD_NS}}}Tuples/"
            f"{{{MD_NS}}}Tuple"
        )
    ]

    if not axis0:
        raise RuntimeError(
            f"{code}: Axis0 has no tuples"
        )

    cube_name = profile["cube"]
    last_data_update = ""

    for element in mdroot.iter():
        name = local_name(
            element.tag
        )

        if name == "CubeName" and element.text:
            cube_name = element.text

        elif (
            name == "LastDataUpdate"
            and element.text
        ):
            last_data_update = element.text

    cells_element = mdroot.find(
        f"{{{MD_NS}}}CellData"
    )

    if cells_element is None:
        raise RuntimeError(
            f"{code}: XMLA response has no CellData"
        )

    results = []
    column_count = len(axis0)

    retrieved_at = datetime.now(
        timezone.utc
    ).strftime("%Y-%m-%dT%H:%M:%SZ")

    for cell in cells_element.findall(
        f"{{{MD_NS}}}Cell"
    ):
        raw_value = cell.findtext(
            f"{{{MD_NS}}}Value"
        )

        if (
            raw_value is None
            or raw_value.strip() == ""
        ):
            continue

        ordinal = int(
            cell.attrib["CellOrdinal"]
        )

        row_index, column_index = divmod(
            ordinal,
            column_count,
        )

        if row_index >= len(axis1):
            raise RuntimeError(
                f"{code}: Invalid CellOrdinal "
                f"{ordinal}"
            )

        column_members = axis0[
            column_index
        ]

        row_members = axis1[
            row_index
        ]

        measure = next(
            (
                member
                for member in column_members
                if member["dimension"] == "[Measures]"
            ),
            None,
        )

        dimension_members = [
            member
            for member in column_members
            if member["dimension"] != "[Measures]"
        ]

        column_member = (
            dimension_members[0]
            if dimension_members
            else empty_member()
        )

        row_member_1 = (
            row_members[0]
            if len(row_members) > 0
            else empty_member()
        )

        row_member_2 = (
            row_members[1]
            if len(row_members) > 1
            else empty_member()
        )

        if row_members and all(
            member["is_all"]
            for member in row_members
        ):
            aggregation_level = "TOTAL"

        elif any(
            member["is_all"]
            for member in row_members
        ):
            aggregation_level = "SUBTOTAL"

        else:
            aggregation_level = "DETAIL"

        key_source = json.dumps(
            [
                profile["catalog"],
                cube_name,
                [
                    member["unique_name"]
                    for member in column_members
                ],
                [
                    member["unique_name"]
                    for member in row_members
                ],
            ],
            ensure_ascii=False,
            separators=(",", ":"),
        )

        record_key = hashlib.sha256(
            key_source.encode("utf-8")
        ).hexdigest()

        results.append({
            "record_key": record_key,

            "insurance_code": code,
            "insurance_type": profile["name"],

            "catalog": profile["catalog"],
            "cube_name": cube_name,

            "column_dimension": (
                column_member["dimension"]
            ),
            "column_member": (
                column_member["caption"]
            ),
            "column_unique_name": (
                column_member["unique_name"]
            ),

            "row_dimension_1": (
                row_member_1["dimension"]
            ),
            "row_member_1": (
                row_member_1["caption"]
            ),
            "row_unique_name_1": (
                row_member_1["unique_name"]
            ),

            "row_dimension_2": (
                row_member_2["dimension"]
            ),
            "row_member_2": (
                row_member_2["caption"]
            ),
            "row_unique_name_2": (
                row_member_2["unique_name"]
            ),

            "aggregation_level": aggregation_level,

            "measure": (
                measure["caption"]
                if measure
                else ""
            ),

            "metric_value": raw_value,
            "metric_raw": raw_value,

            "last_data_update": last_data_update,
            "retrieved_at": retrieved_at,
        })

    if not results:
        raise RuntimeError(
            f"{code}: No nonempty data cells returned"
        )

    return results


# ============================================================
# XLSX WRITER — STANDARD LIBRARY ONLY
# ============================================================

def excel_column_letter(number):
    result = ""

    while number:
        number, remainder = divmod(
            number - 1,
            26,
        )

        result = (
            chr(65 + remainder)
            + result
        )

    return result


def excel_number(value):
    """
    Convert a numeric string to an Excel number only when
    it can be represented without excessive precision loss.

    Long or high-precision values are retained as text.
    The metric_raw column always stores the original value.
    """

    try:
        number = Decimal(
            str(value)
        )

        if not number.is_finite():
            return None

        significant_digits = len(
            number.normalize().as_tuple().digits
        )

        if significant_digits > 15:
            return None

        return number

    except (
        InvalidOperation,
        ValueError,
        TypeError,
    ):
        return None


def excel_cell(address, value, header=False, numeric=False):
    style = ' s="1"' if header else ""

    if numeric:
        number = excel_number(value)

        if number is not None:
            return (
                f'<c r="{address}"{style}>'
                f"<v>{number}</v>"
                "</c>"
            )

    text = escape(
        str(
            value
            if value is not None
            else ""
        )
    )

    return (
        f'<c r="{address}"{style} t="inlineStr">'
        f'<is><t xml:space="preserve">{text}</t></is>'
        "</c>"
    )


def write_xml(archive_stream, text):
    archive_stream.write(
        text.encode("utf-8")
    )


def write_sheet(archive, sheet_index, records):
    sheet_path = (
        f"xl/worksheets/sheet{sheet_index}.xml"
    )

    with archive.open(
        sheet_path,
        "w",
        force_zip64=True,
    ) as stream:

        write_xml(
            stream,
            '<?xml version="1.0" encoding="UTF-8" '
            'standalone="yes"?>'
            '<worksheet xmlns="http://schemas.'
            'openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0" '
            'showGridLines="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            "<sheetData>",
        )

        header_cells = "".join(
            excel_cell(
                f"{excel_column_letter(index)}1",
                column,
                header=True,
            )
            for index, column in enumerate(
                EXCEL_COLUMNS,
                1,
            )
        )

        write_xml(
            stream,
            f'<row r="1">{header_cells}</row>',
        )

        for row_number, record in enumerate(
            records,
            2,
        ):
            cells = []

            for col_number, column in enumerate(
                EXCEL_COLUMNS,
                1,
            ):
                address = (
                    f"{excel_column_letter(col_number)}"
                    f"{row_number}"
                )

                cells.append(
                    excel_cell(
                        address,
                        record.get(
                            column,
                            "",
                        ),
                        numeric=(
                            column == "metric_value"
                        ),
                    )
                )

            write_xml(
                stream,
                f'<row r="{row_number}">'
                f'{"".join(cells)}'
                "</row>",
            )

        last_row = len(records) + 1

        write_xml(
            stream,
            "</sheetData>"
            f'<autoFilter ref="A1:T{last_row}"/>'
            "</worksheet>",
        )


def export_excel(data_by_code):
    if not data_by_code:
        raise RuntimeError(
            "No data available for Excel export"
        )

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    filename = (
        "oic_insurance_"
        + datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + ".xlsx"
    )

    output_path = OUTPUT_DIR / filename

    sheets = list(
        data_by_code.items()
    )

    content_types = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/'
        'package/2006/content-types">',
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-'
        'package.relationships+xml"/>',
        '<Default Extension="xml" '
        'ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-'
        'officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-'
        'officedocument.spreadsheetml.styles+xml"/>',
    ]

    for index in range(
        1,
        len(sheets) + 1,
    ):
        content_types.append(
            f'<Override PartName="/xl/worksheets/'
            f'sheet{index}.xml" '
            'ContentType="application/vnd.openxmlformats-'
            'officedocument.spreadsheetml.worksheet+xml"/>'
        )

    content_types.append(
        "</Types>"
    )

    workbook_sheets = "".join(
        f'<sheet name="{escape(code)}" '
        f'sheetId="{index}" r:id="rId{index}"/>'
        for index, (code, _) in enumerate(
            sheets,
            1,
        )
    )

    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.'
        'org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships">'
        f"<sheets>{workbook_sheets}</sheets>"
        "</workbook>"
    )

    workbook_relationships = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<Relationships xmlns="http://schemas.'
        'openxmlformats.org/package/2006/relationships">',
    ]

    for index in range(
        1,
        len(sheets) + 1,
    ):
        workbook_relationships.append(
            f'<Relationship Id="rId{index}" '
            'Type="http://schemas.openxmlformats.org/'
            'officeDocument/2006/relationships/worksheet" '
            f'Target="worksheets/sheet{index}.xml"/>'
        )

    workbook_relationships.append(
        '<Relationship Id="rIdStyles" '
        'Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/styles" '
        'Target="styles.xml"/>'
    )

    workbook_relationships.append(
        "</Relationships>"
    )

    root_relationships = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.'
        'org/package/2006/relationships">'
        '<Relationship Id="rId1" '
        'Type="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships/officeDocument" '
        'Target="xl/workbook.xml"/>'
        "</Relationships>"
    )

    styles_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<styleSheet xmlns="http://schemas.openxmlformats.'
        'org/spreadsheetml/2006/main">'
        '<fonts count="2">'
        '<font><sz val="11"/><name val="Calibri"/></font>'
        '<font><b/><sz val="11"/>'
        '<color rgb="FFFFFFFF"/>'
        '<name val="Calibri"/></font>'
        "</fonts>"
        '<fills count="2">'
        '<fill><patternFill patternType="none"/></fill>'
        '<fill><patternFill patternType="solid">'
        '<fgColor rgb="FF17324D"/>'
        "</patternFill></fill>"
        "</fills>"
        '<borders count="1"><border/></borders>'
        '<cellStyleXfs count="1">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        "</cellStyleXfs>"
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="1" '
        'borderId="0" applyFont="1" applyFill="1"/>'
        "</cellXfs>"
        "</styleSheet>"
    )

    with zipfile.ZipFile(
        output_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=6,
    ) as archive:

        archive.writestr(
            "[Content_Types].xml",
            "".join(
                content_types
            ),
        )

        archive.writestr(
            "_rels/.rels",
            root_relationships,
        )

        archive.writestr(
            "xl/workbook.xml",
            workbook_xml,
        )

        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            "".join(
                workbook_relationships
            ),
        )

        archive.writestr(
            "xl/styles.xml",
            styles_xml,
        )

        for index, (
            code,
            records,
        ) in enumerate(
            sheets,
            1,
        ):
            log(
                "EXCEL",
                f"Writing {code}: "
                f"{len(records):,} rows",
            )

            write_sheet(
                archive,
                index,
                records,
            )

    return output_path


# ============================================================
# MAIN
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description=(
            "OIC XMLA Collector — "
            "Auto JWT and Excel Only"
        )
    )

    parser.add_argument(
        "--test-only",
        action="store_true",
        help=(
            "Test GenToken and one XMLA request "
            "without exporting Excel"
        ),
    )

    parser.add_argument(
        "--codes",
        nargs="+",
        choices=list(PROFILES.keys()),
        default=list(PROFILES.keys()),
        help="Insurance categories",
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
    )

    parser.add_argument(
        "--retries",
        type=int,
        default=DEFAULT_RETRIES,
    )

    args = parser.parse_args()

    banner()

    config = load_env()

    log(
        "ENV",
        f"Loaded {ENV_PATH}",
    )

    log(
        "ENV",
        "Cookie: "
        + (
            "FOUND"
            if config["cookie"]
            else "NOT FOUND"
        ),
    )

    if config["token"]:
        log(
            "ENV",
            "Existing JWT expires: "
            + format_expiration(
                config["token"]
            ),
        )

    selected_codes = list(
        dict.fromkeys(
            args.codes
        )
    )

    log(
        "MODE",
        "TEST ONLY"
        if args.test_only
        else "LIVE EXCEL",
    )

    log(
        "SCOPE",
        ", ".join(
            selected_codes
        ),
    )

    print("-" * 72)

    client = OICClient(
        cookie=config["cookie"],
        token=config["token"],
        timeout=args.timeout,
        retries=args.retries,
    )

    data_by_code = {}
    failed_codes = {}

    try:
        first_code = selected_codes[0]

        # Always call the official GenToken endpoint first.
        # Never depend on an expired JWT saved in .env.
        client.request_token(
            first_code
        )

        log(
            "TEST",
            f"Testing XMLA request: {first_code}",
        )

        first_root, first_size = client.fetch(
            first_code
        )

        log(
            "TEST",
            "XMLA request: HTTP 200 OK",
        )

        log(
            "TEST",
            f"Response size: {first_size:,} bytes",
        )

        first_rows = parse_dataset(
            first_root,
            first_code,
        )

        log(
            "TEST",
            f"Parsed rows: {len(first_rows):,}",
        )

        log(
            "TEST",
            "SUCCESS",
        )

        if args.test_only:
            print("=" * 72)

            log(
                "DONE",
                "API test passed. "
                "No Excel file created.",
            )

            return

        data_by_code[first_code] = first_rows

        print("-" * 72)

        log(
            "DATA",
            f"{first_code} "
            f"{PROFILES[first_code]['name']}: "
            f"{len(first_rows):,} rows",
        )

        for code in selected_codes[1:]:

            log(
                "FETCH",
                f"{code} "
                f"{PROFILES[code]['name']}",
            )

            try:
                root, response_size = client.fetch(
                    code
                )

                records = parse_dataset(
                    root,
                    code,
                )

                data_by_code[code] = records

                log(
                    "OK",
                    f"{code}: "
                    f"{len(records):,} rows | "
                    f"{response_size:,} bytes",
                )

            except Exception as exc:
                failed_codes[code] = str(exc)

                log(
                    "FAILED",
                    f"{code}: {exc}",
                )

        print("-" * 72)

        log(
            "EXPORT",
            "Creating Excel workbook",
        )

        excel_path = export_excel(
            data_by_code
        )

        total_rows = sum(
            len(records)
            for records in data_by_code.values()
        )

        print("=" * 72)

        log(
            "DONE",
            "Excel created successfully",
        )

        log(
            "FILE",
            str(
                excel_path.resolve()
            ),
        )

        log(
            "SHEETS",
            ", ".join(
                data_by_code.keys()
            ),
        )

        log(
            "ROWS",
            f"{total_rows:,}",
        )

        log(
            "FILE SIZE",
            f"{excel_path.stat().st_size / 1024 / 1024:.2f} MB",
        )

        if failed_codes:
            log(
                "WARNING",
                f"{len(failed_codes)} "
                "categories failed",
            )

            for code, message in failed_codes.items():
                log(
                    "FAILED",
                    f"{code}: {message}",
                )

        print("=" * 72)

    finally:
        client.close()


if __name__ == "__main__":
    try:
        main()

    except KeyboardInterrupt:
        print()

        log(
            "STOP",
            "Interrupted by user",
        )

        sys.exit(130)

    except Exception as exc:
        print()

        log(
            "ERROR",
            str(exc),
        )

        sys.exit(1)
