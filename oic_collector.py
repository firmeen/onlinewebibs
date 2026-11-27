#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC XMLA COLLECTOR
Professional Terminal UI | Textual + Rich | Excel Only

INSTALL:
    python -m pip install -U requests textual rich

RUN TERMINAL UI:
    python oic_collector.py
    python oic_collector.py --cli

RUN HEADLESS:
    python oic_collector.py --test-only
    python oic_collector.py --headless
    python oic_collector.py --headless --codes AH FI

SHORTCUTS:
    Ctrl+R  Reload .env
    Ctrl+T  Test XMLA API
    Ctrl+D  Start collection
    Ctrl+S  Stop collection
    Ctrl+O  Open output directory
    Ctrl+Q  Quit

AUTHENTICATION:
    - Terminal UI opens before any network authentication.
    - JWT expiration is inspected locally and updated live.
    - Cookie validity is checked asynchronously using GenToken.
    - Editing Cookie triggers automatic revalidation.
    - Invalid/expired sessions never close the Terminal UI.
    - Successful GenToken requests update JWT and Cookie.
    - Authentication updates are saved to .env.
    - Session status is periodically revalidated.
    - No Selenium or automatic login is required.

ENVIRONMENT:
    OIC_BEARER_TOKEN=...
    OIC_COOKIE=...
    OIC_OUTPUT_DIR=output

OUTPUT:
    output/oic_insurance_YYYYMMDD_HHMMSS.xlsx

Single Python file.
No SQL Server.
No additional Python source files.
"""

from __future__ import annotations

import sys
from pathlib import Path

# Ensure local project modules take priority.
PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

import argparse
import base64
import hashlib
import json
import os
import re
import sys
import threading
import time
import zipfile
import uuid
import subprocess
import shutil

from oic_query import (PROFILES as QUERY_PROFILES, SCHEMA_VERSION, load_catalog, discover_fields,
                       validate_query, build_mdx as selected_mdx, query_hash,
                        cube_status, cube_profile, cube_options, CUBE_TYPES,
                        verify_field, catalog_verification_summary,
                        export_verified_catalogs)

from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from urllib.parse import urlparse
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape as xml_escape

import requests


# ============================================================
# CONFIGURATION
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
ENV_PATH = BASE_DIR / ".env"

BASE_URL = "https://onlinewebibs.oic.or.th"

PAGE_PATH = (
    "/WebForm/Dashboard/Pivot/"
    "DynamicPivotCube.aspx"
)

TOKEN_URL = (
    BASE_URL
    + "/WebForm/Dashboard/Pivot/"
    "PivotCube.asmx/GenToken"
)

XMLA_URL = (
    BASE_URL
    + "/xmla-proxy/api/xmla-proxy"
)

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"
MD_NS = "urn:schemas-microsoft-com:xml-analysis:mddataset"

DEFAULT_TIMEOUT = 180
DEFAULT_RETRIES = 3

SESSION_RECHECK_SECONDS = 120
COOKIE_DEBOUNCE_SECONDS = 1.5

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)


# ============================================================
# INSURANCE PROFILES
# ============================================================

# Shared multi-cube profiles, including MR's unconfirmed YEAR.
PROFILES = {
    code: dict(name=p["name"], **{"class": p["cls"]}, catalog=p["catalog"],
               cube=p["cube"], column=p["year"] or "[VW PV MR AGGREGATE FY].[STATUS COMPANY PASS]",
               row1=p["row1"], row2=p["row2"], measures=p["measures"])
    for code, p in QUERY_PROFILES.items()
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
# ENVIRONMENT MANAGER
# ============================================================

ENV_LOCK = threading.RLock()


def normalize_token(value: str) -> str:
    value = (value or "").strip()

    if value.lower().startswith("bearer "):
        value = value[7:].strip()

    if value.startswith("REPLACE_WITH"):
        return ""

    return value


def normalize_cookie(value: str) -> str:
    value = (value or "").strip()

    if value.lower().startswith("cookie:"):
        value = value[7:].strip()

    if value.startswith("OIC_COOKIE="):
        value = value[len("OIC_COOKIE="):].strip()

    if (
        len(value) >= 2
        and value[0] == value[-1]
        and value[0] in ('"', "'")
    ):
        value = value[1:-1]

    return value


def read_env() -> dict[str, str]:
    with ENV_LOCK:
        if not ENV_PATH.exists():
            return {}

        result = {}

        content = ENV_PATH.read_text(
            encoding="utf-8-sig"
        )

        for raw_line in content.splitlines():
            line = raw_line.strip()

            if not line or line.startswith("#"):
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

            result[key] = value

        return result


def update_env(updates: dict[str, str]) -> None:
    """
    Update selected keys while preserving unrelated
    configuration such as USER and PASSWORD.
    """

    with ENV_LOCK:
        if ENV_PATH.exists():
            original = ENV_PATH.read_text(
                encoding="utf-8-sig"
            )
        else:
            original = ""

        lines = original.splitlines()
        output_lines = []
        completed = set()

        for line in lines:
            match = re.match(
                r"^(\s*)([A-Za-z_][A-Za-z0-9_]*)(\s*=)",
                line,
            )

            if not match:
                output_lines.append(line)
                continue

            key = match.group(2)

            if key not in updates:
                output_lines.append(line)
                continue

            if key in completed:
                continue

            value = str(updates[key])

            value = value.replace(
                "\r", ""
            ).replace(
                "\n", ""
            )

            if value != value.strip():
                value = json.dumps(
                    value,
                    ensure_ascii=False,
                )

            output_lines.append(
                f"{match.group(1)}"
                f"{key}"
                f"{match.group(3)}"
                f"{value}"
            )

            completed.add(key)

        for key, raw_value in updates.items():
            if key in completed:
                continue

            value = str(raw_value).replace(
                "\r", ""
            ).replace(
                "\n", ""
            )

            output_lines.append(
                f"{key}={value}"
            )

        temporary = ENV_PATH.with_name(
            ".env.oic.tmp"
        )

        try:
            temporary.write_text(
                "\n".join(output_lines).rstrip("\n")
                + "\n",
                encoding="utf-8",
            )

            os.replace(
                temporary,
                ENV_PATH,
            )

        finally:
            temporary.unlink(
                missing_ok=True
            )


def env_signature():
    try:
        stat = ENV_PATH.stat()
        return (
            stat.st_mtime_ns,
            stat.st_size,
        )

    except OSError:
        return None


def resolve_output_path(raw: str) -> Path:
    raw = (raw or "").strip() or "output"

    path = Path(raw).expanduser()

    if not path.is_absolute():
        path = BASE_DIR / path

    return path.resolve()


# ============================================================
# JWT INSPECTOR
# ============================================================

def decode_jwt_payload(token: str) -> dict:
    token = normalize_token(token)

    try:
        parts = token.split(".")

        if len(parts) != 3:
            return {}

        encoded = parts[1]

        encoded += "=" * (
            -len(encoded) % 4
        )

        decoded = base64.urlsafe_b64decode(
            encoded.encode("ascii")
        )

        payload = json.loads(
            decoded.decode("utf-8")
        )

        return (
            payload
            if isinstance(payload, dict)
            else {}
        )

    except Exception:
        return {}


def format_remaining(seconds: int) -> str:
    if seconds <= 0:
        return "00:00:00"

    hours, remainder = divmod(
        seconds,
        3600,
    )

    minutes, seconds = divmod(
        remainder,
        60,
    )

    return (
        f"{hours:02}:"
        f"{minutes:02}:"
        f"{seconds:02}"
    )


def jwt_info(token: str) -> dict:
    token = normalize_token(token)

    result = {
        "status": "MISSING",
        "issued": "—",
        "expires": "—",
        "lifetime": "—",
        "remaining": "—",
        "remaining_seconds": 0,
        "exp": 0,
    }

    if not token:
        return result

    payload = decode_jwt_payload(token)

    try:
        iat = int(payload["iat"])
        exp = int(payload["exp"])

        issued = datetime.fromtimestamp(
            iat
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )

        expires = datetime.fromtimestamp(
            exp
        ).strftime(
            "%Y-%m-%d %H:%M:%S"
        )

    except (
        KeyError,
        TypeError,
        ValueError,
        OverflowError,
        OSError,
    ):
        result["status"] = "INVALID"
        return result

    now = int(time.time())

    remaining = max(
        0,
        exp - now,
    )

    lifetime = (
        exp - iat
    ) / 60

    result.update({
        "status": (
            "VALID"
            if exp > now
            else "EXPIRED"
        ),
        "issued": issued,
        "expires": expires,
        "lifetime": f"{lifetime:g} min",
        "remaining": format_remaining(
            remaining
        ),
        "remaining_seconds": remaining,
        "exp": exp,
    })

    return result


def jwt_is_valid_time(
    token: str,
    margin_seconds: int = 120,
) -> bool:
    details = jwt_info(token)

    return (
        details["status"] == "VALID"
        and details["remaining_seconds"] > margin_seconds
    )


# ============================================================
# MDX QUERY BUILDER
# ============================================================

def build_mdx(profile: dict) -> str:
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

        dx_rows = (
            f"Union({subtotal},{details})"
        )

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


# ============================================================
# SOAP XMLA BUILDER
# ============================================================

def build_soap(profile: dict) -> bytes:
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

def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def child_text(
    element: ET.Element,
    name: str,
) -> str:
    for child in element:
        if local_name(child.tag) == name:
            return child.text or ""

    return ""


def parse_xmla_response(
    content: bytes,
    code: str,
) -> ET.Element:
    try:
        root = ET.fromstring(content)

    except ET.ParseError as exc:
        raise RuntimeError(
            f"{code}: Invalid XML response: {exc}"
        ) from exc

    for element in root.iter():
        if local_name(element.tag) == "Fault":
            message = " ".join(
                (node.text or "").strip()
                for node in element.iter()
                if node.text
            )

            raise RuntimeError(
                f"{code}: SOAP Fault: {message[:400]}"
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
            f"{code}: XMLA multidimensional dataset not found"
        )

    return mdroot


# ============================================================
# AUTHENTICATION ERRORS
# ============================================================

class OICAuthenticationError(RuntimeError):
    pass


class OICSessionExpired(OICAuthenticationError):
    pass


class OICStopped(RuntimeError):
    pass


# ============================================================
# OIC HTTP CLIENT
# ============================================================

class OICClient:
    def __init__(
        self,
        cookie: str = "",
        token: str = "",
        timeout: int = DEFAULT_TIMEOUT,
        retries: int = DEFAULT_RETRIES,
        on_auth=None,
        on_log=None,
        stop_event=None,
    ):
        self.session = requests.Session()

        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Accept-Language": "th-TH,th;q=0.9",
        })

        self.timeout = timeout
        self.retries = max(2, retries)

        self.token = normalize_token(token)

        self.on_auth = on_auth
        self.on_log = on_log

        self.stop_event = (
            stop_event
            if stop_event is not None
            else threading.Event()
        )

        self.load_cookie(
            normalize_cookie(cookie)
        )

    def load_cookie(self, raw: str) -> None:
        hostname = urlparse(
            BASE_URL
        ).hostname

        for part in raw.split(";"):
            name, separator, value = (
                part.strip().partition("=")
            )

            if not separator or not name:
                continue

            self.session.cookies.set(
                name.strip(),
                value.strip(),
                domain=hostname,
                path="/",
                secure=True,
            )

    def cookie_header(self) -> str:
        hostname = urlparse(
            BASE_URL
        ).hostname

        values = {}

        for cookie in self.session.cookies:
            domain = cookie.domain.lstrip(".")

            if domain == hostname:
                values[cookie.name] = cookie.value

        return "; ".join(
            f"{key}={value}"
            for key, value in values.items()
        )

    def page_path(self, code: str) -> str:
        profile = PROFILES[code]

        return (
            f"{PAGE_PATH}"
            f"?Class={profile['class']}"
        )

    def check_stop(self) -> None:
        if self.stop_event.is_set():
            raise OICStopped(
                "Collection stopped by user"
            )

    def log(
        self,
        message: str,
        level: str = "info",
    ) -> None:
        if self.on_log:
            self.on_log(
                message,
                level,
            )

    def gen_headers(self, code: str) -> dict:
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
                BASE_URL + self.page_path(code)
            ),
            "X-Requested-With": "XMLHttpRequest",
        }

    def request_token(self, code: str) -> str:
        self.check_stop()

        self.log(
            f"GenToken / requesting JWT for {code}"
        )

        payload = {
            "_tk": "Init",
            "_pageUrl": self.page_path(code),
        }

        try:
            response = self.session.post(
                TOKEN_URL,
                json=payload,
                headers=self.gen_headers(code),
                timeout=(15, 45),
                allow_redirects=False,
            )

        except requests.RequestException as exc:
            raise OICAuthenticationError(
                "GenToken network error: "
                f"{type(exc).__name__}: {exc}"
            ) from exc

        if response.status_code in (
            301,
            302,
            303,
            307,
            308,
        ):
            raise OICSessionExpired(
                "SESSION_EXPIRED / redirected to login"
            )

        if response.status_code in (401, 403):
            raise OICAuthenticationError(
                f"GenToken HTTP {response.status_code} "
                "/ authorization rejected"
            )

        if response.status_code != 200:
            raise OICAuthenticationError(
                f"GenToken HTTP {response.status_code}"
            )

        try:
            data = response.json()

        except ValueError as exc:
            raise OICAuthenticationError(
                "GenToken returned non-JSON content"
            ) from exc

        token = (
            data.get("d")
            if isinstance(data, dict)
            else None
        )

        if not isinstance(token, str):
            raise OICAuthenticationError(
                "GenToken response has no JWT field 'd'"
            )

        token = normalize_token(token)

        if token == "SESSION_EXPIRED":
            raise OICSessionExpired(
                "SESSION_EXPIRED"
            )

        if len(token.split(".")) != 3:
            status = (
                token
                if re.fullmatch(
                    r"[A-Z_]{3,80}",
                    token,
                )
                else "UNEXPECTED_RESPONSE"
            )

            raise OICAuthenticationError(
                f"GenToken: {status}"
            )

        information = jwt_info(token)

        if information["status"] != "VALID":
            raise OICAuthenticationError(
                "GenToken returned invalid/expired JWT"
            )

        self.token = token

        current_cookie = self.cookie_header()

        if self.on_auth:
            self.on_auth(
                token,
                current_cookie,
            )

        self.log(
            "GenToken SUCCESS / "
            f"expires {information['expires']}",
            "success",
        )

        return token

    def ensure_token(
        self,
        code: str,
        force: bool = False,
    ) -> str:
        if (
            not force
            and jwt_is_valid_time(
                self.token,
                margin_seconds=120,
            )
        ):
            return self.token

        return self.request_token(code)

    def fetch(
        self,
        code: str,
    ) -> tuple[ET.Element, int]:
        self.check_stop()

        self.ensure_token(code)

        payload = build_soap(
            PROFILES[code]
        )

        for attempt in range(
            1,
            self.retries + 1,
        ):
            self.check_stop()

            headers = {
                "Accept": "text/plain, */*; q=0.01",
                "Content-Type": "text/xml",
                "Origin": BASE_URL,
                "Referer": (
                    BASE_URL + self.page_path(code)
                ),
                "X-Requested-With": "XMLHttpRequest",
                "Authorization": (
                    f"Bearer {self.token}"
                ),
            }

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

                delay = min(
                    2 ** attempt,
                    15,
                )

                self.log(
                    f"{code}: network retry "
                    f"{attempt}/{self.retries}",
                    "warning",
                )

                if self.stop_event.wait(delay):
                    self.check_stop()

                continue

            status = response.status_code

            if status == 200:
                return (
                    parse_xmla_response(
                        response.content,
                        code,
                    ),
                    len(response.content),
                )

            if status in (401, 403):
                if attempt == 1:
                    self.log(
                        f"{code}: HTTP {status} / "
                        "refreshing JWT",
                        "warning",
                    )

                    self.ensure_token(
                        code,
                        force=True,
                    )

                    continue

                raise OICAuthenticationError(
                    f"{code}: HTTP {status} "
                    "after JWT refresh"
                )

            if 300 <= status < 400:
                raise OICSessionExpired(
                    f"{code}: HTTP {status} / login redirect"
                )

            if (
                status in (
                    429,
                    500,
                    502,
                    503,
                    504,
                )
                and attempt < self.retries
            ):
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

                self.log(
                    f"{code}: HTTP {status} / "
                    f"retry in {delay}s",
                    "warning",
                )

                if self.stop_event.wait(delay):
                    self.check_stop()

                continue

            raise RuntimeError(
                f"{code}: XMLA HTTP {status}"
            )

        raise RuntimeError(
            f"{code}: XMLA request failed"
        )

    def close(self) -> None:
        self.session.close()


# ============================================================
# XMLA DATA EXTRACTION
# ============================================================

def members_from_tuple(
    element: ET.Element,
) -> list[dict]:
    result = []

    for member in element.findall(
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
            "is_all": unique_name.endswith(
                "[All]"
            ),
        })

    return result


def empty_member() -> dict:
    return {
        "dimension": "",
        "caption": "",
        "unique_name": "",
        "is_all": False,
    }


def parse_dataset(
    mdroot: ET.Element,
    code: str,
) -> list[dict]:
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

    if (
        axis0_element is None
        or axis1_element is None
    ):
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
    ).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )

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
                f"{code}: Invalid CellOrdinal {ordinal}"
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
# XLSX WRITER
# ============================================================

def excel_column_letter(number: int) -> str:
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


def excel_cell(
    address: str,
    value,
    header: bool = False,
    numeric: bool = False,
) -> str:
    style = (
        ' s="1"'
        if header
        else ""
    )

    if numeric:
        number = excel_number(value)

        if number is not None:
            return (
                f'<c r="{address}"{style}>'
                f"<v>{number}</v>"
                "</c>"
            )

    text = xml_escape(
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


def write_sheet(
    archive: zipfile.ZipFile,
    sheet_index: int,
    records: list[dict],
    stop_event=None,
) -> None:
    sheet_path = (
        f"xl/worksheets/sheet{sheet_index}.xml"
    )

    with archive.open(
        sheet_path,
        "w",
        force_zip64=True,
    ) as stream:

        def write_xml(value: str) -> None:
            stream.write(
                value.encode("utf-8")
            )

        write_xml(
            '<?xml version="1.0" encoding="UTF-8" '
            'standalone="yes"?>'
            '<worksheet xmlns="http://schemas.'
            'openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetViews><sheetView workbookViewId="0" '
            'showGridLines="0"/></sheetViews>'
            '<sheetFormatPr defaultRowHeight="15"/>'
            "<sheetData>"
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
            f'<row r="1">{header_cells}</row>'
        )

        for row_number, record in enumerate(
            records,
            2,
        ):
            if (
                stop_event is not None
                and row_number % 100 == 0
                and stop_event.is_set()
            ):
                raise OICStopped(
                    "Excel export stopped"
                )

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
                f'<row r="{row_number}">'
                f'{"".join(cells)}'
                "</row>"
            )

        last_row = len(records) + 1

        write_xml(
            "</sheetData>"
            f'<autoFilter ref="A1:T{last_row}"/>'
            "</worksheet>"
        )


def export_excel(
    data_by_code: dict[str, list[dict]],
    output_dir: Path,
    on_log=None,
    stop_event=None,
) -> Path:
    if not data_by_code:
        raise RuntimeError(
            "No data available for Excel export"
        )

    output_dir.mkdir(
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

    output_path = output_dir / filename

    sheets = list(
        data_by_code.items()
    )

    content_types = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/'
        'package/2006/content-types">',
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.'
        'relationships+xml"/>',
        '<Default Extension="xml" '
        'ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.'
        'spreadsheetml.styles+xml"/>',
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
        f'<sheet name="{xml_escape(code)}" '
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
        '<xf numFmtId="0" fontId="0" fillId="0" '
        'borderId="0"/></cellStyleXfs>'
        '<cellXfs count="2">'
        '<xf numFmtId="0" fontId="0" fillId="0" '
        'borderId="0"/>'
        '<xf numFmtId="0" fontId="1" fillId="1" '
        'borderId="0" applyFont="1" applyFill="1"/>'
        "</cellXfs>"
        "</styleSheet>"
    )

    try:
        with zipfile.ZipFile(
            output_path,
            "w",
            compression=zipfile.ZIP_DEFLATED,
            compresslevel=6,
        ) as archive:

            archive.writestr(
                "[Content_Types].xml",
                "".join(content_types),
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
                "".join(workbook_relationships),
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
                if (
                    stop_event is not None
                    and stop_event.is_set()
                ):
                    raise OICStopped(
                        "Excel export stopped"
                    )

                if on_log:
                    on_log(
                        f"Writing Excel {code}: "
                        f"{len(records):,} rows",
                        "info",
                    )

                write_sheet(
                    archive,
                    index,
                    records,
                    stop_event=stop_event,
                )

        return output_path

    except Exception:
        output_path.unlink(
            missing_ok=True
        )
        raise


# ============================================================
# COLLECTION ENGINE
# ============================================================

def collect(
    codes: list[str],
    cookie: str,
    token: str,
    output_dir: Path,
    timeout: int = DEFAULT_TIMEOUT,
    retries: int = DEFAULT_RETRIES,
    test_only: bool = False,
    stop_event=None,
    on_log=None,
    on_auth=None,
    on_result=None,
    on_progress=None,
) -> dict:
    """
    Run the original XMLA data extraction workflow.

    Session and authentication failures are raised to
    the caller. The Terminal UI handles them without
    closing the application.
    """

    client = OICClient(
        cookie=cookie,
        token=token,
        timeout=timeout,
        retries=retries,
        on_auth=on_auth,
        on_log=on_log,
        stop_event=stop_event,
    )

    data_by_code = {}
    failed_codes = {}

    selected_codes = (
        codes[:1]
        if test_only
        else codes
    )

    try:
        first_code = selected_codes[0]

        client.request_token(
            first_code
        )

        for index, code in enumerate(
            selected_codes,
            1,
        ):
            client.check_stop()

            if on_result:
                on_result(
                    code,
                    "RUNNING",
                    0,
                    0,
                )

            if on_log:
                on_log(
                    f"{code}: requesting XMLA cube",
                    "info",
                )

            try:
                mdroot, response_size = client.fetch(
                    code
                )

                records = parse_dataset(
                    mdroot,
                    code,
                )

                data_by_code[code] = records

                if on_result:
                    on_result(
                        code,
                        "SUCCESS",
                        len(records),
                        response_size,
                    )

                if on_log:
                    on_log(
                        f"{code}: "
                        f"{len(records):,} rows / "
                        f"{response_size:,} bytes",
                        "success",
                    )

            except OICStopped:
                raise

            except OICAuthenticationError:
                raise

            except Exception as exc:
                failed_codes[code] = str(exc)

                if on_result:
                    on_result(
                        code,
                        "FAILED",
                        0,
                        0,
                    )

                if on_log:
                    on_log(
                        f"{code}: {exc}",
                        "error",
                    )

                if test_only:
                    raise

            if on_progress:
                on_progress(
                    index,
                    len(selected_codes),
                )

        total_rows = sum(
            len(records)
            for records in data_by_code.values()
        )

        if test_only:
            return {
                "rows": total_rows,
                "errors": failed_codes,
                "output": None,
                "completed": list(
                    data_by_code
                ),
            }

        client.check_stop()

        if not data_by_code:
            raise RuntimeError(
                "No successful XMLA datasets to export"
            )

        if on_log:
            on_log(
                "Creating Excel workbook...",
                "info",
            )

        output_path = export_excel(
            data_by_code,
            output_dir,
            on_log=on_log,
            stop_event=stop_event,
        )

        return {
            "rows": total_rows,
            "errors": failed_codes,
            "output": output_path,
            "completed": list(
                data_by_code
            ),
        }

    finally:
        client.close()


# ============================================================
# HEADLESS MODE
# ============================================================

def run_headless(args) -> int:
    config = read_env()

    cookie = normalize_cookie(
        config.get(
            "OIC_COOKIE",
            "",
        )
    )

    token = normalize_token(
        config.get(
            "OIC_BEARER_TOKEN",
            "",
        )
    )

    if not cookie:
        print(
            "ERROR: OIC_COOKIE is missing",
            file=sys.stderr,
        )
        return 1

    output_dir = resolve_output_path(
        config.get(
            "OIC_OUTPUT_DIR",
            "output",
        )
    )

    def logger(
        message: str,
        level: str = "info",
    ) -> None:
        now = datetime.now().strftime(
            "%H:%M:%S"
        )

        print(
            f"[{now}] [{level.upper()}] {message}",
            flush=True,
        )

    def auth_updated(
        new_token: str,
        new_cookie: str,
    ) -> None:
        update_env({
            "OIC_BEARER_TOKEN": new_token,
            "OIC_COOKIE": new_cookie,
        })

    try:
        result = collect(
            codes=args.codes,
            cookie=cookie,
            token=token,
            output_dir=output_dir,
            timeout=args.timeout,
            retries=args.retries,
            test_only=args.test_only,
            on_log=logger,
            on_auth=auth_updated,
        )

        logger(
            f"Completed: {result['rows']:,} rows",
            "success",
        )

        if result["output"]:
            logger(
                f"Excel: {result['output']}",
                "success",
            )

        return (
            2
            if result["errors"]
            else 0
        )

    except KeyboardInterrupt:
        logger(
            "Interrupted",
            "warning",
        )
        return 130

    except Exception as exc:
        logger(
            str(exc),
            "error",
        )
        return 1



# ============================================================
# APPLIED QUERY STORAGE — CANONICAL + LEGACY COMPATIBILITY
# ============================================================
# Historical local branches stored each entry in "config" and optionally
# "plan". This branch stores "query" and "task_plan". Both are accepted,
# but only the canonical format is written after backing up the original.

from contextlib import contextmanager

APPLIED_STORAGE_VERSION = 2


@contextmanager
def applied_list_write_lock(path: Path):
    """Serialize Main App and Designer read/modify/write transactions."""
    path.parent.mkdir(parents=True, exist_ok=True)
    lock_path = path.with_name(path.name + ".write.lock")
    with lock_path.open("a+b") as handle:
        handle.seek(0, os.SEEK_END)
        if handle.tell() == 0:
            handle.write(b"\0")
            handle.flush()
        acquired = False
        start = time.monotonic()
        try:
            while not acquired:
                try:
                    handle.seek(0)
                    if os.name == "nt":
                        import msvcrt
                        msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
                    else:
                        import fcntl
                        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                    acquired = True
                except (OSError, BlockingIOError):
                    if time.monotonic() - start >= 10:
                        raise TimeoutError("Applied Query List is busy; retry APPLY")
                    time.sleep(0.08)
            yield
        finally:
            if acquired:
                handle.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(handle.fileno(), fcntl.LOCK_UN)


def _applied_query_number(value) -> int | None:
    if not isinstance(value, str) or not re.fullmatch(r"Q[0-9]{3,}", value):
        return None
    return int(value[1:])


def applied_query_identity(config: dict) -> str:
    """Order-sensitive fingerprint; ignores worker count and run history."""
    values = {
        key: config.get(key, [] if key in ("x", "y", "measures", "filters") else None)
        for key in ("code", "cube", "x", "y", "measures", "filters")
    }
    return json.dumps(values, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def read_applied_queries(path: Path) -> tuple[dict, bool]:
    """Read old 'config'/'plan' and current 'query'/'task_plan' schemas.

    Never modify the file during reading. An invalid legacy record is
    reported with its index rather than discarded or silently overwritten.
    """
    from oic_parallel_collector import build_selected_task_plan

    try:
        original = json.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return {
            "storage_version": APPLIED_STORAGE_VERSION,
            "revision": uuid.uuid4().hex,
            "next_query_sequence": 1,
            "queries": [],
        }, False

    if isinstance(original, list):
        raw_queries = original
        document = {}
    elif isinstance(original, dict):
        document = original
        raw_queries = document.get("queries")
        if raw_queries is None and isinstance(document.get("query"), dict):
            raw_queries = [document]
        if not isinstance(raw_queries, list):
            raise ValueError("Invalid applied_queries.json: expected a queries array")
    else:
        raise ValueError("Invalid applied_queries.json: expected an object or array")

    reserved = set()
    for raw_entry in raw_queries:
        if not isinstance(raw_entry, dict):
            continue
        possible_query = raw_entry.get("query") or raw_entry.get("config")
        if not isinstance(possible_query, dict):
            possible_query = raw_entry
        for candidate in (raw_entry.get("query_id"), possible_query.get("query_id")):
            if _applied_query_number(candidate) is not None:
                reserved.add(candidate)

    stored_sequence = document.get(
        "next_query_sequence", document.get("next_id", 1))
    try:
        sequence = max(1, int(stored_sequence))
    except (TypeError, ValueError):
        sequence = 1
    sequence = max(sequence, (max(
        (_applied_query_number(qid) for qid in reserved), default=0) or 0) + 1)

    normalized = []
    seen = set()
    for index, raw_entry in enumerate(raw_queries, start=1):
        if not isinstance(raw_entry, dict):
            raise ValueError(f"Applied Query item #{index} is not an object")

        query = raw_entry.get("query")
        if not isinstance(query, dict):
            query = raw_entry.get("config")
        if not isinstance(query, dict) and all(
            key in raw_entry for key in ("code", "cube", "x", "y", "measures")
        ):
            query = raw_entry
        if not isinstance(query, dict):
            raise ValueError(
                f"Applied Query item #{index} has no query/config; "
                "original file was not changed"
            )

        config = dict(query)
        if "cube_type" not in config:
            cube = str(config.get("cube") or "")
            config["cube_type"] = next(
                (kind for kind in ("UY", "AY") if cube.endswith("_" + kind)), "FY")

        query_id = raw_entry.get("query_id") or config.get("query_id")
        if _applied_query_number(query_id) is None or query_id in seen:
            while f"Q{sequence:03d}" in reserved or f"Q{sequence:03d}" in seen:
                sequence += 1
            query_id = f"Q{sequence:03d}"
            sequence += 1

        config["query_id"] = query_id
        try:
            validate_query(config)
        except (ValueError, TypeError, KeyError) as exc:
            raise ValueError(
                f"Applied Query {query_id} (item #{index}) cannot be migrated: {exc}; "
                "original file was not changed"
            ) from exc

        # Old plans have different allocation keys; always regenerate them
        # from the actual validated query if stale or structurally different.
        expected_plan = build_selected_task_plan(config)
        plan = raw_entry.get("task_plan")
        if (not isinstance(plan, dict) or
                plan.get("query_hash") != expected_plan["query_hash"] or
                plan.get("assignments") != expected_plan["assignments"]):
            plan = expected_plan

        normalized.append({
            "query_id": query_id,
            "query": config,
            "task_plan": plan,
            "applied_at": raw_entry.get("applied_at")
            if isinstance(raw_entry.get("applied_at"), str)
            else datetime.now(timezone.utc).isoformat(),
        })
        seen.add(query_id)

    sequence = max(sequence, max(
        (_applied_query_number(value) for value in seen), default=0) + 1)
    revision = document.get("revision")
    if not isinstance(revision, str) or not revision:
        revision = uuid.uuid4().hex
    result = {
        "storage_version": APPLIED_STORAGE_VERSION,
        "revision": revision,
        "next_query_sequence": sequence,
        "queries": normalized,
    }
    return result, result != original


def save_applied_queries(path: Path, document: dict, *,
                         backup_legacy: bool = False) -> None:
    """Atomic replacement; first migration keeps the byte-for-byte old JSON."""
    path.parent.mkdir(parents=True, exist_ok=True)
    if backup_legacy and path.exists():
        backup = path.with_name(
            f"{path.stem}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}_"
            f"{uuid.uuid4().hex[:8]}{path.suffix}")
        shutil.copy2(path, backup)
    payload = dict(document)
    payload["storage_version"] = APPLIED_STORAGE_VERSION
    payload["revision"] = uuid.uuid4().hex
    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)
    document.update(payload)


# ============================================================
# TERMINAL APPLICATION
# ============================================================

def launch_terminal_ui(args, return_app: bool = False):
    try:
        from rich.markup import escape as rich_escape
        from rich.text import Text

        from textual import work
        from textual.binding import Binding
        from textual.css.query import NoMatches
        from textual.app import App, ComposeResult
        from textual.containers import Horizontal, Vertical
        from textual.widgets import (
            Button,
            Checkbox,
            DataTable,
            Footer,
            Header,
            Input,
            Label,
            ProgressBar,
            RichLog,
            Static,
            Select,
            Tree,
            OptionList,
        )

    except ImportError as exc:
        raise RuntimeError(
            "Install dependencies: "
            "python -m pip install -U requests textual rich"
        ) from exc

    class OICApp(App):
        TITLE = "OIC XMLA Collector"
        SUB_TITLE = "Insurance Data Workspace"

        BINDINGS = [
            ("ctrl+r", "reload_config", "Reload"),
            ("ctrl+t", "test_api", "Test"),
            ("ctrl+d", "start_collection", "Collect"),
            ("ctrl+s", "stop_collection", "Stop"),
            ("ctrl+o", "open_output", "Output"),
            ("ctrl+q", "quit", "Quit"),
            Binding("a", "field_to_x", "X Axis", show=False),
            Binding("s", "field_to_y", "Y Axis", show=False),
        ]

        CSS = """
        Screen {
            background: #0d141b;
            color: #dce5e9;
        }

        Header {
            background: #17232e;
            color: #dce5e9;
        }

        Footer {
            background: #17232e;
            color: #9fb0ba;
        }

        #workspace {
            height: 1fr;
            padding: 0 2 1 2;
            overflow-y: auto;
        }

        #brand {
            height: 3;
            margin-top: 1;
            padding: 0 2;
            background: #17232e;
            border-left: thick #ac9b79;
            color: #e9e7df;
            content-align: left middle;
            text-style: bold;
        }

        /* AUTHENTICATION STATUS */

        #auth_row {
            height: 8;
            margin-top: 1;
        }

        .auth_panel {
            width: 1fr;
            margin-right: 1;
            padding: 0 2;
            background: #16232e;
            border: solid #304352;
        }

        .auth_heading {
            height: 1;
            color: #c9b996;
            text-style: bold;
        }

        .auth_content {
            height: 1fr;
            color: #d5e1e6;
        }

        /* CREDENTIALS */

        #credentials {
            height: 6;
            margin-top: 1;
        }

        .credential_group {
            width: 1fr;
            margin-right: 1;
        }

        .input_label {
            height: 1;
            padding-left: 1;
            color: #96a8b4;
            text-style: bold;
        }

        .input_hint {
            height: 1;
            padding-left: 1;
            color: #8599a5;
        }

        Input {
            height: 3;
            background: #16232e;
            color: #e5ebed;
            border: tall #304352;
        }

        Input:focus {
            border: tall #ac9b79;
        }

        /* ACTIONS */

        #toolbar {
            height: 4;
            margin-top: 1;
            align: left middle;
            overflow-x: auto;
        }

        #toolbar Button {
            height: 3;
            margin-right: 1;
            text-style: bold;
        }

        Button {
            background: #2c4c5d;
            color: #e6ebed;
        }

        Button:hover {
            background: #375c6c;
        }

        Button:focus {
            background: #416778;
            color: #ffffff;
        }

        Button:disabled {
            opacity: 0.45;
        }

        #check {
            background: #365569;
        }

        #start {
            background: #23685f;
        }

        #stop {
            background: #633d40;
        }

        /* CONFIGURATION */

        #config_row {
            height: 5;
            margin-top: 1;
        }

        #output_group {
            width: 1fr;
            margin-right: 2;
        }

        #category_group {
            width: 2fr;
        }

        #categories {
            height: 3;
            align: left middle;
        }

        #categories Checkbox {
            width: 1fr;
            margin-right: 1;
            color: #dce6e9;
        }

        /* SUMMARY */

        #summary {
            height: 3;
            margin-top: 1;
            padding: 0 2;
            background: #17232e;
            border: solid #293c49;
            color: #d7e0e4;
            content-align: left middle;
        }

        /* DATA CATALOG */

        #catalog_panel {
            height: 1fr;
            min-height: 9;
            margin-top: 1;
            background: #111c25;
            border: solid #293c49;
        }

        #catalog_title {
            height: 2;
            padding: 0 1;
            background: #17232e;
            color: #c9b996;
            content-align: left middle;
            text-style: bold;
        }

        #table {
            height: 1fr;
            background: #111c25;
        }

        DataTable > .datatable--header {
            background: #20303d;
            color: #e4edf1;
            text-style: bold;
        }

        DataTable > .datatable--cursor {
            background: #345265;
            color: #ffffff;
        }

        DataTable > .datatable--odd-row {
            background: #15232e;
        }

        /* PROGRESS */

        #progress_panel {
            height: 4;
            margin-top: 1;
            padding: 0 1;
            border: solid #293c49;
            background: #101a22;
        }

        #status {
            height: 1;
            color: #c7d3da;
        }

        #progress {
            height: 2;
        }

        /* LOG */

        #activity_panel {
            height: 8;
            margin-top: 1;
            border: solid #293c49;
            background: #101a22;
        }

        #activity_title {
            height: 1;
            padding-left: 1;
            background: #17232e;
            color: #c9b996;
            text-style: bold;
        }

        #activity {
            height: 1fr;
            padding: 0 1;
            background: #101a22;
        }

        #categories, #catalog_panel { display: none; }
        #cube_select { width: 1fr; }
        #applied_panel {
            height: 15;
            min-height: 9;
            margin-top: 1;
            border: solid #304352;
            background: #16232e;
            padding: 0 1;
        }
        #applied_title { height: 1; color: #c9b996; text-style: bold; }
        #applied_query_list { height: 1fr; overflow-y: auto; }
        #applied_actions { height: 3; }
        #applied_actions Button { min-width: 12; margin-right: 1; }
        #selected_query_summary { height: 2; padding: 0 1; color: #9fb0ba; }
        #catalog_status { height: auto; min-height: 3; margin-top: 1; padding: 0 2;
                           border: solid #293c49; color: #9fb0ba; }
        #activity_panel { height: 1fr; min-height: 8; }
        #designer_area { height: 1fr; min-height: 12; }
        #field_panel { width: 1fr; padding: 0 1; }
        #axis_panel { width: 1fr; padding: 0 1; }
        #field_table { height: 1fr; }
        .designer_title { height: 1; color: #c9b996; text-style: bold; }
        .axis_bar { height: 3; }
        .axis_bar Button { margin-right: 1; }
        #values_list { height: 3; overflow-y: auto; }
        #preview_text { height: 5; overflow-y: auto; color: #b9c9d2; }
        """

        # ====================================================
        # STATE
        # ====================================================

        def __init__(self):
            super().__init__()

            self.config = {}
            self.selected_code = "MV"
            self.catalog_snapshot = {}
            self.prepare_running = False
            self.field_catalog = {}
            self.selected_field = None
            self.field_rows = {}
            self.axis_x = []
            self.axis_y = []
            self.selected_measures = []
            self.active_run = None
            self.last_run_path = None
            self.applied_config = None
            self.applied_queries = []
            self.selection_revision = None
            self.batch_process = None
            self.query_states = {}
            self.force_confirm_until = 0
            self.catalog_running = False
            self.catalog_statuses = {}

            self.busy = False
            self.stop_event = threading.Event()

            self.cookie_state = "NOT CHECKED"
            self.cookie_message = ""
            self.cookie_checked_at = "—"

            self.active_token = ""

            self.auth_sequence = 0
            self.auth_timer = None

            self.last_cookie_value = ""
            self.last_token_value = ""

            self.last_env_signature = None
            self.dirty = False

            self.rows_by_code = {
                code: 0
                for code in PROFILES
            }

            self.state_by_code = {
                code: "READY"
                for code in PROFILES
            }

            self.failed_count = 0
            self.last_output = None

        # ====================================================
        # LAYOUT
        # ====================================================

        def compose(self) -> ComposeResult:
            yield Header(
                show_clock=True
            )

            with Vertical(id="workspace"):

                yield Static(
                    "OIC  /  XMLA COLLECTOR"
                    "                                  "
                    "INSURANCE INTELLIGENCE WORKSPACE",
                    id="brand",
                )

                # AUTH STATUS

                with Horizontal(id="auth_row"):

                    with Vertical(
                        classes="auth_panel"
                    ):
                        yield Static(
                            "BEARER JWT  /  LIVE INSPECTOR",
                            classes="auth_heading",
                        )

                        yield Static(
                            "Loading...",
                            id="jwt_status",
                            classes="auth_content",
                        )

                    with Vertical(
                        classes="auth_panel"
                    ):
                        yield Static(
                            "COOKIE SESSION  /  LIVE VALIDATION",
                            classes="auth_heading",
                        )

                        yield Static(
                            "Loading...",
                            id="cookie_status",
                            classes="auth_content",
                        )

                # CREDENTIAL EDITOR

                with Horizontal(id="credentials"):

                    with Vertical(
                        classes="credential_group"
                    ):
                        yield Label(
                            "OIC_BEARER_TOKEN",
                            classes="input_label",
                        )

                        yield Input(
                            placeholder="Paste JWT here...",
                            password=True,
                            id="token",
                        )

                        yield Static(
                            "JWT status updates immediately.",
                            classes="input_hint",
                        )

                    with Vertical(
                        classes="credential_group"
                    ):
                        yield Label(
                            "OIC_COOKIE",
                            classes="input_label",
                        )

                        yield Input(
                            placeholder=(
                                "Paste complete Cookie header..."
                            ),
                            password=True,
                            id="cookie",
                        )

                        yield Static(
                            "Cookie changes trigger GenToken validation.",
                            classes="input_hint",
                        )

                # ACTIONS

                with Horizontal(id="toolbar"):

                    yield Button(
                        "RELOAD .ENV",
                        id="reload",
                    )

                    yield Button(
                        "SAVE .ENV",
                        id="save",
                    )

                    yield Button(
                        "CHECK SESSION",
                        id="check",
                    )

                    yield Button(
                        "TEST API",
                        id="test",
                    )

                    yield Button("OPEN QUERY DESIGNER", id="open_designer")
                    yield Button("REFRESH CATALOGS", id="refresh_all_catalogs")
                    yield Button("EXPORT VERIFIED CATALOG", id="export_verified_catalog")
                    yield Button("PREPARE", id="prepare_all_dimensions")
                    yield Button(
                        "START COLLECTION",
                        id="start",
                    )

                    yield Button("STOP", id="stop")
                    yield Button("FORCE STOP", id="force_stop")

                    yield Button(
                        "OPEN OUTPUT",
                        id="open",
                    )
                    yield Button("RESUME", id="resume_run")

                # OUTPUT + CATEGORIES

                with Horizontal(id="config_row"):

                    with Vertical(
                        id="output_group"
                    ):
                        yield Label(
                            "OUTPUT DIRECTORY",
                            classes="input_label",
                        )

                        yield Input(
                            value="output",
                            id="output",
                        )

                    with Vertical(
                        id="category_group"
                    ):
                        yield Label(
                            "INSURANCE CATEGORIES",
                            classes="input_label",
                        )
                        yield Select(
                            [(code + " — " + p["name"], code) for code, p in PROFILES.items()],
                            value="MV", allow_blank=False, id="cube_select",
                        )

                        with Horizontal(
                            id="categories"
                        ):
                            for code in PROFILES:
                                yield Checkbox(
                                    code,
                                    value=(
                                        code in args.codes
                                    ),
                                    id=f"code_{code}",
                                )

                # SUMMARY

                yield Static(
                    "AVAILABLE  5"
                    "       SUCCESS  0"
                    "       FAILED  0"
                    "       RECORDS  0",
                    id="summary",
                )

                # CATALOG

                with Vertical(id="catalog_panel"):

                    yield Static(
                        "INSURANCE DATA CATALOG",
                        id="catalog_title",
                    )

                    yield DataTable(
                        id="table",
                        cursor_type="row",
                        zebra_stripes=True,
                    )

                # Field editing is delegated to a separate Textual window.
                with Vertical(id="applied_panel"):
                    yield Static("APPLIED QUERIES — 0 QUERIES", id="applied_title")
                    yield OptionList(id="applied_query_list", compact=True, markup=False)
                    with Horizontal(id="applied_actions"):
                        yield Button("REMOVE", id="remove_applied")
                        yield Button("MOVE UP", id="move_applied_up")
                        yield Button("MOVE DOWN", id="move_applied_down")
                    yield Static("No applied queries. Open Query Designer.",
                                 id="selected_query_summary")
                yield Static("FIELD CATALOG / Local caches", id="catalog_status")

                # PROGRESS

                with Vertical(id="progress_panel"):

                    yield Static(
                        "READY / Waiting for user action",
                        id="status",
                    )

                    yield ProgressBar(
                        total=5,
                        show_eta=False,
                        show_percentage=True,
                        id="progress",
                    )

                # ACTIVITY LOG

                with Vertical(id="activity_panel"):

                    yield Static(
                        "ACTIVITY LOG",
                        id="activity_title",
                    )

                    yield RichLog(
                        id="activity",
                        markup=True,
                        highlight=False,
                        wrap=True,
                        auto_scroll=True,
                    )

            yield Footer()

        # ====================================================
        # INITIALIZATION
        # ====================================================

        def on_mount(self) -> None:
            table = self.query_one(
                "#table",
                DataTable,
            )

            columns = [
                ("CODE", "code", 9),
                ("INSURANCE CLASS", "type", 28),
                ("CUBE", "cube", 26),
                ("ROWS", "rows", 13),
                ("RESPONSE", "size", 15),
                ("STATUS", "state", 18),
            ]

            for title, key, width in columns:
                table.add_column(
                    title,
                    key=key,
                    width=width,
                )

            for code, profile in PROFILES.items():
                table.add_row(
                    code,
                    profile["name"],
                    profile["cube"],
                    "—",
                    "—",
                    Text(
                        "READY",
                        style="#9aafbd",
                    ),
                    key=code,
                )

            self.query_one("#stop", Button).disabled = True
            self.query_one("#force_stop", Button).disabled = True
            self.select_category("MV")
            self.set_interval(1.0, self.poll_selected_run)
            self.set_interval(1.0, self.poll_applied_selection)
            from oic_all_dim_prepare import load_all_catalogs
            self.catalog_snapshot = load_all_catalogs()
            self.field_catalog = self.catalog_snapshot["MV/FY"][1]
            self.update_catalog_summary()

            # Load local configuration without making
            # any network requests.

            self.reload_from_env(
                auto_check=False
            )
            # Resolve the configured output location before restoring the
            # Applied Query List or reconnecting to an active batch.
            self.poll_applied_selection()
            self.reconnect_active_batch()

            self.info(
                "Terminal UI initialized.",
                "success",
            )

            self.info(
                "JWT expiry is inspected locally. "
                "Cookie requires GenToken verification."
            )

            # Live local JWT countdown.
            self.set_interval(
                1.0,
                self.refresh_auth_cards,
            )

            # Periodically detect external .env changes.
            self.set_interval(
                3.0,
                self.watch_env_file,
            )

            # Revalidate Session while the application
            # remains open. Does not run before UI startup.
            self.set_interval(
                SESSION_RECHECK_SECONDS,
                self.periodic_session_check,
            )

            # Initial network validation starts AFTER
            # the UI has mounted and become usable.
            self.set_timer(
                0.5,
                self.action_check_session,
            )

        # ====================================================
        # LOGGING
        # ====================================================

        def info(
            self,
            message: str,
            level: str = "info",
        ) -> None:
            colors = {
                "info": "#b1c0c9",
                "success": "#79c7aa",
                "warning": "#dfbf86",
                "error": "#ed9898",
            }

            color = colors.get(
                level,
                "#b1c0c9",
            )

            timestamp = datetime.now().strftime(
                "%H:%M:%S"
            )

            self.query_one(
                "#activity",
                RichLog,
            ).write(
                f"[dim]{timestamp}[/dim]  "
                f"[color={color}]"
                f"{rich_escape(str(message))}"
                f"[/color]"
            )

        # ====================================================
        # INPUT VALUES
        # ====================================================

        def token_value(self) -> str:
            return normalize_token(
                self.query_one(
                    "#token",
                    Input,
                ).value
            )

        def cookie_value(self) -> str:
            return normalize_cookie(
                self.query_one(
                    "#cookie",
                    Input,
                ).value
            )

        def output_value(self) -> str:
            return (
                self.query_one(
                    "#output",
                    Input,
                ).value.strip()
                or "output"
            )

        # ====================================================
        # LIVE JWT / COOKIE STATUS
        # ====================================================

        def refresh_auth_cards(self) -> None:
            if not self.is_mounted:
                return
            try:
                self.query_one("#jwt_status", Static)
                self.query_one("#cookie_status", Static)
            except NoMatches:
                return

            details = jwt_info(
                self.token_value()
            )

            jwt_status = details["status"]

            jwt_color = {
                "VALID": "#79c7aa",
                "EXPIRED": "#ed9898",
                "INVALID": "#ed9898",
                "MISSING": "#dfbf86",
            }.get(
                jwt_status,
                "#aebdc6",
            )

            jwt_content = Text()

            jwt_content.append(
                "STATUS       ",
                style="#93a7b4",
            )

            jwt_content.append(
                jwt_status,
                style=f"bold {jwt_color}",
            )

            jwt_content.append(
                "\n"
                f"Issued At    {details['issued']}"
                "\n"
                f"Expires At   {details['expires']}"
                "\n"
                f"Lifetime     {details['lifetime']}"
                "\n"
                f"Time Left    {details['remaining']}"
                "\n"
                "Source       .env / editor"
            )

            self.query_one(
                "#jwt_status",
                Static,
            ).update(jwt_content)

            cookie_colors = {
                "VALID": "#79c7aa",
                "CHECKING": "#dfbf86",
                "SESSION_EXPIRED": "#ed9898",
                "REJECTED": "#ed9898",
                "ERROR": "#ed9898",
                "NOT CHECKED": "#aebdc6",
                "MISSING": "#dfbf86",
            }

            cookie_color = cookie_colors.get(
                self.cookie_state,
                "#aebdc6",
            )

            cookie_content = Text()

            cookie_content.append(
                "STATUS       ",
                style="#93a7b4",
            )

            cookie_content.append(
                self.cookie_state,
                style=f"bold {cookie_color}",
            )

            cookie_content.append(
                "\n"
                "Configured   "
                + (
                    "YES"
                    if self.cookie_value()
                    else "NO"
                )
                + "\n"
                f"Last Check   {self.cookie_checked_at}"
                "\n"
                "Validation   GenToken endpoint"
                "\n"
                "Expiry       Server controlled"
            )

            if self.cookie_message:
                cookie_content.append(
                    "\n"
                    f"Message      {self.cookie_message[:80]}"
                )

            self.query_one(
                "#cookie_status",
                Static,
            ).update(cookie_content)

        # ====================================================
        # CONFIGURATION LOAD
        # ====================================================

        def reload_from_env(
            self,
            auto_check: bool = True,
        ) -> None:
            if self.busy:
                return

            self.cancel_auth_timer()

            self.auth_sequence += 1

            self.config = read_env()

            token = normalize_token(
                self.config.get(
                    "OIC_BEARER_TOKEN",
                    "",
                )
            )

            cookie = normalize_cookie(
                self.config.get(
                    "OIC_COOKIE",
                    "",
                )
            )

            output = self.config.get(
                "OIC_OUTPUT_DIR",
                "output",
            ) or "output"

            # Update tracking values before assigning
            # Input.value to avoid redundant validation.

            self.last_token_value = token
            self.last_cookie_value = cookie

            self.query_one(
                "#token",
                Input,
            ).value = token

            self.query_one(
                "#cookie",
                Input,
            ).value = cookie

            self.query_one(
                "#output",
                Input,
            ).value = output

            self.active_token = ""

            self.cookie_state = (
                "NOT CHECKED"
                if cookie
                else "MISSING"
            )

            self.cookie_message = ""
            self.cookie_checked_at = "—"

            self.dirty = False

            self.last_env_signature = (
                env_signature()
            )

            self.refresh_auth_cards()

            self.info(
                f"Loaded {ENV_PATH.name} / "
                f"JWT: {jwt_info(token)['status']}"
            )

            if auto_check and cookie:
                self.schedule_auth_check(
                    delay=0.3
                )

        def action_reload_config(self) -> None:
            self.reload_from_env(
                auto_check=True
            )

        # ====================================================
        # EXTERNAL .ENV WATCHER
        # ====================================================

        def watch_env_file(self) -> None:
            if self.busy:
                return

            current_signature = (
                env_signature()
            )

            if (
                current_signature
                == self.last_env_signature
            ):
                return

            if self.dirty:
                # Preserve unsaved editor changes.
                return

            self.info(
                "External .env change detected. Reloading.",
                "warning",
            )

            self.reload_from_env(
                auto_check=True
            )

        # ====================================================
        # SAVE .ENV
        # ====================================================

        def save_credentials(self) -> None:
            if self.busy:
                return

            try:
                token = self.token_value()
                cookie = self.cookie_value()

                update_env({
                    "OIC_BEARER_TOKEN": token,
                    "OIC_COOKIE": cookie,
                    "OIC_OUTPUT_DIR": self.output_value(),
                })

                self.last_env_signature = (
                    env_signature()
                )

                self.dirty = False

                self.info(
                    "Credentials saved to .env",
                    "success",
                )

                self.refresh_auth_cards()

            except OSError as exc:
                self.info(
                    f"Cannot save .env: {exc}",
                    "error",
                )

        # ====================================================
        # LIVE INPUT CHANGE EVENTS
        # ====================================================

        def on_input_changed(
            self,
            event: Input.Changed,
        ) -> None:
            input_id = event.input.id

            if input_id == "token":
                current = normalize_token(
                    event.value
                )

                if current != self.last_token_value:
                    self.last_token_value = current
                    self.dirty = True

                    self.refresh_auth_cards()

            elif input_id == "cookie":
                current = normalize_cookie(
                    event.value
                )

                if current == self.last_cookie_value:
                    return

                self.last_cookie_value = current
                self.dirty = True

                # Invalidate all in-flight validations
                # immediately after Cookie changes.

                self.auth_sequence += 1

                self.cancel_auth_timer()

                if not current:
                    self.cookie_state = "MISSING"
                    self.cookie_message = ""
                    self.cookie_checked_at = "—"

                else:
                    self.cookie_state = "NOT CHECKED"
                    self.cookie_message = "Waiting for validation"

                    # Debounce to avoid one request
                    # per character while typing/pasting.
                    self.schedule_auth_check(
                        delay=COOKIE_DEBOUNCE_SECONDS
                    )

                self.refresh_auth_cards()

            elif input_id == "output":
                self.dirty = True

        def on_input_submitted(
            self,
            event: Input.Submitted,
        ) -> None:
            if event.input.id == "cookie":
                self.action_check_session()

            elif event.input.id == "token":
                self.refresh_auth_cards()

        # ====================================================
        # AUTHENTICATION TIMER
        # ====================================================

        def cancel_auth_timer(self) -> None:
            if self.auth_timer is not None:
                try:
                    self.auth_timer.stop()
                except Exception:
                    pass

                self.auth_timer = None

        def schedule_auth_check(
            self,
            delay: float,
        ) -> None:
            self.cancel_auth_timer()

            self.auth_timer = self.set_timer(
                delay,
                self.action_check_session,
            )

        def periodic_session_check(self) -> None:
            if self.busy or self.catalog_running:
                return

            if not self.cookie_value():
                return

            self.action_check_session()

        # ====================================================
        # ASYNCHRONOUS SESSION VALIDATION
        # ====================================================

        def action_check_session(self) -> None:
            """
            Runs only after Terminal UI startup.

            Never closes the application if the session
            is invalid or expired.
            """

            self.cancel_auth_timer()

            if self.busy or self.catalog_running:
                return

            cookie = self.cookie_value()

            if not cookie:
                self.auth_sequence += 1

                self.cookie_state = "MISSING"
                self.cookie_message = "Enter OIC_COOKIE"

                self.refresh_auth_cards()
                return

            self.auth_sequence += 1

            sequence = self.auth_sequence

            self.cookie_state = "CHECKING"
            self.cookie_message = "Contacting GenToken..."

            self.refresh_auth_cards()

            self.info(
                "Checking OIC Session via GenToken..."
            )

            self.validate_session_worker(
                sequence,
                cookie,
            )

        @work(
            thread=True,
            group="auth-validation",
        )
        def validate_session_worker(
            self,
            sequence: int,
            cookie: str,
        ) -> None:
            """
            A separate requests.Session is used for
            each validation.

            Results are applied only when the Cookie
            still matches the value being validated.
            """

            client = OICClient(
                cookie=cookie,
                token="",
                timeout=45,
                retries=2,
            )

            try:
                token = client.request_token(
                    "AH"
                )

                updated_cookie = (
                    client.cookie_header()
                )

                self.call_from_thread(
                    self.session_validation_finished,
                    sequence,
                    cookie,
                    token,
                    updated_cookie,
                    None,
                )

            except Exception as exc:
                self.call_from_thread(
                    self.session_validation_finished,
                    sequence,
                    cookie,
                    None,
                    None,
                    exc,
                )

            finally:
                client.close()

        def session_validation_finished(
            self,
            sequence: int,
            original_cookie: str,
            token: str | None,
            updated_cookie: str | None,
            error,
        ) -> None:
            # Discard stale results if the user edited
            # the Cookie while the request was running.

            if sequence != self.auth_sequence:
                return

            if self.busy:
                return

            if (
                self.cookie_value()
                != original_cookie
            ):
                return

            self.cookie_checked_at = (
                datetime.now().strftime(
                    "%H:%M:%S"
                )
            )

            if error is not None:
                if isinstance(
                    error,
                    OICSessionExpired,
                ):
                    self.cookie_state = (
                        "SESSION_EXPIRED"
                    )

                elif isinstance(
                    error,
                    OICAuthenticationError,
                ):
                    self.cookie_state = "REJECTED"

                else:
                    self.cookie_state = "ERROR"

                message = str(error)

                self.cookie_message = message[:80]

                self.refresh_auth_cards()

                self.info(
                    "Session validation failed: "
                    + message,
                    "error",
                )

                self.query_one(
                    "#status",
                    Static,
                ).update(
                    "AUTHENTICATION REQUIRED / "
                    "Edit Cookie and retry"
                )

                # Intentionally do not exit.
                # All editor controls stay enabled.
                return

            # Session is accepted by GenToken.

            self.cookie_state = "VALID"
            self.cookie_message = "GenToken successful"

            self.active_token = token

            # Preserve all original Cookie values and
            # any updates returned by Set-Cookie.

            final_cookie = (
                updated_cookie
                or original_cookie
            )

            self.last_token_value = token
            self.last_cookie_value = final_cookie

            self.query_one(
                "#token",
                Input,
            ).value = token

            self.query_one(
                "#cookie",
                Input,
            ).value = final_cookie

            self.refresh_auth_cards()

            try:
                update_env({
                    "OIC_BEARER_TOKEN": token,
                    "OIC_COOKIE": final_cookie,
                    "OIC_OUTPUT_DIR": self.output_value(),
                })

                self.last_env_signature = (
                    env_signature()
                )

                self.dirty = False

                self.info(
                    "Session VALID / "
                    "fresh JWT and Cookie saved to .env",
                    "success",
                )

            except OSError as exc:
                self.info(
                    "Session valid, but .env could not "
                    f"be updated: {exc}",
                    "warning",
                )

            self.query_one("#status", Static).update(
                "READY / Session VALID / GenToken successful"
            )

        # ====================================================
        # BUTTON EVENTS
        # ====================================================

        def on_button_pressed(
            self,
            event: Button.Pressed,
        ) -> None:
            button_id = event.button.id

            if button_id == "reload":
                self.action_reload_config()

            elif button_id == "save":
                self.save_credentials()

            elif button_id == "check":
                self.action_check_session()

            elif button_id == "test":
                self.action_test_api()

            elif button_id == "start":
                self.action_start_collection()

            elif button_id == "stop":
                self.action_stop_collection()
            elif button_id == "force_stop":
                self.force_stop_collection()
            elif button_id == "remove_applied":
                self.manage_applied("remove")
            elif button_id == "move_applied_up":
                self.manage_applied("up")
            elif button_id == "move_applied_down":
                self.manage_applied("down")

            elif button_id == "open":
                self.action_open_output()

            elif button_id == "resume_run":
                self.resume_selected_query()
            elif button_id == "open_designer":
                self.open_query_designer()
            elif button_id == "prepare_all_dimensions":
                self.action_prepare_all_dimensions()
            elif button_id == "refresh_all_catalogs":
                if self.busy or self.active_run is not None or self.catalog_running:
                    return
                if self.cookie_state != "VALID":
                    self.info("Check Session before refreshing catalogs.", "warning")
                    return
                self.catalog_running = True
                self.query_one("#refresh_all_catalogs", Button).disabled = True
                self.query_one("#prepare_all_dimensions", Button).disabled = True
                self.query_one("#export_verified_catalog", Button).disabled = True
                self.discover_all_catalogs(force=True)
            elif button_id == "export_verified_catalog":
                button = self.query_one("#export_verified_catalog", Button)
                if button.disabled:
                    return
                if self.busy or self.catalog_running or self.active_run is not None:
                    self.info("Another operation is running; export is not started.", "warning")
                    return
                button.disabled = True
                button.label = "EXPORTING..."
                self.query_one("#refresh_all_catalogs", Button).disabled = True
                self.query_one("#prepare_all_dimensions", Button).disabled = True
                self.info("Exporting saved field verification reports...")
                self.export_catalog_worker()

        # ====================================================
        # SELECTED INSURANCE CLASSES
        # ====================================================

        def selected_codes(self) -> list[str]:
            return [self.selected_code]

        # ====================================================
        # COLLECTION ACTIONS
        # ====================================================

        def action_test_api(self) -> None:
            self.start_job(
                mode="test"
            )

        def action_start_collection(self) -> None:
            if self.busy or self.catalog_running or self.query_one("#export_verified_catalog", Button).disabled:
                return
            self.start_selected_query()

        def action_stop_collection(self) -> None:
            if self.active_run is not None:
                (self.active_run / "stop.flag").touch()
                self.info("STOPPING / No new queries or tasks; workers finish active requests.", "warning")
                self.query_one("#status", Static).update("STOPPING / Waiting for workers")
                return
            if not self.busy:
                return

            self.stop_event.set()

            self.info(
                "Stop requested. "
                "Waiting for the current request...",
                "warning",
            )

            self.query_one(
                "#status",
                Static,
            ).update(
                "STOP REQUESTED / "
                "Finishing current operation"
            )

        def force_stop_collection(self) -> None:
            if self.active_run is None:
                return
            if time.monotonic() > self.force_confirm_until:
                self.force_confirm_until = time.monotonic() + 8
                self.info("FORCE STOP will terminate active workers. Press again within 8 seconds to confirm.", "warning")
                return
            (self.active_run / "stop.flag").touch()
            (self.active_run / "force_stop.flag").touch()
            self.force_confirm_until = 0
            self.info("FORCE STOP requested; completed chunks remain saved.", "warning")

        # ====================================================
        # OPEN OUTPUT DIRECTORY
        # ====================================================

        def action_open_output(self) -> None:
            output_dir = resolve_output_path(
                self.output_value()
            ) / "Excel"

            try:
                output_dir.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                if sys.platform == "win32":
                    os.startfile(
                        str(output_dir)
                    )

                else:
                    import subprocess

                    subprocess.Popen(
                        (
                            ["open", str(output_dir)]
                            if sys.platform == "darwin"
                            else ["xdg-open", str(output_dir)]
                        )
                    )

                self.info(
                    f"Opened output: {output_dir}",
                    "success",
                )

            except Exception as exc:
                self.info(
                    f"Cannot open output: {exc}",
                    "error",
                )

        # ====================================================
        # BUSY STATE
        # ====================================================

        def set_busy(
            self,
            busy: bool,
        ) -> None:
            self.busy = busy

            for button_id in (
                "reload",
                "save",
                "check",
                "test",
                "start",
                "open",
                "prepare_all_dimensions",
                "refresh_all_catalogs",
                "export_verified_catalog",
                "open_designer",
                "resume_run",
            ):
                self.query_one(
                    f"#{button_id}",
                    Button,
                ).disabled = busy

            for input_id in (
                "token",
                "cookie",
                "output",
            ):
                self.query_one(
                    f"#{input_id}",
                    Input,
                ).disabled = busy

            for code in PROFILES:
                self.query_one(
                    f"#code_{code}",
                    Checkbox,
                ).disabled = busy

            self.query_one(
                "#stop",
                Button,
            ).disabled = not busy

        # ====================================================
        # DATA TABLE
        # ====================================================

        def update_row(
            self,
            code: str,
            state: str,
            count: int = 0,
            size: int = 0,
        ) -> None:
            table = self.query_one(
                "#table",
                DataTable,
            )

            colors = {
                "READY": "#9aafbd",
                "RUNNING": "#dfbf86",
                "SUCCESS": "#79c7aa",
                "FAILED": "#ed9898",
                "STOPPED": "#dfbf86",
            }

            self.state_by_code[code] = state

            if state == "SUCCESS":
                self.rows_by_code[code] = count

            elif state == "FAILED":
                self.rows_by_code[code] = 0

            table.update_cell(
                code,
                "rows",
                (
                    f"{count:,}"
                    if count
                    else "—"
                ),
            )

            table.update_cell(
                code,
                "size",
                (
                    f"{size:,} B"
                    if size
                    else "—"
                ),
            )

            table.update_cell(
                code,
                "state",
                Text(
                    state,
                    style=colors.get(
                        state,
                        "#a9bbc7",
                    ),
                ),
            )

            self.update_summary()

        # ====================================================
        # SUMMARY
        # ====================================================

        def update_summary(self) -> None:
            total_rows = sum(
                self.rows_by_code.values()
            )

            success_count = sum(
                1
                for state in self.state_by_code.values()
                if state == "SUCCESS"
            )

            failed_count = sum(
                1
                for state in self.state_by_code.values()
                if state == "FAILED"
            )

            running_count = sum(
                1
                for state in self.state_by_code.values()
                if state == "RUNNING"
            )

            summary = (
                f"AVAILABLE  {len(PROFILES)}"
                "       "
                f"SUCCESS  {success_count}"
                "       "
                f"RUNNING  {running_count}"
                "       "
                f"FAILED  {failed_count}"
                "       "
                f"RECORDS  {total_rows:,}"
            )

            self.query_one(
                "#summary",
                Static,
            ).update(summary)

        # ====================================================
        # PROGRESS
        # ====================================================

        def set_progress(
            self,
            current: int,
            total: int,
        ) -> None:
            self.query_one(
                "#progress",
                ProgressBar,
            ).update(
                total=max(total, 1),
                progress=current,
            )

            self.query_one(
                "#status",
                Static,
            ).update(
                "COLLECTING / "
                f"{current} of {total} categories"
            )

        # ====================================================
        # START COLLECTION
        # ====================================================

        def start_job(
            self,
            mode: str,
        ) -> None:
            if self.busy or self.catalog_running or self.active_run is not None or self.query_one("#export_verified_catalog", Button).disabled:
                return

            codes = self.selected_codes()

            if not codes:
                self.info(
                    "Select at least one category.",
                    "warning",
                )
                return

            cookie = self.cookie_value()
            token = self.token_value()

            if not cookie:
                self.cookie_state = "MISSING"
                self.cookie_message = (
                    "Enter OIC_COOKIE"
                )

                self.refresh_auth_cards()

                self.info(
                    "OIC_COOKIE is missing. "
                    "Enter a valid session Cookie.",
                    "error",
                )

                return

            if mode == "test":
                codes = codes[:1]

            output_dir = resolve_output_path(
                self.output_value()
            )

            # Cancel pending Cookie validation.
            # Collection will request GenToken itself.

            self.cancel_auth_timer()
            self.auth_sequence += 1

            self.stop_event.clear()

            self.rows_by_code = {
                code: 0
                for code in PROFILES
            }

            self.state_by_code = {
                code: "READY"
                for code in PROFILES
            }

            self.failed_count = 0

            for code in PROFILES:
                self.update_row(
                    code,
                    "READY",
                )

            self.set_busy(True)

            self.query_one(
                "#progress",
                ProgressBar,
            ).update(
                total=max(len(codes), 1),
                progress=0,
            )

            self.query_one(
                "#status",
                Static,
            ).update(
                f"{mode.upper()} / "
                f"{', '.join(codes)}"
            )

            self.info(
                f"Starting {mode.upper()} / "
                f"{', '.join(codes)}"
            )

            self.collection_worker(
                mode,
                codes,
                cookie,
                token,
                output_dir,
            )

        # ====================================================
        # AUTH REFRESHED DURING COLLECTION
        # ====================================================

        def collection_auth_updated(
            self,
            token: str,
            cookie: str,
        ) -> None:
            self.cookie_state = "VALID"
            self.cookie_message = (
                "GenToken successful"
            )

            self.cookie_checked_at = (
                datetime.now().strftime(
                    "%H:%M:%S"
                )
            )

            self.active_token = token

            self.last_token_value = token
            self.last_cookie_value = cookie

            self.query_one(
                "#token",
                Input,
            ).value = token

            self.query_one(
                "#cookie",
                Input,
            ).value = cookie

            self.refresh_auth_cards()

            try:
                update_env({
                    "OIC_BEARER_TOKEN": token,
                    "OIC_COOKIE": cookie,
                    "OIC_OUTPUT_DIR": self.output_value(),
                })

                self.last_env_signature = (
                    env_signature()
                )

                self.dirty = False

                self.info(
                    "Authentication refreshed / "
                    "credentials saved to .env",
                    "success",
                )

            except OSError as exc:
                self.info(
                    f"Could not save refreshed JWT: {exc}",
                    "warning",
                )

        # ====================================================
        # COLLECTION FINISHED
        # ====================================================

        def collection_finished(
            self,
            mode: str,
            result=None,
            error=None,
            stopped: bool = False,
        ) -> None:
            self.set_busy(False)

            if stopped:
                for code, state in list(
                    self.state_by_code.items()
                ):
                    if state == "RUNNING":
                        self.update_row(
                            code,
                            "STOPPED",
                        )

                self.query_one(
                    "#status",
                    Static,
                ).update(
                    "STOPPED / Operation interrupted"
                )

                self.info(
                    "Collection stopped. "
                    "No incomplete Excel file retained.",
                    "warning",
                )

                return

            if error is not None:
                message = str(error)

                if isinstance(
                    error,
                    OICSessionExpired,
                ):
                    self.cookie_state = (
                        "SESSION_EXPIRED"
                    )

                    self.cookie_message = message

                elif isinstance(
                    error,
                    OICAuthenticationError,
                ):
                    self.cookie_state = "REJECTED"
                    self.cookie_message = message

                for code, state in list(
                    self.state_by_code.items()
                ):
                    if state == "RUNNING":
                        self.update_row(
                            code,
                            "FAILED",
                        )

                self.refresh_auth_cards()

                self.query_one(
                    "#status",
                    Static,
                ).update(
                    "FAILED / " + message[:120]
                )

                self.info(
                    f"{mode.upper()} failed: {message}",
                    "error",
                )

                self.info(
                    "Terminal UI remains active. "
                    "Edit Cookie and retry.",
                    "warning",
                )

                # Do not close or exit the application.
                return

            self.last_output = (
                result.get("output")
                if result
                else None
            )

            if mode == "test":
                message = (
                    "TEST PASSED / "
                    f"{result['rows']:,} rows / "
                    "No Excel created"
                )

            else:
                message = (
                    "COMPLETE / "
                    f"{result['rows']:,} rows / "
                    f"{len(result['errors'])} failed"
                )

            self.query_one(
                "#status",
                Static,
            ).update(message)

            self.info(
                message,
                (
                    "success"
                    if not result["errors"]
                    else "warning"
                ),
            )

            if result and result.get("output"):
                output_path = result["output"]

                file_size = (
                    output_path.stat().st_size
                    / 1024
                    / 1024
                )

                self.info(
                    "Excel created: "
                    f"{output_path.name} "
                    f"({file_size:.2f} MB)",
                    "success",
                )

                self.info(
                    f"Output: {output_path}"
                )

            self.update_summary()

        # ====================================================
        # COLLECTION WORKER
        # ====================================================

        @work(
            thread=True,
            group="collection",
            exclusive=True,
        )
        def collection_worker(
            self,
            mode: str,
            codes: list[str],
            cookie: str,
            token: str,
            output_dir: Path,
        ) -> None:

            def log_callback(
                message: str,
                level: str = "info",
            ) -> None:
                self.call_from_thread(
                    self.info,
                    message,
                    level,
                )

            def auth_callback(
                new_token: str,
                new_cookie: str,
            ) -> None:
                self.call_from_thread(
                    self.collection_auth_updated,
                    new_token,
                    new_cookie,
                )

            def result_callback(
                code: str,
                state: str,
                count: int,
                size: int,
            ) -> None:
                self.call_from_thread(
                    self.update_row,
                    code,
                    state,
                    count,
                    size,
                )

            def progress_callback(
                current: int,
                total: int,
            ) -> None:
                self.call_from_thread(
                    self.set_progress,
                    current,
                    total,
                )

            try:
                result = collect(
                    codes=codes,
                    cookie=cookie,
                    token=token,
                    output_dir=output_dir,
                    timeout=args.timeout,
                    retries=args.retries,
                    test_only=(
                        mode == "test"
                    ),
                    stop_event=self.stop_event,
                    on_log=log_callback,
                    on_auth=auth_callback,
                    on_result=result_callback,
                    on_progress=progress_callback,
                )

                self.call_from_thread(
                    self.collection_finished,
                    mode,
                    result,
                )

            except OICStopped:
                self.call_from_thread(
                    self.collection_finished,
                    mode,
                    None,
                    None,
                    True,
                )

            except Exception as exc:
                self.call_from_thread(
                    self.collection_finished,
                    mode,
                    None,
                    exc,
                )

        # ====================================================
        # MAIN WINDOW: CONFIG BRIDGE AND CATALOG DISCOVERY
        # ====================================================

        def selection_file(self) -> Path:
            return resolve_output_path(self.output_value()) / "applied_queries.json"

        def select_category(self, code: str) -> None:
            if code in QUERY_PROFILES:
                self.selected_code = code

        def on_select_changed(self, event: Select.Changed) -> None:
            if event.select.id == "cube_select" and event.value in QUERY_PROFILES:
                self.select_category(str(event.value))

        def update_query_summary(self) -> None:
            if not self.is_mounted:
                return
            from textual.widgets.option_list import Option
            listing = self.query_one("#applied_query_list", OptionList)
            previous = listing.highlighted_option
            focused = previous.id if previous else None
            listing.clear_options()
            for item in self.applied_queries:
                cfg = item["query"]
                label = (f"{item['query_id']}  {cfg['code']}/{cfg.get('cube_type','FY')}"
                         f"  {cfg['cube']}  X:{len(cfg['x'])} Y:{len(cfg['y'])}"
                         f" V:{len(cfg['measures'])}  W:{cfg['workers']}  "
                         f"{self.query_states.get(item['query_id'], 'PLANNED')}")
                listing.add_option(Option(label, id=item["query_id"]))
            if self.applied_queries:
                ids = [q["query_id"] for q in self.applied_queries]
                listing.highlighted = ids.index(focused) if focused in ids else 0
            self.query_one("#applied_title", Static).update(
                f"APPLIED QUERIES — {len(self.applied_queries)} QUERIES")
            self.query_one("#selected_query_summary", Static).update(
                "APPLY creates tasks; START executes the prepared batch. "
                "Duplicate checks use only this list.")

        def query_config(self) -> dict:
            if not self.applied_queries:
                raise ValueError("Apply at least one query in Query Designer")
            return validate_query(dict(self.applied_queries[0]["query"]))

        def _write_applied_list(self, entries: list[dict]) -> None:
            path = self.selection_file()
            with applied_list_write_lock(path):
                envelope, legacy = read_applied_queries(path)
                # This helper is only used for changes initiated from the App;
                # refuse to overwrite if the on-disk revision moved.
                if (self.selection_revision is not None
                        and path.exists()
                        and envelope["revision"] != self.selection_revision):
                    raise ValueError("Applied List changed in Designer; refresh and retry")
                envelope["queries"] = entries
                save_applied_queries(path, envelope, backup_legacy=legacy)
            self.applied_queries = entries
            self.selection_revision = envelope["revision"]
            self.update_query_summary()

        def manage_applied(self, action: str) -> None:
            if self.active_run is not None:
                self.info("Cannot edit applied list during collection.", "warning")
                return
            listing = self.query_one("#applied_query_list", OptionList)
            selected = listing.highlighted_option
            if selected is None:
                self.info("Select an Applied Query first.", "warning")
                return
            path = self.selection_file()
            try:
                with applied_list_write_lock(path):
                    envelope, legacy = read_applied_queries(path)
                    entries = envelope["queries"]
                    index = next((i for i, item in enumerate(entries)
                                  if item["query_id"] == selected.id), -1)
                    if index < 0:
                        raise ValueError("Applied Query changed; select it again")
                    if action == "remove":
                        entries.pop(index)
                    elif action in ("up", "down"):
                        other = index + (-1 if action == "up" else 1)
                        if not 0 <= other < len(entries):
                            return
                        entries[index], entries[other] = entries[other], entries[index]
                    else:
                        raise ValueError("Unknown Applied Query action")
                    save_applied_queries(path, envelope, backup_legacy=legacy)
                self.applied_queries = envelope["queries"]
                self.selection_revision = envelope["revision"]
                self.applied_config = (self.applied_queries[-1]["query"]
                                       if self.applied_queries else None)
                self.update_query_summary()
                self.info(f"Applied Query {selected.id}: {action}", "success")
            except (OSError, ValueError, TimeoutError) as exc:
                self.info(f"Cannot edit Applied List: {exc}", "warning")

        def poll_applied_selection(self) -> None:
            path = self.selection_file()
            if not path.is_file():
                return
            try:
                envelope, legacy = read_applied_queries(path)
                revision = envelope["revision"]
                if revision == self.selection_revision:
                    return
                entries = envelope["queries"]
            except (ValueError, TypeError, KeyError, OSError) as exc:
                message = str(exc)
                if getattr(self, "_applied_last_error", None) != message:
                    self._applied_last_error = message
                    self.info(f"Applied List cannot be read: {message}", "warning")
                return
            self._applied_last_error = None
            self.applied_queries = entries
            self.applied_config = entries[-1]["query"] if entries else None
            self.selection_revision = revision
            self.update_query_summary()
            if legacy:
                self.info(
                    f"Loaded {len(entries)} legacy Applied Queries. "
                    "Original file will be backed up and upgraded on next edit/APPLY.",
                    "warning",
                )
            else:
                self.info(f"Applied Query List updated: {len(entries)} queries.", "success")


        # oic_collector.py
        # Replace OICApp.open_query_designer() inside launch_terminal_ui().

        def open_query_designer(self) -> None:
            if self.busy or self.catalog_running:
                return
            if sys.platform != "win32":
                self.info(
                    "Dedicated Query Designer requires Windows.",
                    "error",
                )
                return

            from oic_parallel_collector import launch_terminal_window

            selection = self.selection_file().resolve()
            lock_path = selection.with_suffix(".lock")

            # Prevent opening another instance when the existing
            # Query Designer process is still active.

            if lock_path.exists():
                try:
                    pid = int(
                        lock_path.read_text(
                            encoding="ascii"
                        ).strip()
                    )

                    import ctypes

                    process_handle = ctypes.windll.kernel32.OpenProcess(
                        0x1000,
                        False,
                        pid,
                    )

                    if process_handle:
                        ctypes.windll.kernel32.CloseHandle(
                            process_handle
                        )

                        self.info(
                            "Query Designer is already open.",
                            "warning",
                        )
                        return

                except (OSError, ValueError):
                    pass

                try:
                    lock_path.unlink(missing_ok=True)
                except OSError as exc:
                    self.info(
                        f"Cannot clear stale Designer lock: {exc}",
                        "error",
                    )
                    return

            try:
                selection.parent.mkdir(
                    parents=True,
                    exist_ok=True,
                )

                python_path = str(
                    Path(sys.executable).resolve()
                )

                script_path = str(
                    Path(__file__).resolve()
                )

                designer_command = [
                    python_path,
                    script_path,
                    "--designer",
                    "--selection-file",
                    str(selection),
                    "--category",
                    self.selected_code,
                ]

                mode = launch_terminal_window(
                    [designer_command],
                    title="OIC | Query Designer",
                    cwd=BASE_DIR,
                    tab_titles=["OIC | Query Designer"],
                    keep_open=False,
                    maximized=True,
                )

                self.info(
                    f"Query Designer opened via {mode}.",
                    "success",
                )

            except (OSError, RuntimeError, ValueError) as exc:
                self.info(
                    f"Query Designer launch failed: {exc}",
                    "error",
                )

        def action_prepare_all_dimensions(self) -> None:
            if (self.busy or self.prepare_running or self.catalog_running or
                    self.active_run is not None or self.query_one("#export_verified_catalog", Button).disabled):
                self.info("Another operation is running.", "warning")
                return
            if self.cookie_state != "VALID":
                self.info("Check Session before PREPARE.", "warning")
                return
            missing = [key for key, (profile, catalog, fields) in self.catalog_snapshot.items()
                       if profile["availability"] != "NOT_AVAILABLE" and
                       (profile["availability"] != "AVAILABLE" or catalog.get("source") == "PROFILE" or not fields)]
            if len(self.catalog_snapshot) != len(QUERY_PROFILES) * len(CUBE_TYPES) or missing:
                self.info("Saved catalog incomplete; use REFRESH CATALOGS: " + ", ".join(missing), "warning")
                return
            cookie, token = self.cookie_value(), self.token_value()
            try:
                # The curl benchmark/extraction clients read the same existing
                # .env. Persist the current UI credentials before the handoff.
                update_env({"OIC_COOKIE": cookie, "OIC_BEARER_TOKEN": token,
                            "OIC_OUTPUT_DIR": self.output_value()})
                self.last_env_signature = env_signature()
                self.dirty = False
            except OSError:
                self.info("Cannot save credentials for PREPARE.", "error")
                return
            self.cancel_auth_timer()
            self.prepare_running = True
            self.stop_event.clear()
            self.set_busy(True)
            self.query_one("#progress", ProgressBar).update(total=1, progress=0)
            self.query_one("#status", Static).update("PREPARING / Reusing saved catalog")
            self.prepare_dimensions_worker(self.catalog_snapshot, cookie, token)

        @work(thread=True, group="prepare-all-dimensions", exclusive=True)
        def prepare_dimensions_worker(self, catalogs, cookie, token) -> None:
            client = None
            result = 1
            error = None
            failed = False
            try:
                from oic_all_dim import prepare_from_collector
                client = OICClient(cookie=cookie, token=token, stop_event=self.stop_event,
                    on_auth=lambda jwt, cookies: self.call_from_thread(
                        self.collection_auth_updated, jwt, cookies),
                    on_log=lambda message, level="info": self.call_from_thread(
                        self.info, message, level))
                result = prepare_from_collector(catalogs, client,
                    lambda event, **data: self.call_from_thread(self.prepare_progress, event, data),
                    self.stop_event)
            except Exception as exc:
                failed = True
                error = str(exc) if type(exc).__name__ == "CollectionError" else type(exc).__name__
            finally:
                try:
                    if client is not None:
                        client.close()
                finally:
                    self.call_from_thread(self.prepare_finished, result, error, failed)

        def prepare_progress(self, event, data) -> None:
            if event == "log":
                self.info(data["message"])
            elif event == "phase":
                self.query_one("#status", Static).update(
                    f"PREPARING / {data.get('cube', '')} / {data['phase']} / {data.get('detail', '')}")
            elif event == "cube":
                if "total" in data and data["total"] is not None:
                    self.query_one("#progress", ProgressBar).update(total=max(1,data["total"]), progress=data.get("done",0))
                elif "done" in data:
                    self.query_one("#progress", ProgressBar).update(progress=data["done"])
            elif event == "plan_job":
                self.query_one("#progress", ProgressBar).update(total=max(1,data["total"]),progress=data["index"]-1)
                self.query_one("#status", Static).update(
                    f"PREPARING / {data['cube']} / REQUEST PLAN {data['index']}/{data['total']}")
            elif event == "summary":
                summary=data["summary"]
                self.info(f"PREPARE SUMMARY | Members {summary.get('discovered_members',0):,} | Jobs {summary.get('jobs',0):,}")

        def prepare_finished(self, result, error, failed=False) -> None:
            self.prepare_running = False
            self.set_busy(False)
            status = "FAILED" if failed else "READY" if result == 0 else "STOPPED" if result == 130 else "INCOMPLETE"
            self.query_one("#status", Static).update("PREPARE " + status)
            self.info("PREPARE " + status + (f" / {error}" if error else ""),
                      "success" if status=="READY" else "error" if failed else "warning")
            if status=="READY":
                self.query_one("#progress", ProgressBar).update(total=1,progress=1)
            # Benchmarks may have renewed the shared .env. Adopt the persisted
            # credentials through the existing watcher (no catalog reload).
            self.watch_env_file()
            if failed and error and error.startswith(("AUTH_COOKIE_EXPIRED", "ACCESS_DENIED", "AUTH_TOKEN_EXPIRED")):
                self.cookie_state = "NOT CHECKED"
                self.cookie_message = "Check Session before retrying"
                self.refresh_auth_cards()

        def update_catalog_summary(self) -> None:
            totals = dict(total=0, verified=0, mapped=0, unverified=0,
                          unsupported=0, selectable=0)
            available = 0
            lines = []
            for code in QUERY_PROFILES:
                values = []
                for kind in CUBE_TYPES:
                    profile, catalog, _ = self.catalog_snapshot[f"{code}/{kind}"]
                    summary = catalog_verification_summary(catalog)
                    status = profile["availability"]
                    self.catalog_statuses[f"{code}/{kind}"] = status
                    available += status == "AVAILABLE"
                    if catalog.get("source") == "PROFILE":
                        values.append(f"{kind}: {status} / NO SAVED CATALOG")
                        continue
                    for key in totals:
                        totals[key] += summary[key]
                    values.append(f"{kind}: {summary['verified']}/{summary['total']}"
                                  f" ({status})")
                lines.append(f"{code}  " + "  |  ".join(values))
            coverage = 100 * totals["verified"] / totals["total"] if totals["total"] else 0
            self.query_one("#catalog_status", Static).update(
                "FIELD CATALOG | REPORTED VERIFICATION\n"
                f"FIELDS: {totals['total']:,} | VERIFIED: {totals['verified']:,} "
                f"({coverage:.2f}%) | MAPPED: {totals['mapped']:,} | "
                f"UNVERIFIED: {totals['unverified']:,} | UNSUPPORTED: {totals['unsupported']:,}\n"
                f"SELECTABLE: {totals['selectable']:,} | CUBES AVAILABLE: "
                f"{available}/{len(QUERY_PROFILES) * len(CUBE_TYPES)}\n"
                + "\n".join(lines))

        @work(thread=True, group="catalog-export", exclusive=True)
        def export_catalog_worker(self) -> None:
            result = None
            error = None
            try:
                result = export_verified_catalogs()
            except Exception as exc:
                error = str(exc) or type(exc).__name__
            finally:
                self.call_from_thread(self.catalog_export_finished, result, error)

        def catalog_export_finished(self, result: dict | None, error: str | None) -> None:
            button = self.query_one("#export_verified_catalog", Button)
            button.label = "EXPORT VERIFIED CATALOG"
            button.disabled = False
            self.query_one("#refresh_all_catalogs", Button).disabled = False
            self.query_one("#prepare_all_dimensions", Button).disabled = False
            if error is not None:
                self.info(f"Catalog export failed: {error}", "error")
                return
            if result is None:
                self.info("Catalog export returned no result.", "error")
                return
            self.info(
                f"EXPORT COMPLETED | CUBES: {result['cubes']} | "
                f"FIELDS: {result['fields']:,} | DESTINATION: {result['directory']}",
                "success",
            )
            self.update_catalog_summary()

        @work(thread=True, group="five-cube-discovery", exclusive=True)
        def discover_all_catalogs(self, force: bool = False) -> None:
            if not force or not self.catalog_running:
                return
            sequence = self.auth_sequence
            cookie, token = self.cookie_value(), self.token_value()
            try:
                if self.cookie_state != "VALID":
                    if force:
                        self.call_from_thread(self.info, "Saved catalogs loaded; live refresh requires a valid session.", "warning")
                    return
                for code in QUERY_PROFILES:
                    for cube_type in CUBE_TYPES:
                        if sequence != self.auth_sequence:
                            return
                        try:
                            result = discover_fields(code, token, cookie, cube_type=cube_type)
                            self.call_from_thread(self.catalog_discovery_event,
                                                  code, cube_type, result, None)
                        except Exception as exc:
                            error = str(exc)[:150]
                            self.call_from_thread(self.catalog_discovery_event,
                                                  code, cube_type, None, error)
                            if any(x in error.upper() for x in ("SESSION_EXPIRED", "SESSION EXPIRED", "HTTP 401", "HTTP 403")):
                                return
            finally:
                # This is the only post-startup snapshot reload, triggered by
                # the explicit REFRESH CATALOGS button.
                from oic_all_dim_prepare import load_all_catalogs
                try:
                    snapshot = load_all_catalogs()
                except Exception as exc:
                    self.call_from_thread(self.info, "Saved catalog reload failed: " + type(exc).__name__, "error")
                    snapshot = None
                self.call_from_thread(self.catalog_refresh_finished, snapshot)

        def catalog_refresh_finished(self, snapshot) -> None:
            if snapshot is not None:
                self.catalog_snapshot = snapshot
                self.field_catalog = snapshot["MV/FY"][1]
            self.catalog_running = False
            for button_id in ("refresh_all_catalogs","prepare_all_dimensions","export_verified_catalog"):
                self.query_one(f"#{button_id}", Button).disabled = self.busy
            self.update_catalog_summary()

        def catalog_discovery_event(self, code: str, cube_type: str,
                                    catalog: dict | None, error: str | None) -> None:
            key = f"{code}/{cube_type}"
            if error:
                self.info(f"{key} refresh failed; saved catalog kept: {error}", "warning")
            elif catalog:
                summary = catalog_verification_summary(catalog)
                self.info(f"{key}: VERIFIED {summary['verified']}/{summary['total']}; "
                          f"{summary['selectable']} selectable.", "success")

        # ====================================================
        # SELECTED DATA EXTRACTION / WINDOW HANDOFF
        # ====================================================

        @staticmethod
        def _batch_pid_alive(pid) -> bool:
            try:
                pid = int(pid)
                if pid <= 0 or sys.platform != "win32":
                    return False
                import ctypes
                handle = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)
                if not handle:
                    return False
                ctypes.windll.kernel32.CloseHandle(handle)
                return True
            except (TypeError, ValueError, OSError):
                return False

        def reconnect_active_batch(self) -> None:
            root = resolve_output_path(self.output_value()).resolve() / "Runs"
            if not root.is_dir():
                return
            runs = sorted((p for p in root.glob("*") if (p / "batch_config.json").is_file()),
                          reverse=True)
            self.last_run_path = runs[0] if runs else None
            for path in runs:
                manifest_path = path / "batch_manifest.json"
                try:
                    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
                except (OSError, ValueError):
                    continue
                if (manifest.get("status") in ("RUNNING", "STARTING", "STOPPING")
                        and self._batch_pid_alive(manifest.get("pid"))):
                    self.active_run = path
                    self.last_run_path = path
                    self.query_one("#start", Button).disabled = True
                    self.query_one("#stop", Button).disabled = False
                    self.query_one("#force_stop", Button).disabled = False
                    self.info(f"Reconnected to active batch: {path.name}", "success")
                    self.poll_selected_run()
                    return

        def _launch_batch_process(self, path: Path) -> None:
            script = BASE_DIR / "oic_parallel_collector.py"
            if not script.is_file():
                raise FileNotFoundError(str(script))
            logfile = path / "batch.log"
            with logfile.open("a", encoding="utf-8") as handle:
                self.batch_process = subprocess.Popen(
                    [sys.executable, str(script), "--batch-run", str(path)],
                    cwd=str(BASE_DIR),
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                    stdout=handle,
                    stderr=subprocess.STDOUT,
                    close_fds=True,
                )
            self.active_run = path
            self.last_run_path = path
            self.query_one("#start", Button).disabled = True
            self.query_one("#stop", Button).disabled = False
            self.query_one("#force_stop", Button).disabled = False
            self.query_one("#status", Static).update("STARTING / " + path.name)

        def start_selected_query(self) -> None:
            if self.busy or self.catalog_running:
                return
            if self.active_run is not None:
                self.info("A batch is active. Stop or finish it before restarting.", "warning")
                return
            if self.cookie_state != "VALID" or jwt_info(self.token_value())["status"] != "VALID":
                self.info("Validate Cookie and JWT before START.", "error")
                return
            if not self.applied_queries:
                self.info("APPLY at least one query before START.", "warning")
                return
            try:
                self.save_credentials()
                root = resolve_output_path(self.output_value()).resolve()
                run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
                path = root / "Runs" / run_id
                path.mkdir(parents=True, exist_ok=False)
                entries = []
                from oic_parallel_collector import write_json
                for item in self.applied_queries:
                    qid = item["query_id"]
                    config = dict(validate_query(item["query"]))
                    plan = dict(item["task_plan"])
                    if plan.get("query_hash") != query_hash(config):
                        raise ValueError(f"{qid} changed after APPLY; re-apply to plan tasks")
                    config["run_id"] = run_id
                    config["created_at"] = datetime.now(timezone.utc).isoformat()
                    config["output_root"] = str(root)
                    qdir = path / qid
                    qdir.mkdir()
                    write_json(qdir / "query_config.json", config)
                    write_json(qdir / "task_plan.json", plan)
                    entries.append(dict(query_id=qid, code=config["code"],
                                        cube_type=config.get("cube_type", "FY"),
                                        cube=config["cube"], workers=config["workers"],
                                        status="PENDING"))
                write_json(path / "batch_config.json", {
                    "run_id": run_id, "output_root": str(root),
                    "created_at": datetime.now(timezone.utc).isoformat(),
                    "queries": entries,
                })
                self._launch_batch_process(path)
                self.info(f"Batch started: {run_id} / {len(entries)} queries.", "success")
            except Exception as exc:
                self.active_run = None
                self.info(f"Cannot start batch: {type(exc).__name__}: {exc}", "error")

        def resume_selected_query(self) -> None:
            if self.busy or self.catalog_running or self.query_one("#export_verified_catalog", Button).disabled:
                return
            if self.active_run is not None:
                self.info("Stop or finish the active batch before RESUME.", "warning")
                return
            if self.cookie_state != "VALID":
                self.info("Validate Cookie session before RESUME.", "warning")
                return
            root = resolve_output_path(self.output_value()).resolve()
            runs_dir = root / "Runs"
            eligible = sorted((p for p in runs_dir.glob("*")
                               if (p / "batch_config.json").is_file()),
                              reverse=True) if runs_dir.is_dir() else []
            path = self.last_run_path if self.last_run_path and (
                self.last_run_path / "batch_config.json").is_file() else (
                eligible[0] if eligible else None)
            if path is None:
                self.info("No previous batch to resume.", "warning")
                return
            manifest = json.loads((path / "batch_manifest.json").read_text(
                encoding="utf-8")) if (path / "batch_manifest.json").is_file() else {}
            if (manifest.get("status") in ("RUNNING", "STARTING", "STOPPING")
                    and self._batch_pid_alive(manifest.get("pid"))):
                self.info("That batch is still running; do not launch duplicate workers.", "warning")
                return
            if manifest.get("status") == "COMPLETED":
                self.info("The previous batch is already complete.", "warning")
                return
            try:
                (path / "stop.flag").unlink(missing_ok=True)
                (path / "force_stop.flag").unlink(missing_ok=True)
                self._launch_batch_process(path)
                self.info(f"Resuming batch {path.name} from saved chunks.", "success")
            except Exception as exc:
                self.active_run = None
                self.info(f"RESUME failed: {exc}", "error")

        def poll_selected_run(self) -> None:
            if self.active_run is None:
                return
            manifest = {}
            try:
                manifest = json.loads((self.active_run / "batch_manifest.json").read_text(
                    encoding="utf-8"))
            except (FileNotFoundError, OSError, ValueError, TypeError):
                pass
            status = manifest.get("status", "STARTING")
            items = manifest.get("queries", [])
            query_states = {item["query_id"]: item.get("status", "PENDING")
                            for item in items if item.get("query_id")}
            if query_states != self.query_states:
                self.query_states = query_states
                self.update_query_summary()
                completed = sum(value == "COMPLETED" for value in query_states.values())
                self.query_one("#progress", ProgressBar).update(
                    total=max(1, len(query_states)), progress=completed)
            details = " / ".join(f"{q['query_id']}:{q.get('status','?')}"
                                 for q in items[:5])
            self.query_one("#status", Static).update(
                f"BATCH {self.active_run.name} / {status} / {details}"[:190])
            process_failed = (
                self.batch_process is not None and
                self.batch_process.poll() is not None and
                status not in ("COMPLETED", "PARTIAL", "FAILED", "STOPPED")
            )
            if (self.batch_process is None and
                    status in ("RUNNING", "STARTING", "STOPPING") and
                    manifest.get("pid") and
                    not self._batch_pid_alive(manifest.get("pid"))):
                process_failed = True
            if process_failed:
                status = "FAILED"
                self.info(f"Batch supervisor exited unexpectedly. See {self.active_run / 'batch.log'}", "error")
            if status in ("COMPLETED", "PARTIAL", "FAILED", "STOPPED"):
                self.info(f"Batch ended: {self.active_run.name} / {status}.", "success" if status == "COMPLETED" else "warning")
                self.active_run = None
                self.batch_process = None
                self.query_one("#start", Button).disabled = False
                self.query_one("#stop", Button).disabled = True
                self.query_one("#force_stop", Button).disabled = True

    # ========================================================
    # DEDICATED QUERY DESIGNER — same Textual framework, separate process.
    # ========================================================

    class OICQueryDesigner(App):
        TITLE = "OIC | Query Designer"
        SUB_TITLE = "AH / FI / MR / MI / MV | FY / UY / AY"
        BINDINGS = [
            Binding("a", "field_to_x", "X / ROWS", show=False),
            Binding("s", "field_to_y", "Y / COLUMNS", show=False),
            Binding("d", "field_to_values", "VALUES", show=False),
            Binding("escape", "cancel", "Cancel", show=True),
        ]

        CSS = """
        Screen {
            background: #0d141b;
            color: #dce5e9;
        }

        Header, Footer {
            background: #17232e;
        }

        #designer_root {
            height: 1fr;
            padding: 0 2 1 2;
        }

        #designer_title {
            height: 3;
            background: #17232e;
            padding: 1 2;
            color: #c9b996;
            text-style: bold;
        }

        #designer_header {
            height: 3;
        }

        #designer_header Select {
            width: 1fr;
        }
        #designer_cube { width: 15; }

        #designer_status {
            width: 2fr;
            padding: 0 1;
        }

        #designer_panels {
            height: 1fr;
            min-height: 15;
        }

        #catalog_left {
            width: 1fr;
            margin-right: 1;
            padding: 0 1;
            border: solid #304352;
            background: #111c25;
        }

        #query_right {
            width: 1fr;
            padding: 0 1;
            border: solid #304352;
            background: #111c25;
        }

        #catalog_search {
            height: 3;
        }

        #catalog_tree {
            height: 1fr;
            overflow-y: auto;
        }

        #field_description {
            height: 3;
            color: #a8bac6;
        }

        .designer_section {
            height: 1;
            color: #c9b996;
            text-style: bold;
        }

        .designer_axis {
            height: 1fr;
            min-height: 4;
            margin-bottom: 1;
        }

        .designer_axis OptionList {
            width: 1fr;
            height: 1fr;
            border: solid #304352;
            background: #13202b;
        }

        .designer_axis_actions {
            width: 7;
            height: 1fr;
            margin-left: 1;
        }

        .designer_axis_actions Button {
            width: 7;
            min-width: 7;
            height: 3;
            min-height: 3;
            padding: 0;
            content-align: center middle;
        }

        #designer_add_controls {
            height: 3;
        }

        #designer_add_controls Button {
            min-width: 8;
            margin-right: 1;
        }

        #designer_preview {
            height: 5;
            min-height: 3;
            overflow-y: auto;
            color: #b7c9d4;
        }

        #designer_settings {
            height: 3;
        }

        #designer_settings Input {
            width: 1fr;
        }

        #designer_settings Label {
            width: auto;
            padding: 1 1;
        }

        #designer_footer {
            height: 3;
        }

        #designer_footer Button {
            margin-right: 1;
        }
        #duplicate_confirm { background: #665334; }
        #duplicate_cancel { background: #633d40; }

        Button {
            background: #2c4c5d;
            color: #e6ebed;
        }

        Button:focus {
            background: #416778;
        }

        Input {
            background: #16232e;
            color: #e5ebed;
        }
        """

        def __init__(self):
            super().__init__()
            self.code = args.category if args.category in QUERY_PROFILES else "MV"
            self.selection_path = Path(args.selection_file).resolve()
            self.cube_type = "FY"
            self.catalog = load_catalog(self.code, self.cube_type)
            self.fields = {}
            self.chosen = {"x": [], "y": [], "measures": []}
            self.by_category = {}
            self.highlighted_id = None
            self.pending_duplicate = None
            self._changing_cube = False

        def compose(self) -> ComposeResult:
            from textual.widgets import OptionList

            yield Header(show_clock=True)

            with Vertical(id="designer_root"):
                yield Static(
                    "OIC QUERY DESIGNER   /   X = ROWS   •   Y = COLUMNS",
                    id="designer_title",
                )

                with Horizontal(id="designer_header"):
                    yield Select(
                        [
                            (f"{code} — {profile['name']}", code)
                            for code, profile in QUERY_PROFILES.items()
                        ],
                        value=self.code,
                        allow_blank=False,
                        id="designer_category",
                    )

                    yield Select(
                        [(f"{kind} / {status}", kind)
                         for kind, status in cube_options(self.code)],
                        value=self.cube_type, allow_blank=False,
                        id="designer_cube",
                    )
                    yield Static(
                        "CATALOG: CACHE",
                        id="designer_status",
                    )

                with Horizontal(id="designer_panels"):

                    with Vertical(id="catalog_left"):
                        yield Static(
                            "FIELD EXPLORER / OIC CAPTION",
                            classes="designer_section",
                        )

                        yield Input(
                            placeholder="Search CAPTION / GROUP / DATAFIELD / MDX",
                            id="catalog_search",
                        )

                        yield Tree(
                            "OIC FIELD CATALOG",
                            id="catalog_tree",
                        )

                        yield Static(
                            "A = ROWS / S = COLUMNS / D = VALUES",
                            id="field_description",
                        )

                        with Horizontal(id="designer_add_controls"):
                            yield Button(
                                "+ X",
                                id="add_to_x",
                            )

                            yield Button(
                                "+ Y",
                                id="add_to_y",
                            )

                            yield Button(
                                "+ VALUES",
                                id="add_to_values",
                            )

                            yield Button(
                                "RELOAD",
                                id="reload_designer_catalog",
                            )
                            yield Button(
                                "VERIFY",
                                id="verify_designer_field",
                            )

                    with Vertical(id="query_right"):

                        yield Static(
                            "X / ROWS (Axis1)",
                            classes="designer_section",
                        )

                        with Horizontal(classes="designer_axis"):
                            yield OptionList(
                                id="designer_x",
                                compact=True,
                                markup=False,
                            )

                            with Vertical(classes="designer_axis_actions"):
                                yield Button("↑", id="x_up")
                                yield Button("↓", id="x_down")
                                yield Button("×", id="x_remove")

                        yield Static(
                            "Y / COLUMNS (Axis0)",
                            classes="designer_section",
                        )

                        with Horizontal(classes="designer_axis"):
                            yield OptionList(
                                id="designer_y",
                                compact=True,
                                markup=False,
                            )

                            with Vertical(classes="designer_axis_actions"):
                                yield Button("↑", id="y_up")
                                yield Button("↓", id="y_down")
                                yield Button("×", id="y_remove")

                        yield Static(
                            "VALUES / MEASURES",
                            classes="designer_section",
                        )

                        with Horizontal(classes="designer_axis"):
                            yield OptionList(
                                id="designer_values",
                                compact=True,
                                markup=False,
                            )

                            with Vertical(classes="designer_axis_actions"):
                                yield Button("↑", id="value_up")
                                yield Button("↓", id="value_down")
                                yield Button("×", id="value_remove")

                        yield Static(
                            "FILTERS: Not supported without validated mappings",
                            classes="designer_section",
                        )

                        with Horizontal(id="designer_settings"):
                            yield Label("WORKERS")

                            yield Input(
                                value="4",
                                id="designer_workers",
                                type="integer",
                            )

                            yield Label("X CHUNK")

                            yield Input(
                                value="1",
                                id="designer_x_chunk",
                                type="integer",
                            )

                            yield Label("Y CHUNK")

                            yield Input(
                                value="20",
                                id="designer_y_chunk",
                                type="integer",
                            )

                        yield Static(
                            "QUERY PREVIEW / LOCAL ONLY",
                            classes="designer_section",
                        )

                        yield Static(
                            "Choose mapped fields, then Preview.",
                            id="designer_preview",
                        )

                with Horizontal(id="designer_footer"):
                    yield Button(
                        "PREVIEW",
                        id="designer_preview_button",
                    )

                    yield Button(
                        "APPLY SELECTION",
                        id="designer_apply",
                        variant="success",
                    )

                    yield Button(
                        "ADD ANYWAY",
                        id="duplicate_confirm",
                    )
                    yield Button(
                        "CANCEL DUPLICATE",
                        id="duplicate_cancel",
                    )
                    yield Button(
                        "CLOSE DESIGNER",
                        id="designer_cancel",
                    )

            yield Footer()

        def on_mount(self) -> None:
            self.selection_path.parent.mkdir(parents=True, exist_ok=True)
            self.selection_path.with_suffix(".lock").write_text(
                str(os.getpid()), encoding="ascii")
            self.query_one("#duplicate_confirm", Button).display = False
            self.query_one("#duplicate_cancel", Button).display = False
            self._restore_previous_selection()
            self.update_catalog(self.code, self.cube_type)

        def on_unmount(self) -> None:
            lock = self.selection_path.with_suffix(".lock")
            try:
                if lock.read_text(encoding="ascii") == str(os.getpid()):
                    lock.unlink(missing_ok=True)
            except OSError:
                pass

        def _restore_previous_selection(self) -> None:
            try:
                envelope, _ = read_applied_queries(self.selection_path)
                items = envelope["queries"]
                if not items:
                    return
                previous = validate_query(items[-1]["query"])
                if int(previous.get("schema_version", 1)) != SCHEMA_VERSION:
                    return
                self.code = previous["code"]
                self.cube_type = previous.get("cube_type", "FY")
                self.by_category[(self.code, self.cube_type)] = {
                    axis: list(previous["selection_ids"].get(axis, []))
                    for axis in self.chosen
                }
                self._changing_cube = True
                self.query_one("#designer_category", Select).value = self.code
                control = self.query_one("#designer_cube", Select)
                control.set_options([(f"{kind} / {status}", kind)
                                     for kind, status in cube_options(self.code)])
                control.value = self.cube_type
                self._changing_cube = False
                for name, key in (("workers", "designer_workers"),
                                  ("x_chunk", "designer_x_chunk"),
                                  ("y_chunk", "designer_y_chunk")):
                    self.query_one("#" + key, Input).value = str(previous[name])
            except (OSError, ValueError, TypeError, KeyError):
                self._changing_cube = False

        def update_catalog(self, code: str, cube_type: str | None = None) -> None:
            self.code = code
            self.cube_type = cube_type or self.cube_type
            self.catalog = load_catalog(code, self.cube_type)
            self.fields = {field["id"]: field for field in self.catalog["fields"]}
            previous = self.by_category.get((code, self.cube_type), {})
            self.chosen = {
                axis: [item for item in previous.get(axis, [])
                       if item in self.fields and
                       self.fields[item].get("status") in ("MAPPED", "VERIFIED")]
                for axis in ("x", "y", "measures")
            }
            self.highlighted_id = None
            self.refresh_tree()
            self.refresh_layout()
            availability = cube_status(code, self.cube_type)["status"]
            summary = catalog_verification_summary(self.catalog)
            self.query_one("#designer_status", Static).update(
                f"CATALOG {code}/{self.cube_type}: {availability} | {self.catalog['source']} | "
                f"VERIFIED {summary['verified']}/{summary['total']} "
                f"({summary['coverage']:.2f}%) | MAPPED {summary['mapped']} | "
                f"UNVERIFIED {summary['unverified']} | UNSUPPORTED {summary['unsupported']} | "
                f"SELECTABLE {summary['selectable']}")

        def on_select_changed(self, event: Select.Changed) -> None:
            if self._changing_cube:
                return
            if event.select.id not in ("designer_category", "designer_cube"):
                return
            next_code = (str(event.value) if event.select.id == "designer_category"
                         else self.code)
            next_type = (str(event.value) if event.select.id == "designer_cube"
                         else "FY" if next_code != self.code else self.cube_type)
            if next_code not in QUERY_PROFILES or next_type not in CUBE_TYPES:
                return
            if (next_code, next_type) == (self.code, self.cube_type):
                return
            self.by_category[(self.code, self.cube_type)] = {
                axis: list(value) for axis, value in self.chosen.items()
            }
            if next_code != self.code:
                self._changing_cube = True
                select = self.query_one("#designer_cube", Select)
                select.set_options([(f"{kind} / {status}", kind)
                                    for kind, status in cube_options(next_code)])
                select.value = next_type
                self._changing_cube = False
            self.update_catalog(next_code, next_type)
            self.cancel_duplicate()

        @work(thread=True, group="designer-discovery", exclusive=True)
        def refresh_catalog_worker(self) -> None:
            code, cube_type = self.code, self.cube_type
            env = read_env()
            try:
                result = discover_fields(code, env.get("OIC_BEARER_TOKEN", ""),
                                         env.get("OIC_COOKIE", ""),
                                         cube_type=cube_type)
                self.call_from_thread(self.catalog_refreshed, code, cube_type, result, None)
            except Exception as exc:
                self.call_from_thread(self.catalog_refreshed, code, cube_type, None, str(exc)[:200])

        @work(thread=True, group="designer-verify", exclusive=True)
        def verify_selected_field(self) -> None:
            item = self.highlighted_id
            field = self.fields.get(item)
            if not field:
                self.call_from_thread(self.show_error, "Select a field in the Tree first")
                return
            env = read_env()
            code, cube_type = self.code, self.cube_type
            try:
                result = verify_field(code, cube_type, item,
                                      env.get("OIC_BEARER_TOKEN", ""),
                                      env.get("OIC_COOKIE", ""))
                self.call_from_thread(
                    self.field_verified, code, cube_type, result["caption"])
            except Exception as exc:
                self.call_from_thread(self.show_error, f"VERIFY FAILED: {exc}")

        def field_verified(self, code: str, cube_type: str, caption: str) -> None:
            if (code, cube_type) == (self.code, self.cube_type):
                self.by_category[(code, cube_type)] = {
                    axis: list(v) for axis, v in self.chosen.items()
                }
                self.update_catalog(code, cube_type)
            self.query_one("#designer_preview", Static).update(
                f"VERIFIED {code}/{cube_type}: {caption} — XMLA Execute successful")

        def on_input_changed(self, event: Input.Changed) -> None:
            if event.input.id == "catalog_search":
                self.refresh_tree()

        def refresh_tree(self) -> None:
            search = (
                self.query_one("#catalog_search", Input)
                .value.casefold()
                .strip()
            )

            tree = self.query_one("#catalog_tree", Tree)
            tree.root.remove_children()

            groups = {}
            folders = {}
            xmla_groups = {}
            xmla_folders = {}

            visible = 0

            fields = self.catalog["fields"]

            # Existing selectable OIC identifiers.
            # Compare MDX identifiers, not display captions.
            existing_mdx = {
                (
                    field["type"],
                    str(field["mdx"]).strip().casefold(),
                )
                for field in fields
                if field.get("source") == "GETSTRUCTURE"
                and field.get("status") in ("MAPPED", "VERIFIED")
                and field.get("mdx")
            }

            # Preserve the original OIC field order.
            # Display additional XMLA fields after existing fields.
            ordered_fields = sorted(
                fields,
                key=lambda field: (
                    field.get("source") == "XMLA_DISCOVER"
                ),
            )

            displayed_xmla = set()

            for field in ordered_fields:
                caption = field.get(
                    "caption",
                    field.get("display", ""),
                )

                field_type = field.get("type")
                source = field.get("source")
                status = field.get("status")
                mdx = str(field.get("mdx") or "").strip()

                details = field.get("details") or {}

                is_xmla = source == "XMLA_DISCOVER"

                # Hide duplicate XMLA identifiers when the same
                # selectable identifier already exists in GetStructure.
                if is_xmla and mdx:
                    identifier = (
                        field_type,
                        mdx.casefold(),
                    )

                    if identifier in existing_mdx:
                        continue

                    if identifier in displayed_xmla:
                        continue

                # Preserve original OIC groups and folders.
                group = field.get("group_dimension") or (
                    "ข้อมูล"
                    if field_type == "MEASURE"
                    else "มุมมอง"
                )

                folder = field.get("dimension") or ""

                xmla_folder = ""

                if is_xmla:
                    # Apply dynamically to all insurance categories
                    # and all FY / UY / AY cubes.
                    group = (
                        "ข้อมูล"
                        if field_type == "MEASURE"
                        else "มุมมอง"
                    )

                    if field_type == "DIMENSION":
                        # Extract the actual Dimension name.
                        #
                        # Example:
                        # [DIM CT COUNTRY INS].[COUNTRY NAME EN]
                        #
                        # Folder:
                        # DIM CT COUNTRY INS

                        match = re.match(
                            r"^\[((?:[^\]]|\]\])+)\]\.",
                            mdx,
                        )

                        if match:
                            xmla_folder = (
                                match.group(1)
                                .replace("]]", "]")
                                .strip()
                            )
                        else:
                            xmla_folder = (
                                folder
                                or details.get("DIMENSION_UNIQUE_NAME")
                                or "OTHER DIMENSIONS"
                            )

                    elif field_type == "MEASURE":
                        # Use the existing Measure Display Folder
                        # when available.
                        xmla_folder = (
                            details.get("MEASURE_DISPLAY_FOLDER")
                            or folder
                            or ""
                        )

                # Search original captions, metadata, identifiers,
                # and additional XMLA folders.
                terms = " ".join(
                    str(value or "")
                    for value in (
                        caption,
                        group,
                        folder,
                        xmla_folder,
                        mdx,
                        source,
                        details.get("DATAFIELD"),
                        details.get("DATATABLE"),
                    )
                )

                if search and search not in terms.casefold():
                    continue

                # Create the main OIC group.
                if group not in groups:
                    groups[group] = tree.root.add(
                        group,
                        expand=True,
                    )

                parent = groups[group]

                # ============================================
                # ADDITIONAL XMLA FIELDS
                # ============================================

                if is_xmla:
                    if group not in xmla_groups:
                        xmla_groups[group] = parent.add(
                            "XMLA ADDITIONAL FIELDS",
                            expand=bool(search),
                        )

                    parent = xmla_groups[group]

                    if xmla_folder:
                        folder_key = (
                            group,
                            xmla_folder,
                        )

                        if folder_key not in xmla_folders:
                            xmla_folders[folder_key] = parent.add(
                                xmla_folder,
                                expand=bool(search),
                            )

                        parent = xmla_folders[folder_key]

                # ============================================
                # ORIGINAL OIC FIELDS
                # ============================================

                elif folder:
                    folder_key = (
                        group,
                        folder,
                    )

                    if folder_key not in folders:
                        folders[folder_key] = parent.add(
                            folder,
                            expand=True,
                        )

                    parent = folders[folder_key]

                # ============================================
                # FIELD STATUS
                # ============================================

                if status == "UNSUPPORTED":
                    suffix = "  [UNSUPPORTED]"

                elif status not in ("MAPPED", "VERIFIED"):
                    suffix = "  [UNVERIFIED]"

                else:
                    suffix = ""

                parent.add_leaf(
                    caption + suffix,
                    data=field["id"],
                )

                if is_xmla and mdx:
                    displayed_xmla.add(
                        (
                            field_type,
                            mdx.casefold(),
                        )
                    )

                visible += 1

            tree.root.expand()

            self.query_one(
                "#field_description",
                Static,
            ).update(
                f"{visible} visible / "
                f"{len(fields)} total"
                " · A = ROWS / S = COLUMNS / D = VALUES"
            )

        def on_tree_node_highlighted(self, event: Tree.NodeHighlighted) -> None:
            if event.control.id == "catalog_tree":
                self.highlighted_id = event.node.data if isinstance(event.node.data, str) else None

        def on_tree_node_selected(self, event: Tree.NodeSelected) -> None:
            if event.control.id != "catalog_tree":
                return
            item = event.node.data
            if not isinstance(item, str) or item not in self.fields:
                return
            self.highlighted_id = item
            field = self.fields[item]
            if field["type"] == "MEASURE":
                self.toggle_value(item)
            else:
                self.query_one("#field_description", Static).update(
                    f"{field['caption']} · {field.get('status')} · "
                    f"{field.get('mapping_reason', 'Select A=ROWS / S=COLUMNS')}"
                )

        def action_field_to_x(self) -> None:
            self.move_field("x")

        def action_field_to_y(self) -> None:
            self.move_field("y")

        def action_field_to_values(self) -> None:
            if isinstance(self.focused, Input):
                return

            item = self.highlighted_id
            field = self.fields.get(item)

            if field and field["type"] == "MEASURE":
                self.toggle_value(item)

        def move_field(self, axis: str) -> None:
            if isinstance(self.focused, Input):
                return

            item = self.highlighted_id
            field = self.fields.get(item)

            if not field or field["type"] != "DIMENSION":
                return

            if (
                field.get("status") not in ("MAPPED", "VERIFIED")
                or not field.get("mdx")
            ):
                self.show_error(
                    f"{field['caption']}: "
                    f"mapping not available ({field.get('status')}): "
                    f"{field.get('mapping_reason', 'No verified MDX identifier')}"
                )
                return

            if item in self.chosen[axis]:
                return

            for selected in self.chosen["x"] + self.chosen["y"]:
                if (
                    selected != item
                    and self.fields[selected].get("mdx") == field["mdx"]
                ):
                    self.show_error(
                        "The hierarchy is already selected "
                        "under another display item"
                    )
                    return

            other_axis = "y" if axis == "x" else "x"

            if item in self.chosen[other_axis]:
                self.chosen[other_axis].remove(item)
                self.refresh_axis(other_axis)

            self.chosen[axis].append(item)

            if (
                axis == "y"
                and field["mdx"] == QUERY_PROFILES[self.code].get("year")
                and self.query_one(
                    "#designer_y_chunk",
                    Input,
                ).value == "20"
            ):
                self.query_one(
                    "#designer_y_chunk",
                    Input,
                ).value = "1"

            self.refresh_axis(
                axis,
                select_id=item,
            )

        def toggle_value(self, item: str) -> None:
            field = self.fields.get(item)

            if not field or field["type"] != "MEASURE":
                return

            if item in self.chosen["measures"]:
                self.chosen["measures"].remove(item)

            elif (
                field.get("status") not in ("MAPPED", "VERIFIED")
                or not field.get("mdx")
            ):
                self.show_error(
                    f"{field['caption']}: "
                    f"{field.get('status')} — "
                    f"{field.get('mapping_reason', 'mapping not verified')}"
                )
                return

            elif any(
                self.fields[selected].get("mdx") == field["mdx"]
                for selected in self.chosen["measures"]
            ):
                self.show_error(
                    "This base Measure is already selected"
                )
                return

            else:
                self.chosen["measures"].append(item)

            self.refresh_axis(
                "measures",
                select_id=item,
            )

        def refresh_axis(
            self,
            axis: str,
            select_id: str | None = None,
        ) -> None:
            from textual.widgets import OptionList
            from textual.widgets.option_list import Option

            control = {
                "x": "designer_x",
                "y": "designer_y",
                "measures": "designer_values",
            }[axis]

            widget = self.query_one(
                f"#{control}",
                OptionList,
            )

            items = self.chosen[axis]

            previous = widget.highlighted_option
            previous_id = (
                previous.id
                if previous is not None
                else None
            )

            current_ids = [
                option.id
                for option in widget.options
            ]

            if current_ids == items:
                return

            widget.clear_options()
            widget.add_options(
                [
                    Option(
                        f"{index + 1}. "
                        f"{self.fields[item]['caption']}",
                        id=item,
                    )
                    for index, item in enumerate(items)
                ]
            )

            if not items:
                widget.highlighted = None
                return

            target = (
                select_id
                if select_id in items
                else previous_id
                if previous_id in items
                else items[-1]
            )

            widget.highlighted = items.index(target)

        def refresh_layout(self) -> None:
            for axis in ("x", "y", "measures"):
                self.refresh_axis(axis)

        def modify_axis(self, operation: str) -> None:
            from textual.widgets import OptionList

            axis = (
                "measures"
                if operation.startswith("value_")
                else operation[0]
            )

            control = {
                "x": "designer_x",
                "y": "designer_y",
                "measures": "designer_values",
            }[axis]

            widget = self.query_one(
                f"#{control}",
                OptionList,
            )

            selected = widget.highlighted_option

            if selected is None:
                self.show_error(
                    "Select a field in that axis first"
                )
                return

            item = selected.id
            items = self.chosen[axis]

            if item not in items:
                self.show_error(
                    "Select a field in that axis first"
                )
                return

            index = items.index(item)

            if operation.endswith("remove"):
                items.pop(index)

                target = (
                    items[min(index, len(items) - 1)]
                    if items
                    else None
                )

            else:
                next_index = (
                    index - 1
                    if operation.endswith("up")
                    else index + 1
                )

                if not 0 <= next_index < len(items):
                    return

                items[index], items[next_index] = (
                    items[next_index],
                    items[index],
                )

                target = item

            self.refresh_axis(
                axis,
                select_id=target,
            )
        def build_query(self) -> dict:
            profile = cube_profile(self.code, self.cube_type)

            selections = {
                axis: list(self.chosen[axis])
                for axis in ("x", "y", "measures")
            }

            if not selections["x"]:
                raise ValueError("Select at least one X / ROWS field")

            if not selections["y"]:
                raise ValueError("Select at least one Y / COLUMNS field")

            if not selections["measures"]:
                raise ValueError("Select at least one VALUES / MEASURES field")

            selected_ids = list(dict.fromkeys(
                selections["x"]
                + selections["y"]
                + selections["measures"]
            ))

            selected_fields = []

            for item in selected_ids:
                field = self.fields.get(item)

                if field is None:
                    raise ValueError(
                        f"Selected field not found: {item}"
                    )

                if (
                    field.get("status") not in ("MAPPED", "VERIFIED")
                    or not field.get("mdx")
                ):
                    raise ValueError(
                        f"Field cannot be applied: "
                        f"{field.get('caption', item)}"
                    )

                selected_fields.append({
                    "id": field["id"],
                    "mdx": field["mdx"],
                    "caption": field["caption"],
                    "type": field["type"],
                    "status": field["status"],
                })

            config = {
                "schema_version": SCHEMA_VERSION,
                "code": self.code,
                "cube_type": self.cube_type,
                "cube": profile["cube"],
                "catalog": profile["catalog"],
                "x": [
                    self.fields[item]["mdx"]
                    for item in selections["x"]
                ],
                "y": [
                    self.fields[item]["mdx"]
                    for item in selections["y"]
                ],
                "measures": [
                    self.fields[item]["mdx"]
                    for item in selections["measures"]
                ],
                "selection_ids": selections,
                "selected_fields": selected_fields,
                "filters": [],
                "workers": int(
                    self.query_one(
                        "#designer_workers", Input
                    ).value
                ),
                "x_chunk": int(
                    self.query_one(
                        "#designer_x_chunk", Input
                    ).value
                ),
                "y_chunk": int(
                    self.query_one(
                        "#designer_y_chunk", Input
                    ).value
                ),
                "timeout": 60,
                "delay": 1.0,
            }

            return validate_query(config)
        
        def preview_query(self) -> None:
            try:
                query = self.build_query()
                mdx = selected_mdx(query)

                self.query_one(
                    "#designer_preview",
                    Static,
                ).update(
                    f"CODE: {query['code']} / {query.get('cube_type', 'FY')} / {query['cube']}\n"
                    f"X ROWS (Axis1): {len(query['x'])}   "
                    f"Y COLUMNS (Axis0): {len(query['y'])}   "
                    f"VALUES: {len(query['measures'])}\n"
                    f"Workers: {query['workers']}  "
                    f"Chunks: {query['x_chunk']} × {query['y_chunk']}\n"
                    f"NOTE: CrossJoin queries may be expensive.\n"
                    f"{mdx}"
                )

            except Exception as exc:
                self.show_error(
                    f"PREVIEW FAILED: {type(exc).__name__}: {exc}"
                )

        def show_error(self, message: str) -> None:
            self.query_one("#designer_preview", Static).update("VALIDATION: " + message[:450])

        def cancel_duplicate(self) -> None:
            self.pending_duplicate = None
            self.query_one("#duplicate_confirm", Button).display = False
            self.query_one("#duplicate_cancel", Button).display = False

        def apply_selection(self, force_duplicate: bool = False) -> None:
            try:
                from oic_parallel_collector import build_selected_task_plan
                config = self.build_query()
                fingerprint = applied_query_identity(config)
                selection = self.selection_path

                if force_duplicate and self.pending_duplicate != fingerprint:
                    self.cancel_duplicate()
                    raise ValueError(
                        "Query changed after duplicate warning; press APPLY again"
                    )

                with applied_list_write_lock(selection):
                    envelope, legacy = read_applied_queries(selection)
                    entries = envelope["queries"]

                    duplicate = next(
                        (item for item in entries
                         if applied_query_identity(item["query"]) == fingerprint),
                        None,
                    )
                    if duplicate and not force_duplicate:
                        self.pending_duplicate = fingerprint
                        self.query_one("#duplicate_confirm", Button).display = True
                        self.query_one("#duplicate_cancel", Button).display = True
                        self.show_error(
                            f"DUPLICATE: matches {duplicate['query_id']} "
                            "(same category, cube, X, Y, VALUES, Filters). "
                            "Press ADD ANYWAY or CANCEL DUPLICATE."
                        )
                        return

                    # Stable, monotonic IDs: never reuse removed Query IDs.
                    sequence = int(envelope["next_query_sequence"])
                    existing_ids = {item["query_id"] for item in entries}
                    query_id = f"Q{sequence:03d}"
                    while query_id in existing_ids:
                        sequence += 1
                        query_id = f"Q{sequence:03d}"
                    config["query_id"] = query_id
                    plan = build_selected_task_plan(config)
                    entries.append({
                        "query_id": query_id,
                        "query": config,
                        "task_plan": plan,
                        "applied_at": datetime.now(timezone.utc).isoformat(),
                    })
                    envelope["next_query_sequence"] = sequence + 1
                    save_applied_queries(
                        selection, envelope, backup_legacy=legacy,
                    )

                self.cancel_duplicate()
                self.query_one("#designer_preview", Static).update(
                    f"APPLIED {query_id} / {config['code']}/"
                    f"{config.get('cube_type', 'FY')} / {config['cube']}\n"
                    f"Prepared {len(plan['assignments'])} worker allocations. "
                    "Task count is DYNAMIC until member discovery.\n"
                    f"APPLIED QUERIES: {len(entries)}. "
                    "Designer remains open."
                )
            except Exception as exc:
                self.show_error(
                    f"APPLY FAILED: {type(exc).__name__}: {exc}"
                )

        def catalog_refreshed(self, code: str, cube_type: str,
                              catalog: dict | None, error: str | None):
            if code == self.code and cube_type == self.cube_type:
                self.by_category[(code, cube_type)] = {
                    axis: list(v) for axis, v in self.chosen.items()
                }
                self._changing_cube = True
                control = self.query_one("#designer_cube", Select)
                control.set_options([(f"{kind} / {status}", kind)
                                     for kind, status in cube_options(code)])
                control.value = cube_type
                self._changing_cube = False
                self.update_catalog(code, cube_type)
            if error:
                self.show_error("Live refresh failed; saved catalog loaded: " + error)

        def action_cancel(self) -> None:
            self.exit()

        def on_button_pressed(
            self,
            event: Button.Pressed,
        ) -> None:
            operation = event.button.id

            if operation == "add_to_x":
                self.move_field("x")

            elif operation == "add_to_y":
                self.move_field("y")

            elif operation == "add_to_values":
                self.action_field_to_values()

            elif operation in (
                "x_up",
                "x_down",
                "x_remove",
                "y_up",
                "y_down",
                "y_remove",
                "value_up",
                "value_down",
                "value_remove",
            ):
                self.modify_axis(operation)

            elif operation == "reload_designer_catalog":
                self.refresh_catalog_worker()
            elif operation == "verify_designer_field":
                self.verify_selected_field()

            elif operation == "designer_preview_button":
                self.preview_query()

            elif operation == "designer_apply":
                self.apply_selection()

            elif operation == "duplicate_confirm":
                self.apply_selection(force_duplicate=True)
            elif operation == "duplicate_cancel":
                self.cancel_duplicate()
            elif operation == "designer_cancel":
                self.action_cancel()

    # ========================================================
    # LAUNCH TEXTUAL APPLICATION
    # ========================================================

    if return_app:
        return OICApp()
    if getattr(args, "designer", False):
        OICQueryDesigner().run()
    else:
        OICApp().run()


# ============================================================
# ENTRY POINT
# ============================================================

def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "OIC XMLA Collector "
            "/ Live Authentication Terminal UI"
        )
    )

    parser.add_argument("--designer", action="store_true",
                        help="Open the dedicated Textual query designer")
    parser.add_argument("--selection-file", type=str, default="output/applied_queries.json")
    parser.add_argument("--category", choices=list(PROFILES), default="MV")

    parser.add_argument(
        "--cli",
        action="store_true",
        help=(
            "Open interactive Terminal UI "
            "(same as running without arguments)"
        ),
    )

    parser.add_argument(
        "--headless",
        action="store_true",
        help=(
            "Run collection without Terminal UI"
        ),
    )

    parser.add_argument(
        "--test-only",
        action="store_true",
        help=(
            "Run a headless XMLA API test "
            "without exporting Excel"
        ),
    )

    parser.add_argument(
        "--codes",
        nargs="+",
        choices=list(PROFILES),
        default=list(PROFILES),
        help="Selected insurance categories",
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
        help="XMLA read timeout in seconds",
    )

    parser.add_argument(
        "--retries",
        type=int,
        default=DEFAULT_RETRIES,
        help="Maximum XMLA request attempts",
    )

    args = parser.parse_args()

    if (
        args.timeout < 1
        or args.retries < 1
    ):
        parser.error(
            "--timeout and --retries must be positive"
        )

    # --cli is now an interactive UI alias.
    # It never starts authentication before UI startup.
    #
    # Explicit headless mode remains available for
    # scripts and scheduled jobs.

    if args.headless or args.test_only:
        return run_headless(args)

    try:
        launch_terminal_ui(args)
        return 0

    except RuntimeError as exc:
        print(
            f"ERROR: {exc}",
            file=sys.stderr,
        )
        return 1

    except KeyboardInterrupt:
        return 130


if __name__ == "__main__":
    sys.exit(main())