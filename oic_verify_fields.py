
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC FIELD VERIFIER
==================
File: oic_verify_fields.py

Standalone OIC field discovery and XMLA verification.

Categories:
    AH / FI / MR / MI / MV

Cube types:
    FY / UY / AY

Authentication (.env):
    OIC_COOKIE=...
    OIC_BEARER_TOKEN=...

Install:
    pip install curl_cffi rich

Commands:
    python oic_verify_fields.py
    python oic_verify_fields.py --code MV
    python oic_verify_fields.py --code MV --cube-type FY
    python oic_verify_fields.py --unverified-only
    python oic_verify_fields.py --force
    python oic_verify_fields.py --summary

Output:
    output/field_verification/
        summary.json
        all_fields.csv
        AH_FY.json
        AH_UY.json
        AH_AY.json
        ...

Persistence:
    - Windows-safe atomic writes with retry.
    - Automatic recovery file if destination is locked.
    - Save progress every 10 fields by default.
    - Preserve completed verification results.
    - Resume without repeating VERIFIED fields.
    - Do not turn a Cube into ERROR for a temporary file lock.
    - Do not modify the main OIC catalog.
    - Do not persist authentication secrets.
"""

from __future__ import annotations

import argparse
import base64
import csv
import errno
import hashlib
import json
import os
import re
import shutil
import sys
import time
import uuid

from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse
from xml.etree import ElementTree as ET

try:
    from curl_cffi import requests
    from curl_cffi.requests.exceptions import RequestException
except ImportError:
    print(
        "Missing dependency: curl_cffi\n"
        "Install: python -m pip install curl_cffi"
    )
    raise SystemExit(1)

try:
    from rich.console import Console
    from rich.panel import Panel
    from rich.progress import (
        BarColumn,
        Progress,
        SpinnerColumn,
        TaskProgressColumn,
        TextColumn,
        TimeElapsedColumn,
    )
    from rich.table import Table
    from rich.text import Text

    RICH_AVAILABLE = True

except ImportError:
    RICH_AVAILABLE = False


# ============================================================
# CONFIGURATION
# ============================================================

ROOT = Path(__file__).resolve().parent
ENV_FILE = ROOT / ".env"

OUTPUT_DIR = ROOT / "output" / "field_verification"

BASE_URL = "https://onlinewebibs.oic.or.th"

WEBFORM_PATH = (
    "/WebForm/Dashboard/Pivot/DynamicPivotCube.aspx"
)

WEBFORM_URL = BASE_URL + WEBFORM_PATH

TOKEN_URL = (
    BASE_URL
    + "/WebForm/Dashboard/Pivot/"
    + "PivotCube.asmx/GenToken"
)

XMLA_URL = (
    BASE_URL
    + "/xmla-proxy/api/xmla-proxy"
)

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"

MD_NS = (
    "urn:schemas-microsoft-com:"
    "xml-analysis:mddataset"
)

ET.register_namespace("soap", SOAP_NS)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)

CATEGORIES = {
    "AH": {
        "name": "อุบัติเหตุและสุขภาพ",
        "web_class": "AH",
        "catalog": "PIVOT_202609_AH",
    },
    "FI": {
        "name": "อัคคีภัยและทรัพย์สิน",
        "web_class": "FI",
        "catalog": "PIVOT_202609_FI",
    },
    "MR": {
        "name": "ทางทะเลและขนส่ง",
        "web_class": "MR",
        "catalog": "PIVOT_202609_MR",
    },
    "MI": {
        "name": "เบ็ดเตล็ด",
        "web_class": "MI",
        "catalog": "PIVOT_202609_MI",
    },
    "MV": {
        "name": "รถยนต์",
        "web_class": "MO",
        "catalog": "PIVOT_202609_MV",
    },
}

CUBE_TYPES = {
    "FY": "ปีบัญชี",
    "UY": "ปีรับประกัน",
    "AY": "ปีอุบัติเหตุ",
}

WEBFORM_METHODS = (
    "GetStructure",
    "GET_MENUTEMPLATE_LIST",
    "BindChangeTemplate",
)

XMLA_DISCOVER_TYPES = (
    "MDSCHEMA_HIERARCHIES",
    "MDSCHEMA_MEASURES",
    "MDSCHEMA_LEVELS",
)

DIMENSION_RE = re.compile(
    r"^\[[^\]\r\n]+\]\.\[[^\]\r\n]+\]$"
)

MEASURE_RE = re.compile(
    r"^\[Measures\]\.\[[^\]\r\n]+\]$",
    re.IGNORECASE,
)

GROUP_TYPES = {
    "มุมมอง": "DIMENSION",
    "ข้อมูล": "MEASURE",
}

CSV_COLUMNS = [
    "insurance_code",
    "insurance_type",
    "cube_type",
    "cube_type_th",
    "catalog",
    "cube",
    "field_id",
    "field_type",
    "field_caption",
    "field_folder",
    "data_table",
    "data_field",
    "mdx",
    "status",
    "source",
    "mapping_source",
    "mapping_reason",
    "verification_method",
    "verified_at",
    "error",
]

# Save after every N field verifications.
DEFAULT_CHECKPOINT_EVERY = 10

# Windows sharing violation / antivirus / indexing retry.
FILE_REPLACE_DELAYS = (
    0.10,
    0.20,
    0.35,
    0.50,
    0.75,
    1.00,
    1.50,
    2.00,
)

if RICH_AVAILABLE:
    console = Console(highlight=False)
else:
    console = None


# ============================================================
# TERMINAL
# ============================================================

def log(
    message: str,
    level: str = "INFO",
) -> None:
    current = datetime.now().strftime("%H:%M:%S")

    if not RICH_AVAILABLE:
        print(
            f"{current} | {level:<7} | {message}",
            flush=True,
        )
        return

    styles = {
        "INFO": "cyan",
        "OK": "green",
        "WARN": "yellow",
        "ERROR": "red",
    }

    line = Text()

    line.append(
        current,
        style="dim",
    )

    line.append("  ")

    line.append(
        f"{level:<7}",
        style=styles.get(level, "white"),
    )

    line.append("  ")
    line.append(str(message))

    console.print(line)


def print_header() -> None:
    if not RICH_AVAILABLE:
        print("=" * 72)
        print("OIC FIELD VERIFIER")
        print("FIELD DISCOVERY + XMLA EXECUTE")
        print("=" * 72)
        return

    console.print(
        Panel(
            "[bold cyan]OIC FIELD VERIFIER[/]\n"
            "[white]FIELD DISCOVERY + XMLA EXECUTE[/]\n"
            "[dim]AH / FI / MR / MI / MV | "
            "FY / UY / AY | AUTO RESUME[/]",
            border_style="cyan",
        )
    )


def print_cube_result(result: dict) -> None:
    counts = result.get("counts", {})

    log(
        f"{result['insurance_code']}/"
        f"{result['cube_type']} | "
        f"{result['status']} | "
        f"TOTAL={counts.get('TOTAL', 0)} | "
        f"VERIFIED={counts.get('VERIFIED', 0)} | "
        f"MAPPED={counts.get('MAPPED', 0)} | "
        f"UNVERIFIED={counts.get('UNVERIFIED', 0)} | "
        f"UNSUPPORTED={counts.get('UNSUPPORTED', 0)}",
        (
            "OK"
            if result["status"] == "AVAILABLE"
            else "WARN"
        ),
    )


def print_summary(summary: dict) -> None:
    cubes = summary.get("cubes", [])
    totals = summary.get("totals", {})

    if not RICH_AVAILABLE:
        print("\nOIC FIELD VERIFICATION SUMMARY")
        print("-" * 72)

        for cube in cubes:
            counts = cube.get("counts", {})

            print(
                f"{cube['insurance_code']}/"
                f"{cube['cube_type']:<2} "
                f"{cube['status']:<18} "
                f"TOTAL={counts.get('TOTAL', 0)} "
                f"VERIFIED={counts.get('VERIFIED', 0)} "
                f"MAPPED={counts.get('MAPPED', 0)}"
            )

        print("-" * 72)
        print(
            json.dumps(
                totals,
                ensure_ascii=False,
                indent=2,
            )
        )
        return

    table = Table(
        title="OIC FIELD VERIFICATION SUMMARY",
        expand=True,
    )

    table.add_column("CUBE", style="cyan")
    table.add_column("STATUS")
    table.add_column("FIELDS", justify="right")
    table.add_column("VERIFIED", justify="right")
    table.add_column("MAPPED", justify="right")
    table.add_column("UNVERIFIED", justify="right")
    table.add_column("UNSUPPORTED", justify="right")

    for item in cubes:
        counts = item.get("counts", {})

        table.add_row(
            f"{item['insurance_code']}/{item['cube_type']}",
            str(item["status"]),
            str(counts.get("TOTAL", 0)),
            str(counts.get("VERIFIED", 0)),
            str(counts.get("MAPPED", 0)),
            str(counts.get("UNVERIFIED", 0)),
            str(counts.get("UNSUPPORTED", 0)),
        )

    console.print(table)

    total = totals.get("TOTAL", 0)
    verified = totals.get("VERIFIED", 0)

    percentage = (
        100.0 * verified / total
        if total
        else 0.0
    )

    console.print(
        f"FIELDS: {total:,}  |  "
        f"VERIFIED: {verified:,} ({percentage:.2f}%)  |  "
        f"MAPPED: {totals.get('MAPPED', 0):,}  |  "
        f"UNVERIFIED: {totals.get('UNVERIFIED', 0):,}  |  "
        f"UNSUPPORTED: {totals.get('UNSUPPORTED', 0):,}"
    )


# ============================================================
# UTILITIES
# ============================================================

def utc_now() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat(timespec="seconds")


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def clean(value) -> str:
    return str(
        "" if value is None else value
    ).strip()


def normalized(value) -> str:
    return " ".join(
        clean(value).casefold().split()
    )


def valid_identifier(
    value: str,
    field_type: str,
) -> bool:
    if not isinstance(value, str):
        return False

    pattern = (
        MEASURE_RE
        if field_type == "MEASURE"
        else DIMENSION_RE
    )

    return bool(
        pattern.fullmatch(value)
    )


def make_field_id(
    code: str,
    cube_type: str,
    source: str,
    record: dict,
) -> str:
    encoded = json.dumps(
        {
            "code": code,
            "cube_type": cube_type,
            "source": source,
            "record": record,
        },
        ensure_ascii=False,
        sort_keys=True,
        default=str,
    ).encode("utf-8")

    return hashlib.sha256(
        encoded
    ).hexdigest()[:24]


def normalize_cookie(value: str) -> str:
    value = clean(value)

    if value.lower().startswith("cookie:"):
        value = value[7:].strip()

    if value.startswith("OIC_COOKIE="):
        value = value[len("OIC_COOKIE="):].strip()

    return value


def normalize_token(value: str) -> str:
    value = clean(value)

    if value.lower().startswith("bearer "):
        value = value[7:].strip()

    return value


def jwt_exp(token: str) -> int:
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)

        decoded = base64.urlsafe_b64decode(
            payload
        )

        data = json.loads(
            decoded.decode("utf-8")
        )

        return int(
            data.get("exp") or 0
        )

    except Exception:
        return 0


def safe_error(
    exception: Exception,
    credentials: tuple[str, ...] = (),
) -> str:
    message = (
        f"{type(exception).__name__}: {exception}"
    )

    for secret in credentials:
        if secret and len(secret) >= 8:
            message = message.replace(
                secret,
                "[REDACTED]",
            )

    return message[:600]


def read_env() -> dict[str, str]:
    result = {}

    if ENV_FILE.is_file():
        content = ENV_FILE.read_text(
            encoding="utf-8-sig"
        )

        for original in content.splitlines():
            line = original.strip()

            if not line or line.startswith("#"):
                continue

            if line.startswith("export "):
                line = line[7:].strip()

            if "=" not in line:
                continue

            key, value = line.split("=", 1)

            key = key.strip()
            value = value.strip()

            if (
                len(value) >= 2
                and value[0] == value[-1]
                and value[0] in ("'", '"')
            ):
                value = value[1:-1]

            result[key] = value

    for key in (
        "OIC_COOKIE",
        "OIC_BEARER_TOKEN",
    ):
        environment_value = os.environ.get(key)

        if environment_value:
            result[key] = environment_value

    return result


# ============================================================
# WINDOWS-SAFE FILE PERSISTENCE
# ============================================================

def is_windows_file_lock(exc: OSError) -> bool:
    if isinstance(exc, PermissionError):
        return True

    if getattr(exc, "winerror", None) in (
        5,
        32,
        33,
    ):
        return True

    return (
        getattr(exc, "errno", None)
        in (
            errno.EACCES,
            errno.EPERM,
        )
    )


def recovery_pattern(path: Path) -> str:
    return (
        f"{path.stem}.recovery.*"
        f"{path.suffix}"
    )


def recovery_files(path: Path) -> list[Path]:
    try:
        return [
            item
            for item in path.parent.glob(
                recovery_pattern(path)
            )
            if item.is_file()
        ]
    except OSError:
        return []


def read_json(
    path: Path,
    default=None,
):
    try:
        with path.open(
            "r",
            encoding="utf-8-sig",
        ) as stream:
            return json.load(stream)

    except (
        OSError,
        ValueError,
        UnicodeError,
    ):
        return default


def read_latest_json(
    path: Path,
    default=None,
    *,
    expected_code: str | None = None,
    expected_cube_type: str | None = None,
):
    candidates = [path] + recovery_files(path)

    candidates.sort(
        key=lambda item: (
            item.stat().st_mtime_ns
            if item.exists()
            else 0
        ),
        reverse=True,
    )

    for candidate in candidates:
        value = read_json(
            candidate,
            None,
        )

        if not isinstance(value, dict):
            continue

        if (
            expected_code is not None
            and value.get("insurance_code")
            != expected_code
        ):
            continue

        if (
            expected_cube_type is not None
            and value.get("cube_type")
            != expected_cube_type
        ):
            continue

        return value

    return default


def write_temp_file(
    path: Path,
    payload: bytes,
) -> Path:
    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary = path.with_name(
        path.name
        + "."
        + uuid.uuid4().hex
        + ".tmp"
    )

    with temporary.open("xb") as stream:
        stream.write(payload)
        stream.flush()
        os.fsync(stream.fileno())

    return temporary


def write_recovery_file(
    path: Path,
    temporary: Path,
) -> Path:
    """
    If the destination JSON is locked, preserve the
    complete new report under a unique recovery name.

    The next run automatically reads the newest valid
    report, including recovery files.
    """

    recovery = path.with_name(
        f"{path.stem}.recovery."
        f"{time.time_ns()}."
        f"{uuid.uuid4().hex[:8]}"
        f"{path.suffix}"
    )

    with temporary.open("rb") as source:
        with recovery.open("xb") as target:
            shutil.copyfileobj(
                source,
                target,
                length=1024 * 1024,
            )

            target.flush()
            os.fsync(target.fileno())

    return recovery


def cleanup_recovery_files(
    path: Path,
) -> None:
    for candidate in recovery_files(path):
        try:
            candidate.unlink(
                missing_ok=True
            )
        except OSError:
            pass


def commit_file(
    path: Path,
    temporary: Path,
) -> Path:
    """
    Commit a temporary file using os.replace().

    WinError 5 / 32 / 33:
        Retry first.
        If still locked, write a unique recovery file.

    A temporary destination lock must not discard
    already-verified fields.
    """

    last_error = None

    try:
        for delay in FILE_REPLACE_DELAYS:
            try:
                os.replace(
                    temporary,
                    path,
                )

                cleanup_recovery_files(path)

                return path

            except OSError as exc:
                if not is_windows_file_lock(exc):
                    raise

                last_error = exc
                time.sleep(delay)

        try:
            os.replace(
                temporary,
                path,
            )

            cleanup_recovery_files(path)

            return path

        except OSError as exc:
            if not is_windows_file_lock(exc):
                raise

            last_error = exc

        # A locked destination is not a lost report.
        # Save the new version in a separate recovery file.

        recovery = write_recovery_file(
            path,
            temporary,
        )

        log(
            f"FILE LOCK: {path.name} | "
            f"saved recovery: {recovery.name}",
            "WARN",
        )

        return recovery

    except OSError as exc:
        if last_error is not None:
            raise OSError(
                f"Cannot save {path.name}; "
                f"original error: {last_error}; "
                f"latest error: {exc}"
            ) from exc

        raise

    finally:
        try:
            temporary.unlink(
                missing_ok=True
            )
        except OSError:
            pass


def atomic_json(
    path: Path,
    value: dict,
) -> Path:
    payload = json.dumps(
        value,
        ensure_ascii=False,
        indent=2,
        default=str,
    ).encode("utf-8")

    temporary = write_temp_file(
        path,
        payload,
    )

    return commit_file(
        path,
        temporary,
    )


def report_path(
    code: str,
    cube_type: str,
) -> Path:
    return (
        OUTPUT_DIR
        / f"{code}_{cube_type}.json"
    )


def load_cube_report(
    code: str,
    cube_type: str,
) -> dict | None:
    return read_latest_json(
        report_path(
            code,
            cube_type,
        ),
        None,
        expected_code=code,
        expected_cube_type=cube_type,
    )


# ============================================================
# XMLA REQUEST BUILDERS
# ============================================================

def make_discover_request(
    catalog: str,
    request_type: str,
    cube: str | None = None,
) -> bytes:
    supported = {
        "MDSCHEMA_CUBES",
        "MDSCHEMA_DIMENSIONS",
        "MDSCHEMA_HIERARCHIES",
        "MDSCHEMA_LEVELS",
        "MDSCHEMA_MEASURES",
    }

    if request_type not in supported:
        raise ValueError(
            f"Unsupported XMLA request: {request_type}"
        )

    envelope = ET.Element(
        f"{{{SOAP_NS}}}Envelope"
    )

    body = ET.SubElement(
        envelope,
        f"{{{SOAP_NS}}}Body",
    )

    discover = ET.SubElement(
        body,
        f"{{{XMLA_NS}}}Discover",
    )

    ET.SubElement(
        discover,
        f"{{{XMLA_NS}}}RequestType",
    ).text = request_type

    restrictions = ET.SubElement(
        ET.SubElement(
            discover,
            f"{{{XMLA_NS}}}Restrictions",
        ),
        f"{{{XMLA_NS}}}RestrictionList",
    )

    ET.SubElement(
        restrictions,
        f"{{{XMLA_NS}}}CATALOG_NAME",
    ).text = catalog

    if cube and request_type != "MDSCHEMA_CUBES":
        ET.SubElement(
            restrictions,
            f"{{{XMLA_NS}}}CUBE_NAME",
        ).text = cube

    properties = ET.SubElement(
        ET.SubElement(
            discover,
            f"{{{XMLA_NS}}}Properties",
        ),
        f"{{{XMLA_NS}}}PropertyList",
    )

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}Catalog",
    ).text = catalog

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}LocaleIdentifier",
    ).text = "9"

    return ET.tostring(
        envelope,
        encoding="utf-8",
    )


def make_execute_request(
    catalog: str,
    mdx: str,
) -> bytes:
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
    ).text = mdx

    properties = ET.SubElement(
        ET.SubElement(
            execute,
            f"{{{XMLA_NS}}}Properties",
        ),
        f"{{{XMLA_NS}}}PropertyList",
    )

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}Catalog",
    ).text = catalog

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}ShowHiddenCubes",
    ).text = "True"

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}LocaleIdentifier",
    ).text = "9"

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}Timeout",
    ).text = "60"

    return ET.tostring(
        envelope,
        encoding="utf-8",
    )


def parse_soap(
    content: bytes,
) -> ET.Element:
    root = ET.fromstring(content)

    for element in root.iter():
        if local_name(element.tag) != "Fault":
            continue

        message = " ".join(
            clean(item.text)
            for item in element.iter()
            if item.text
        )

        raise ValueError(
            f"XMLA SOAP Fault: {message[:400]}"
        )

    return root


def parse_discover_rows(
    content: bytes,
) -> list[dict[str, str]]:
    root = parse_soap(content)

    rows = []

    for element in root.iter():
        if local_name(element.tag).lower() != "row":
            continue

        row = {
            local_name(child.tag).upper(): (
                child.text or ""
            )
            for child in element
        }

        if row:
            rows.append(row)

    return rows


def validate_execute_response(
    content: bytes,
) -> dict:
    root = parse_soap(content)

    dataset = next(
        (
            node
            for node in root.iter()
            if (
                local_name(node.tag) == "root"
                and node.tag.startswith(
                    "{" + MD_NS + "}"
                )
            )
        ),
        None,
    )

    if dataset is None:
        raise ValueError(
            "XMLA Execute did not return "
            "a multidimensional dataset"
        )

    axes = [
        node
        for node in dataset.iter()
        if local_name(node.tag) == "Axis"
    ]

    if not axes:
        raise ValueError(
            "XMLA Execute response has no axes"
        )

    return {
        "valid": True,
        "axes": len(axes),
    }


# ============================================================
# WEBFORM AUTHORIZATION DISCOVERY
# ============================================================

AUTH_PATTERN = re.compile(
    r"""
    ['"]?Authorization['"]?
    \s*:\s*
    ['"](?P<token>[A-Fa-f0-9]{32,128})['"]
    """,
    re.IGNORECASE | re.VERBOSE,
)

METHOD_PATTERN = re.compile(
    r"""
    DynamicPivotCube\.aspx/
    (?P<method>
        GetStructure
        |GET_MENUTEMPLATE_LIST
        |BindChangeTemplate
    )
    """,
    re.IGNORECASE | re.VERBOSE,
)

SCRIPT_PATTERN = re.compile(
    r"""<script\b[^>]*\bsrc\s*=\s*["']([^"']+)["']""",
    re.IGNORECASE,
)


def extract_method_authorizations(
    source: str,
) -> dict[str, list[str]]:
    result = {
        method: []
        for method in WEBFORM_METHODS
    }

    lookup = {
        method.upper(): method
        for method in WEBFORM_METHODS
    }

    matches = list(
        METHOD_PATTERN.finditer(source)
    )

    for index, match in enumerate(matches):
        method = lookup.get(
            match.group("method").upper()
        )

        if not method:
            continue

        next_start = (
            matches[index + 1].start()
            if index + 1 < len(matches)
            else len(source)
        )

        end = min(
            match.end() + 2500,
            next_start,
        )

        fragment = source[
            match.end():end
        ]

        for found in AUTH_PATTERN.finditer(fragment):
            token = found.group("token")

            if token not in result[method]:
                result[method].append(token)

    return result


def decode_webform_response(
    response: requests.Response,
    method: str,
) -> list:
    if response.status_code != 200:
        raise ValueError(
            f"{method}: HTTP {response.status_code}"
        )

    try:
        data = response.json()

    except ValueError as exc:
        raise ValueError(
            f"{method}: invalid JSON response"
        ) from exc

    if (
        not isinstance(data, dict)
        or "d" not in data
    ):
        raise ValueError(
            f"{method}: missing ASP.NET 'd' property"
        )

    payload = data["d"]

    for _ in range(3):
        if not isinstance(payload, str):
            break

        try:
            payload = json.loads(payload)

        except (ValueError, TypeError) as exc:
            raise ValueError(
                f"{method}: invalid serialized JSON"
            ) from exc

    if not isinstance(payload, list):
        raise ValueError(
            f"{method}: expected array, got "
            f"{type(payload).__name__}"
        )

    return payload


# ============================================================
# OIC AUTHENTICATED CLIENT
# ============================================================

class OICClient:

    def __init__(
        self,
        cookie: str,
        token: str,
        timeout: int = 30,
    ):
        self.cookie = normalize_cookie(cookie)
        self.token = normalize_token(token)
        self.timeout = timeout

        if not self.cookie:
            raise ValueError(
                "OIC_COOKIE is missing from .env"
            )

        self.session = requests.Session()

        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Accept-Language": "th-TH,th;q=0.9",
        })

        hostname = urlparse(
            BASE_URL
        ).hostname

        for part in self.cookie.split(";"):
            name, separator, value = (
                part.strip().partition("=")
            )

            if not separator or not name.strip():
                continue

            self.session.cookies.set(
                name.strip(),
                value.strip(),
                domain=hostname,
                path="/",
                secure=True,
            )

        self.webform_auth = {}

    def close(self) -> None:
        self.session.close()

    def referer(self, code: str) -> str:
        return (
            WEBFORM_URL
            + "?Class="
            + CATEGORIES[code]["web_class"]
        )

    def refresh_token(self, code: str) -> None:
        if getattr(self,"shared_auth",False):
            from oic_all_dim import ensure_shared_auth
            cookie,token=ensure_shared_auth(code,getattr(self,"auth_stop",lambda:False),validate=True,allow_login=False,rejected_token=self.token)
            if cookie!=self.cookie:
                self.session.cookies.clear()
                self.webform_auth.clear()
                for part in cookie.split(";"):
                    name,sep,value=part.strip().partition("=")
                    if sep and name:
                        self.session.cookies.set(name.strip(),value.strip(),domain=urlparse(BASE_URL).hostname,path="/",secure=True)
            self.cookie,self.token=cookie,token
            return
        page_url = (
            WEBFORM_PATH
            + "?Class="
            + CATEGORIES[code]["web_class"]
        )

        response = self.session.post(
            TOKEN_URL,
            json={
                "_tk": "Init",
                "_pageUrl": page_url,
            },
            headers={
                "Content-Type": (
                    "application/json; charset=UTF-8"
                ),
                "X-Requested-With": "XMLHttpRequest",
                "Origin": BASE_URL,
                "Referer": self.referer(code),
            },
            timeout=(10, self.timeout),
            allow_redirects=False,
        )

        if response.status_code != 200:
            raise ValueError(
                "GenToken HTTP "
                f"{response.status_code}"
            )

        try:
            payload = response.json()

        except ValueError as exc:
            raise ValueError(
                "GenToken returned invalid JSON"
            ) from exc

        if isinstance(payload, dict):
            payload = payload.get("d")

        if not isinstance(payload, str):
            raise ValueError(
                "GenToken returned unexpected data"
            )

        if payload == "SESSION_EXPIRED":
            raise ValueError(
                "OIC_COOKIE session expired"
            )

        token = normalize_token(payload)

        if (
            token.startswith('"')
            and token.endswith('"')
        ):
            token = token[1:-1]

        if len(token.split(".")) != 3:
            raise ValueError(
                "GenToken did not return a JWT"
            )

        expiry = jwt_exp(token)

        if expiry and expiry <= time.time() + 30:
            raise ValueError(
                "GenToken returned expired JWT"
            )

        self.token = token
        self.cookie="; ".join(f"{c.name}={c.value}" for c in self.session.cookies.jar)

        log(
            f"{code}: JWT refreshed successfully",
            "OK",
        )

    def bearer(self, code: str) -> str:
        expires = jwt_exp(self.token)

        if (
            not self.token
            or (
                expires
                and expires <= time.time() + 60
            )
        ):
            self.refresh_token(code)

        return self.token

    def xmla(
        self,
        code: str,
        body: bytes,
    ) -> bytes:
        last_error = None

        for attempt in range(3):
            token = self.bearer(code)

            try:
                response = self.session.post(
                    XMLA_URL,
                    data=body,
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "text/xml",
                        "Accept": (
                            "text/plain, */*; q=0.01"
                        ),
                        "X-Requested-With": "XMLHttpRequest",
                        "Origin": BASE_URL,
                        "Referer": self.referer(code),
                    },
                    timeout=(10, self.timeout),
                    allow_redirects=False,
                )

            except RequestException as exc:
                last_error = exc

                if attempt < 2:
                    time.sleep(attempt + 1)
                    continue

                raise

            if response.status_code == 200:
                from oic_all_dim import extract_error_details,CollectionError
                if extract_error_details(response.content)=="AUTH_COOKIE_EXPIRED":
                    raise CollectionError("AUTH_COOKIE_EXPIRED","XMLA session expired")
                return response.content

            if response.status_code in (401, 403):
                if attempt == 0:
                    self.refresh_token(code)
                    continue

                raise ValueError(
                    "XMLA authentication rejected: "
                    f"HTTP {response.status_code}"
                )

            if response.status_code in (
                408, 429, 500, 502, 503, 504
            ):
                if attempt < 2:
                    retry_after = response.headers.get(
                        "Retry-After"
                    )

                    try:
                        delay = min(
                            float(retry_after),
                            30.0,
                        )

                    except (TypeError, ValueError):
                        delay = 2 ** attempt

                    time.sleep(
                        max(delay, 0.5)
                    )
                    continue

            raise ValueError(
                f"XMLA HTTP {response.status_code}"
            )

        raise ValueError(
            "XMLA request failed: "
            f"{type(last_error).__name__}"
        )

    def discover(
        self,
        code: str,
        request_type: str,
        cube: str | None = None,
    ) -> list[dict]:
        catalog = CATEGORIES[code]["catalog"]

        body = make_discover_request(
            catalog,
            request_type,
            cube,
        )

        response = self.xmla(
            code,
            body,
        )

        return parse_discover_rows(
            response
        )

    def execute(
        self,
        code: str,
        cube: str,
        mdx: str,
    ) -> dict:
        body = make_execute_request(
            CATEGORIES[code]["catalog"],
            mdx,
        )

        response = self.xmla(
            code,
            body,
        )

        return validate_execute_response(
            response
        )

    def load_webform_auth(
        self,
        code: str,
    ) -> dict[str, list[str]]:
        if code in self.webform_auth:
            return self.webform_auth[code]

        referer = self.referer(code)

        response = self.session.get(
            referer,
            headers={
                "Accept": (
                    "text/html,application/xhtml+xml,"
                    "application/xml;q=0.9,*/*;q=0.8"
                ),
                "Referer": BASE_URL + "/",
            },
            timeout=(10, self.timeout),
            allow_redirects=False,
        )

        if response.status_code != 200:
            raise ValueError(
                f"{code}: WebForm HTTP "
                f"{response.status_code}"
            )

        html = response.text

        if not html.strip():
            raise ValueError(
                f"{code}: empty WebForm page"
            )

        auth = extract_method_authorizations(
            html
        )

        urls = []
        seen = set()

        for match in SCRIPT_PATTERN.finditer(html):
            url = urljoin(
                referer,
                match.group(1),
            )

            parsed = urlparse(url)

            if parsed.scheme != "https":
                continue

            if parsed.hostname != urlparse(
                BASE_URL
            ).hostname:
                continue

            if url in seen:
                continue

            seen.add(url)
            urls.append(url)

        if not all(auth.values()):
            for url in urls[:20]:
                if all(auth.values()):
                    break

                try:
                    script = self.session.get(
                        url,
                        headers={
                            "Referer": referer,
                            "Accept": (
                                "application/javascript,"
                                "text/javascript,*/*"
                            ),
                        },
                        timeout=(10, self.timeout),
                        allow_redirects=False,
                    )

                    if script.status_code != 200:
                        continue

                    discovered = (
                        extract_method_authorizations(
                            script.text
                        )
                    )

                    for method, tokens in discovered.items():
                        for token in tokens:
                            if token not in auth[method]:
                                auth[method].append(token)

                except RequestException:
                    continue

        self.webform_auth[code] = auth

        return auth

    def webform(
        self,
        code: str,
        method: str,
        payload: dict,
    ) -> list:
        if method not in WEBFORM_METHODS:
            raise ValueError(
                f"Unknown WebForm method: {method}"
            )

        authorizations = self.load_webform_auth(
            code
        ).get(method, [])

        if not authorizations:
            raise ValueError(
                f"{code}: {method} authorization not found"
            )

        last_error = None

        for authorization in authorizations:
            response = self.session.post(
                f"{WEBFORM_URL}/{method}",
                json=payload,
                headers={
                    "Authorization": authorization,
                    "Content-Type": (
                        "application/json; charset=UTF-8"
                    ),
                    "Accept": (
                        "application/json, text/javascript,"
                        "*/*; q=0.01"
                    ),
                    "X-Requested-With": "XMLHttpRequest",
                    "Origin": BASE_URL,
                    "Referer": self.referer(code),
                },
                timeout=(10, self.timeout),
                allow_redirects=False,
            )

            try:
                return decode_webform_response(
                    response,
                    method,
                )

            except ValueError as exc:
                last_error = exc

                if response.status_code not in (
                    200, 401, 403
                ):
                    raise

        raise ValueError(
            f"{code}: {method} failed: {last_error}"
        )


# ============================================================
# METADATA NORMALIZATION
# ============================================================

def parse_template_items(
    code: str,
    cube: str,
    records: list,
) -> list[dict]:
    profile = CATEGORIES[code]
    items = []

    for record in records:
        if not isinstance(record, dict):
            continue

        record_class = record.get("CLASS")

        if record_class and record_class not in (
            code,
            profile["web_class"],
        ):
            continue

        record_cube = record.get("CUBE_NAME")

        if record_cube and record_cube != cube:
            continue

        raw = (
            record.get("TEMPLATE_DETAIL")
            or record.get("JSON")
        )

        if not raw:
            continue

        try:
            decoded = (
                json.loads(raw)
                if isinstance(raw, str)
                else raw
            )

        except (ValueError, TypeError):
            continue

        if not isinstance(decoded, list):
            continue

        for item in decoded:
            if not isinstance(item, dict):
                continue

            mdx = clean(
                item.get("dataField")
            )

            if not mdx:
                continue

            items.append({
                "mdx": mdx,
                "caption": clean(
                    item.get("caption")
                ),
                "folder": clean(
                    item.get("displayFolder")
                ),
                "area": clean(
                    item.get("area")
                ).lower(),
            })

    return items


def normalize_xmla_fields(
    hierarchies: list[dict],
    measures: list[dict],
) -> list[dict]:
    fields = []
    seen = set()

    for field_type, records in (
        ("DIMENSION", hierarchies),
        ("MEASURE", measures),
    ):
        for record in records:
            if field_type == "DIMENSION":
                mdx = clean(
                    record.get(
                        "HIERARCHY_UNIQUE_NAME"
                    )
                )

                caption = clean(
                    record.get(
                        "HIERARCHY_CAPTION"
                    )
                )

                folder = clean(
                    record.get(
                        "HIERARCHY_DISPLAY_FOLDER"
                    )
                )

            else:
                mdx = clean(
                    record.get(
                        "MEASURE_UNIQUE_NAME"
                    )
                )

                caption = clean(
                    record.get(
                        "MEASURE_CAPTION"
                    )
                )

                folder = clean(
                    record.get(
                        "MEASURE_DISPLAY_FOLDER"
                    )
                )

            if not valid_identifier(
                mdx,
                field_type,
            ):
                continue

            identity = (
                field_type,
                mdx.casefold(),
            )

            if identity in seen:
                continue

            seen.add(identity)

            fields.append({
                "type": field_type,
                "mdx": mdx,
                "caption": caption,
                "folder": folder,
                "metadata": record,
            })

    return fields


def build_xmla_indexes(
    fields: list[dict],
) -> dict:
    by_mdx = {}
    by_caption = defaultdict(set)
    by_caption_folder = defaultdict(set)

    for field in fields:
        field_type = field["type"]
        mdx = field["mdx"]

        by_mdx[
            (field_type, mdx.casefold())
        ] = mdx

        by_caption[
            (
                field_type,
                normalized(field["caption"]),
            )
        ].add(mdx)

        by_caption_folder[
            (
                field_type,
                normalized(field["caption"]),
                normalized(field["folder"]),
            )
        ].add(mdx)

    return {
        "by_mdx": by_mdx,
        "by_caption": by_caption,
        "by_caption_folder": by_caption_folder,
    }


def find_mapping(
    row: dict,
    field_type: str,
    xmla_index: dict,
    template_items: list[dict],
) -> tuple[str | None, str, str]:
    caption = clean(
        row.get("CAPTION")
    )

    folder = clean(
        row.get("DIMENSION")
    )

    table_name = clean(
        row.get("DATATABLE")
    ).replace("_", " ")

    data_field = clean(
        row.get("DATAFIELD")
    ).replace("_", " ")

    by_mdx = xmla_index["by_mdx"]

    def confirmed(candidates):
        return {
            by_mdx[
                (field_type, item.casefold())
            ]
            for item in candidates
            if (
                field_type,
                item.casefold(),
            ) in by_mdx
        }

    candidate_templates = set()

    for template in template_items:
        if (
            normalized(template["caption"])
            != normalized(caption)
        ):
            continue

        expected_areas = (
            {"data"}
            if field_type == "MEASURE"
            else {"row", "column", "filter"}
        )

        if template["area"] not in expected_areas:
            continue

        if (
            folder
            and template["folder"]
            and (
                normalized(template["folder"])
                != normalized(folder)
            )
        ):
            continue

        candidate_templates.add(
            template["mdx"]
        )

    matches = confirmed(
        candidate_templates
    )

    if len(matches) == 1:
        return (
            next(iter(matches)),
            "OIC_TEMPLATE",
            "Exact mapping confirmed in XMLA metadata",
        )

    if len(matches) > 1:
        return (
            None,
            "AMBIGUOUS",
            "Multiple OIC template mappings",
        )

    direct = set()

    if field_type == "MEASURE":
        direct.add(
            f"[Measures].[{data_field}]"
        )

    elif table_name and data_field:
        direct.add(
            f"[{table_name}].[{data_field}]"
        )

    matches = confirmed(direct)

    if len(matches) == 1:
        return (
            next(iter(matches)),
            "XMLA_IDENTIFIER",
            "Exact source identifier found in XMLA",
        )

    if len(matches) > 1:
        return (
            None,
            "AMBIGUOUS",
            "Multiple matching source identifiers",
        )

    matches = xmla_index[
        "by_caption_folder"
    ].get(
        (
            field_type,
            normalized(caption),
            normalized(folder),
        ),
        set(),
    )

    if len(matches) == 1:
        return (
            next(iter(matches)),
            "XMLA_CAPTION_FOLDER",
            "Unique exact caption and folder",
        )

    if len(matches) > 1:
        return (
            None,
            "AMBIGUOUS",
            "Multiple matching caption/folder identifiers",
        )

    matches = xmla_index[
        "by_caption"
    ].get(
        (
            field_type,
            normalized(caption),
        ),
        set(),
    )

    if len(matches) == 1:
        return (
            next(iter(matches)),
            "XMLA_CAPTION",
            "Unique exact caption in XMLA metadata",
        )

    if len(matches) > 1:
        return (
            None,
            "AMBIGUOUS",
            "Caption maps to multiple XMLA identifiers",
        )

    return (
        None,
        "NOT_MAPPED",
        "No confirmed MDX identifier found",
    )


def build_field_catalog(
    code: str,
    cube_type: str,
    structure: list[dict],
    template_items: list[dict],
    xmla_fields: list[dict],
    levels: list[dict],
) -> list[dict]:
    indexes = build_xmla_indexes(
        xmla_fields
    )

    fields = []
    represented = set()

    for index, row in enumerate(structure):
        if not isinstance(row, dict):
            continue

        field_type = GROUP_TYPES.get(
            clean(row.get("GROUP_DIMENSION"))
        )

        if not field_type:
            continue

        caption = clean(
            row.get("CAPTION")
            or row.get("DATAFIELD")
            or "(unnamed)"
        )

        folder = clean(
            row.get("DIMENSION")
        )

        mdx, mapping_source, reason = find_mapping(
            row,
            field_type,
            indexes,
            template_items,
        )

        calculated = bool(
            clean(row.get("CALCULATE"))
        )

        summary_type = clean(
            row.get("SUMMARYTYPE")
        ).lower()

        derived = (
            calculated
            or summary_type not in ("", "sum")
        )

        if mdx:
            status = "MAPPED"

            represented.add(
                (
                    field_type,
                    mdx.casefold(),
                )
            )

        elif derived:
            status = "UNSUPPORTED"

        else:
            status = "UNVERIFIED"

        identity = {
            "group": row.get("GROUP_DIMENSION"),
            "caption": caption,
            "table": row.get("DATATABLE"),
            "field": row.get("DATAFIELD"),
            "seq": row.get("SEQ", index),
        }

        fields.append({
            "id": make_field_id(
                code,
                cube_type,
                "GETSTRUCTURE",
                identity,
            ),
            "type": field_type,
            "caption": caption,
            "folder": folder,
            "data_table": clean(
                row.get("DATATABLE")
            ),
            "data_field": clean(
                row.get("DATAFIELD")
            ),
            "mdx": mdx,
            "status": status,
            "source": "GETSTRUCTURE",
            "mapping_source": mapping_source,
            "mapping_reason": reason,
            "verification_method": None,
            "verified_at": None,
            "error": None,
            "details": row,
        })

    # Include fields returned directly by XMLA Discover.

    for item in xmla_fields:
        identity = (
            item["type"],
            item["mdx"].casefold(),
        )

        if identity in represented:
            continue

        represented.add(identity)

        fields.append({
            "id": make_field_id(
                code,
                cube_type,
                "XMLA_DISCOVER",
                {
                    "type": item["type"],
                    "mdx": item["mdx"],
                },
            ),
            "type": item["type"],
            "caption": (
                item["caption"]
                or item["mdx"]
            ),
            "folder": item["folder"],
            "data_table": "",
            "data_field": "",
            "mdx": item["mdx"],
            "status": "MAPPED",
            "source": "XMLA_DISCOVER",
            "mapping_source": "XMLA_UNIQUE_NAME",
            "mapping_reason": (
                "Identifier returned by XMLA Discover"
            ),
            "verification_method": None,
            "verified_at": None,
            "error": None,
            "details": item["metadata"],
        })

    level_lookup = defaultdict(list)

    for row in levels:
        hierarchy = clean(
            row.get("HIERARCHY_UNIQUE_NAME")
        )

        if hierarchy:
            level_lookup[
                hierarchy.casefold()
            ].append({
                "unique_name": row.get(
                    "LEVEL_UNIQUE_NAME"
                ),
                "caption": row.get(
                    "LEVEL_CAPTION"
                ),
                "number": row.get(
                    "LEVEL_NUMBER"
                ),
            })

    for field in fields:
        if field["type"] == "DIMENSION":
            field["levels"] = level_lookup.get(
                clean(field["mdx"]).casefold(),
                [],
            )
        else:
            field["levels"] = []

    return fields


# ============================================================
# FIELD EXECUTE VERIFICATION
# ============================================================

def make_field_mdx(
    cube: str,
    field: dict,
    anchor_measure: str | None,
) -> str:
    mdx = field.get("mdx")

    if not mdx:
        raise ValueError(
            "Field has no confirmed MDX mapping"
        )

    if not valid_identifier(
        mdx,
        field["type"],
    ):
        raise ValueError(
            "Field identifier is invalid"
        )

    if field["type"] == "MEASURE":
        return (
            f"SELECT {{{mdx}}} ON COLUMNS "
            f"FROM [{cube}] "
            "CELL PROPERTIES VALUE"
        )

    if not anchor_measure:
        raise ValueError(
            "Cannot verify hierarchy: "
            "no discovered Measure available"
        )

    return (
        f"SELECT {{{anchor_measure}}} ON COLUMNS, "
        f"Subset({mdx}.AllMembers, 0, 1) ON ROWS "
        f"FROM [{cube}] "
        "CELL PROPERTIES VALUE"
    )


def verify_one_field(
    client: OICClient,
    code: str,
    cube: str,
    field: dict,
    anchor_measure: str | None,
) -> dict:
    if not field.get("mdx"):
        return field

    mdx = make_field_mdx(
        cube,
        field,
        anchor_measure,
    )

    try:
        result = client.execute(
            code,
            cube,
            mdx,
        )

        if not result.get("valid"):
            raise ValueError(
                "XMLA Execute did not validate"
            )

        field["status"] = "VERIFIED"
        field["verified_at"] = utc_now()
        field["verification_method"] = (
            "XMLA_EXECUTE"
        )
        field["mapping_reason"] = (
            "Confirmed MDX mapping and "
            "successful XMLA Execute"
        )
        field["error"] = None

    except (
        RequestException,
        ValueError,
        ET.ParseError,
    ) as exc:
        from oic_all_dim import is_session_error
        if is_session_error(exc) or any(x in str(exc).upper() for x in ("HTTP 401","HTTP 403","AUTHENTICATION REJECTED")):
            raise
        field["status"] = "MAPPED"
        field["verified_at"] = None
        field["verification_method"] = (
            "XMLA_EXECUTE_FAILED"
        )
        field["error"] = safe_error(
            exc,
            (
                client.cookie,
                client.token,
            ),
        )

    return field


# ============================================================
# VERIFICATION REPORT
# ============================================================

def count_fields(
    fields: list[dict],
) -> dict:
    counts = {
        "TOTAL": len(fields),
        "VERIFIED": 0,
        "MAPPED": 0,
        "UNVERIFIED": 0,
        "UNSUPPORTED": 0,
    }

    for field in fields:
        status = field.get("status")

        if status in counts:
            counts[status] += 1

    return counts


def create_cube_report(
    code: str,
    cube_type: str,
) -> dict:
    profile = CATEGORIES[code]

    return {
        "insurance_code": code,
        "insurance_type": profile["name"],
        "cube_type": cube_type,
        "cube_type_th": CUBE_TYPES[cube_type],
        "catalog": profile["catalog"],
        "cube": f"{code}_AGGREGATE_{cube_type}",
        "status": "PENDING",
        "started_at": utc_now(),
        "updated_at": utc_now(),
        "completed_at": None,
        "metadata_sources": [],
        "errors": [],
        "counts": count_fields([]),
        "fields": [],
    }


def persist_report(
    report: dict,
) -> Path:
    report["updated_at"] = utc_now()

    report["counts"] = count_fields(
        report.get("fields", [])
    )

    return atomic_json(
        report_path(
            report["insurance_code"],
            report["cube_type"],
        ),
        report,
    )


def verify_saved_catalog(client,code,kind,catalog,stop=lambda:False,progress=None):
    """Verify only the current website fields; checkpoint without rediscovery."""
    import oic_query as query
    report=create_cube_report(code,kind)
    report.update(status="AVAILABLE",catalog=catalog["catalog"],cube=catalog["cube"],
                  fields=[dict(f) for f in catalog["fields"] if f.get("source")=="GETSTRUCTURE"])
    fields=report["fields"]
    anchor=next((f["mdx"] for f in fields if f["type"]=="MEASURE"
                 and f.get("mdx") and f["status"] in ("MAPPED","VERIFIED")),None)
    evidence={}
    for field in fields:
        if (field.get("status")=="VERIFIED" and field.get("verified_at")
                and field.get("verification_method")=="XMLA_EXECUTE"):
            evidence[(field["type"],field.get("mdx"))]=field
    try:
        for index,field in enumerate(fields,1):
            if stop():
                raise RuntimeError("Verification stopped")
            key=(field["type"],field.get("mdx"))
            if field.get("status") in ("MAPPED","VERIFIED") and field.get("mdx"):
                if key in evidence:
                    for name in ("status","verified_at","verification_method","mapping_reason"):
                        field[name]=evidence[key].get(name)
                elif field["type"]=="MEASURE" or anchor:
                    verify_one_field(client,code,report["cube"],field,anchor)
                    error=str(field.get("error") or "").upper()
                    if any(marker in error for marker in ("SESSION_EXPIRED","SESSION EXPIRED","HTTP 401","HTTP 403","AUTHENTICATION REJECTED")):
                        raise RuntimeError("Session expired or authentication rejected during verification")
                    if field["status"]=="VERIFIED":
                        evidence[key]=field
                field["mapping_status"]=field["status"]
            if index%25==0:
                persist_report(report)
            if progress:
                progress(index,len(fields))
    finally:
        persist_report(report)
    report["completed_at"]=utc_now()
    persist_report(report)
    # The verification loader merges matching evidence, never old membership.
    current=query.load_web_catalog(code,kind)
    return current


def restore_verified_fields(
    current: list[dict],
    previous: dict,
    *,
    force: bool = False,
) -> int:
    if force:
        return 0

    old_fields = {
        field.get("id"): field
        for field in previous.get("fields", [])
        if (
            isinstance(field, dict)
            and field.get("status") == "VERIFIED"
            and field.get("verified_at")
            and field.get("verification_method")
            == "XMLA_EXECUTE"
        )
    }

    restored = 0

    for field in current:
        previous_field = old_fields.get(
            field["id"]
        )

        if not previous_field:
            continue

        if (
            previous_field.get("mdx")
            != field.get("mdx")
        ):
            continue

        if (
            previous_field.get("type")
            != field.get("type")
        ):
            continue

        field["status"] = "VERIFIED"
        field["verified_at"] = (
            previous_field["verified_at"]
        )
        field["verification_method"] = (
            "XMLA_EXECUTE"
        )
        field["mapping_reason"] = (
            "Previously verified with XMLA Execute"
        )
        field["error"] = None

        restored += 1

    return restored


# ============================================================
# FULL CUBE VERIFICATION
# ============================================================

def verify_cube(
    client: OICClient,
    code: str,
    cube_type: str,
    args: argparse.Namespace,
    known_cubes: set[str] | None = None,
) -> dict:
    previous = load_cube_report(
        code,
        cube_type,
    ) or {}

    report = create_cube_report(
        code,
        cube_type,
    )

    cube = report["cube"]

    log(
        f"DISCOVER {code}/{cube_type} | {cube}"
    )

    structure = []
    menu_records = []
    binding_records = []

    hierarchy_records = []
    measure_records = []
    level_records = []

    getstructure_ok = False
    hierarchy_ok = False
    measures_ok = False

    # --------------------------------------------------------
    # 1. GETSTRUCTURE
    # --------------------------------------------------------

    try:
        structure = client.webform(
            code,
            "GetStructure",
            {
                "Template": [
                    code,
                    cube,
                ]
            },
        )

        getstructure_ok = True

        report["metadata_sources"].append(
            "GetStructure"
        )

        log(
            f"{code}/{cube_type}: "
            f"GetStructure={len(structure)} fields"
        )

    except Exception as exc:
        report["errors"].append(
            "GetStructure: "
            + safe_error(
                exc,
                (
                    client.cookie,
                    client.token,
                ),
            )
        )

    # --------------------------------------------------------
    # 2. TEMPLATE METADATA
    # --------------------------------------------------------

    try:
        menu_records = client.webform(
            code,
            "GET_MENUTEMPLATE_LIST",
            {
                "Template": [
                    CATEGORIES[code]["web_class"],
                ]
            },
        )

        report["metadata_sources"].append(
            "GET_MENUTEMPLATE_LIST"
        )

    except Exception as exc:
        report["errors"].append(
            "GET_MENUTEMPLATE_LIST: "
            + safe_error(
                exc,
                (
                    client.cookie,
                    client.token,
                ),
            )
        )

    try:
        binding_records = client.webform(
            code,
            "BindChangeTemplate",
            {
                "Template": [
                    code,
                    cube,
                ]
            },
        )

        report["metadata_sources"].append(
            "BindChangeTemplate"
        )

    except Exception as exc:
        report["errors"].append(
            "BindChangeTemplate: "
            + safe_error(
                exc,
                (
                    client.cookie,
                    client.token,
                ),
            )
        )

    # --------------------------------------------------------
    # 3. XMLA DISCOVER
    # --------------------------------------------------------

    targets = (
        ("MDSCHEMA_HIERARCHIES", "hierarchies"),
        ("MDSCHEMA_MEASURES", "measures"),
        ("MDSCHEMA_LEVELS", "levels"),
    )

    for request_type, name in targets:
        try:
            records = client.discover(
                code,
                request_type,
                cube,
            )

            if name == "hierarchies":
                hierarchy_records = records
                hierarchy_ok = True

            elif name == "measures":
                measure_records = records
                measures_ok = True

            else:
                level_records = records

            report["metadata_sources"].append(
                request_type
            )

            log(
                f"{code}/{cube_type}: "
                f"{name}={len(records)}"
            )

        except Exception as exc:
            report["errors"].append(
                request_type + ": "
                + safe_error(
                    exc,
                    (
                        client.cookie,
                        client.token,
                    ),
                )
            )

    # --------------------------------------------------------
    # 4. CONFIRM CUBE STATUS
    # --------------------------------------------------------

    xmla_fields = normalize_xmla_fields(
        hierarchy_records,
        measure_records,
    )

    has_live_metadata = bool(
        structure or xmla_fields
    )

    if not has_live_metadata:
        confirmed_missing = (
            known_cubes is not None
            and cube.upper() not in known_cubes
            and getstructure_ok
            and hierarchy_ok
            and measures_ok
        )

        report["status"] = (
            "NOT_AVAILABLE"
            if confirmed_missing
            else "UNKNOWN"
        )

        report["completed_at"] = utc_now()

        # Preserve older successful data on temporary
        # metadata failures.

        if (
            report["status"] == "UNKNOWN"
            and previous.get("fields")
        ):
            previous.setdefault(
                "errors",
                [],
            ).extend(report["errors"])

            previous["status"] = "ERROR"

            persist_report(previous)

            return previous

        persist_report(report)

        return report

    # --------------------------------------------------------
    # 5. BUILD FIELD CATALOG
    # --------------------------------------------------------

    template_items = parse_template_items(
        code,
        cube,
        menu_records + binding_records,
    )

    fields = build_field_catalog(
        code,
        cube_type,
        structure,
        template_items,
        xmla_fields,
        level_records,
    )

    report["fields"] = fields
    report["status"] = "AVAILABLE"

    restored = restore_verified_fields(
        fields,
        previous,
        force=args.force,
    )

    log(
        f"{code}/{cube_type}: "
        f"RESTORED VERIFIED={restored}",
        "INFO",
    )

    # Initial checkpoint before network verification.
    persist_report(report)

    # --------------------------------------------------------
    # 6. ANCHOR MEASURE
    # --------------------------------------------------------

    anchor_measure = next(
        (
            item["mdx"]
            for item in xmla_fields
            if (
                item["type"] == "MEASURE"
                and valid_identifier(
                    item["mdx"],
                    "MEASURE",
                )
            )
        ),
        None,
    )

    # --------------------------------------------------------
    # 7. PENDING FIELDS
    # --------------------------------------------------------

    pending = [
        field
        for field in fields
        if (
            field.get("mdx")
            and (
                args.force
                or field.get("status") != "VERIFIED"
            )
        )
    ]

    log(
        f"{code}/{cube_type}: "
        f"TOTAL={len(fields)} | "
        f"PENDING={len(pending)}"
    )

    checkpoint_every = max(
        1,
        args.checkpoint_every,
    )

    processed = 0
    last_checkpoint = 0

    def checkpoint(
        *,
        force: bool = False,
    ) -> None:
        nonlocal last_checkpoint

        if not force:
            if (
                processed - last_checkpoint
                < checkpoint_every
            ):
                return

        persist_report(report)

        last_checkpoint = processed

    def process_field(field: dict) -> None:
        nonlocal processed

        verify_one_field(
            client,
            code,
            cube,
            field,
            anchor_measure,
        )

        processed += 1

        if field["status"] == "VERIFIED":
            log(
                f"{code}/{cube_type}: "
                f"VERIFIED {field['caption']}",
                "OK",
            )

        else:
            log(
                f"{code}/{cube_type}: "
                f"{field['status']} "
                f"{field['caption']}",
                "WARN",
            )

        checkpoint()

        if args.delay > 0:
            time.sleep(args.delay)

    # --------------------------------------------------------
    # 8. VERIFY WITH CHECKPOINTS
    # --------------------------------------------------------

    try:
        if RICH_AVAILABLE and pending:
            with Progress(
                SpinnerColumn(),
                TextColumn(
                    "[cyan]{task.description}"
                ),
                BarColumn(),
                TaskProgressColumn(),
                TimeElapsedColumn(),
                console=console,
                transient=True,
            ) as progress:
                task = progress.add_task(
                    f"VERIFY {code}/{cube_type}",
                    total=len(pending),
                )

                for field in pending:
                    process_field(field)

                    progress.advance(
                        task,
                        1,
                    )

        else:
            for field in pending:
                process_field(field)

    except KeyboardInterrupt:
        report["status"] = "INTERRUPTED"
        report["completed_at"] = None

        checkpoint(force=True)

        raise

    except Exception:
        report["status"] = "ERROR"
        report["completed_at"] = None

        checkpoint(force=True)

        raise

    # --------------------------------------------------------
    # 9. FINAL CHECKPOINT
    # --------------------------------------------------------

    report["status"] = "AVAILABLE"
    report["completed_at"] = utc_now()

    checkpoint(force=True)

    return report


# ============================================================
# CSV EXPORT
# ============================================================

def export_all_fields(
    reports: list[dict],
) -> tuple[int, Path]:
    destination = (
        OUTPUT_DIR / "all_fields.csv"
    )

    destination.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary = destination.with_name(
        destination.name
        + "."
        + uuid.uuid4().hex
        + ".tmp"
    )

    total = 0

    try:
        with temporary.open(
            "w",
            encoding="utf-8-sig",
            newline="",
        ) as stream:
            writer = csv.DictWriter(
                stream,
                fieldnames=CSV_COLUMNS,
                extrasaction="ignore",
            )

            writer.writeheader()

            for report in reports:
                for field in report.get(
                    "fields",
                    [],
                ):
                    row = {
                        "insurance_code": report[
                            "insurance_code"
                        ],
                        "insurance_type": report[
                            "insurance_type"
                        ],
                        "cube_type": report[
                            "cube_type"
                        ],
                        "cube_type_th": report[
                            "cube_type_th"
                        ],
                        "catalog": report[
                            "catalog"
                        ],
                        "cube": report[
                            "cube"
                        ],
                        "field_id": field.get("id"),
                        "field_type": field.get("type"),
                        "field_caption": field.get(
                            "caption"
                        ),
                        "field_folder": field.get(
                            "folder"
                        ),
                        "data_table": field.get(
                            "data_table"
                        ),
                        "data_field": field.get(
                            "data_field"
                        ),
                        "mdx": field.get("mdx"),
                        "status": field.get("status"),
                        "source": field.get("source"),
                        "mapping_source": field.get(
                            "mapping_source"
                        ),
                        "mapping_reason": field.get(
                            "mapping_reason"
                        ),
                        "verification_method": field.get(
                            "verification_method"
                        ),
                        "verified_at": field.get(
                            "verified_at"
                        ),
                        "error": field.get("error"),
                    }

                    # Prevent CSV formula injection.

                    for key, value in row.items():
                        if (
                            isinstance(value, str)
                            and value.startswith(
                                ("=", "+", "-", "@")
                            )
                        ):
                            row[key] = "'" + value

                    writer.writerow(row)

                    total += 1

            stream.flush()
            os.fsync(stream.fileno())

        output_path = commit_file(
            destination,
            temporary,
        )

        return total, output_path

    finally:
        try:
            temporary.unlink(
                missing_ok=True
            )
        except OSError:
            pass


# ============================================================
# SUMMARY GENERATOR
# ============================================================

def load_all_reports() -> list[dict]:
    reports = []

    for code in CATEGORIES:
        for cube_type in CUBE_TYPES:
            report = load_cube_report(
                code,
                cube_type,
            )

            if not isinstance(report, dict):
                continue

            reports.append(report)

    return reports


def build_summary(
    reports: list[dict],
) -> dict:
    cube_records = []

    totals = {
        "TOTAL": 0,
        "VERIFIED": 0,
        "MAPPED": 0,
        "UNVERIFIED": 0,
        "UNSUPPORTED": 0,
    }

    cube_statuses = defaultdict(int)

    for report in reports:
        counts = count_fields(
            report.get("fields", [])
        )

        cube_records.append({
            "insurance_code": report[
                "insurance_code"
            ],
            "cube_type": report[
                "cube_type"
            ],
            "cube_type_th": report[
                "cube_type_th"
            ],
            "cube": report[
                "cube"
            ],
            "status": report[
                "status"
            ],
            "counts": counts,
            "errors": report.get(
                "errors",
                [],
            ),
        })

        cube_statuses[
            report["status"]
        ] += 1

        for key in totals:
            totals[key] += counts.get(
                key,
                0,
            )

    return {
        "generated_at": utc_now(),
        "total_cube_targets": (
            len(CATEGORIES) * len(CUBE_TYPES)
        ),
        "reported_cubes": len(reports),
        "cube_statuses": dict(cube_statuses),
        "totals": totals,
        "cubes": cube_records,
    }


def save_summary() -> tuple[dict, Path, Path]:
    reports = load_all_reports()

    summary = build_summary(
        reports
    )

    _, csv_path = export_all_fields(
        reports
    )

    json_path = atomic_json(
        OUTPUT_DIR / "summary.json",
        summary,
    )

    return summary, json_path, csv_path


# ============================================================
# COMMAND LINE
# ============================================================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="oic_verify_fields.py",
        description=(
            "Discover and verify OIC fields "
            "for AH/FI/MR/MI/MV and FY/UY/AY"
        ),
    )

    parser.add_argument(
        "--code",
        choices=list(CATEGORIES),
        type=str.upper,
        help="Insurance category",
    )

    parser.add_argument(
        "--cube-type",
        choices=list(CUBE_TYPES),
        type=str.upper,
        help="FY, UY or AY",
    )

    parser.add_argument(
        "--unverified-only",
        action="store_true",
        help=(
            "Reuse verified results and "
            "verify remaining mapped fields"
        ),
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-verify previously verified fields",
    )

    parser.add_argument(
        "--summary",
        action="store_true",
        help="Show saved verification summary only",
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="HTTP read timeout in seconds",
    )

    parser.add_argument(
        "--delay",
        type=float,
        default=0.3,
        help="Delay between field requests",
    )

    parser.add_argument(
        "--checkpoint-every",
        type=int,
        default=DEFAULT_CHECKPOINT_EVERY,
        help=(
            "Save verification progress every N fields "
            "(default: 10)"
        ),
    )

    return parser.parse_args()


# ============================================================
# ERROR RECOVERY
# ============================================================

def preserve_cube_error(
    code: str,
    cube_type: str,
    message: str,
) -> None:
    """
    Keep already-verified fields after unexpected errors.

    Only fatal Cube processing failures produce ERROR.
    A temporary Windows file lock is handled separately
    by commit_file() and recovery checkpoints.
    """

    previous = load_cube_report(
        code,
        cube_type,
    )

    if isinstance(previous, dict):
        report = previous

    else:
        report = create_cube_report(
            code,
            cube_type,
        )

    errors = report.setdefault(
        "errors",
        [],
    )

    if message not in errors:
        errors.append(message)

    report["status"] = "ERROR"
    report["completed_at"] = None

    persist_report(report)


# ============================================================
# APPLICATION
# ============================================================

def main() -> int:
    args = parse_args()

    print_header()

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    if args.summary:
        summary, json_path, csv_path = save_summary()

        print_summary(summary)

        log(f"REPORT: {json_path}")
        log(f"FIELDS: {csv_path}")

        return 0

    if args.timeout < 1:
        log(
            "--timeout must be greater than zero",
            "ERROR",
        )
        return 2

    if args.delay < 0:
        log(
            "--delay cannot be negative",
            "ERROR",
        )
        return 2

    if args.checkpoint_every < 1:
        log(
            "--checkpoint-every must be >= 1",
            "ERROR",
        )
        return 2

    credentials = read_env()

    cookie = normalize_cookie(
        credentials.get(
            "OIC_COOKIE",
            "",
        )
    )

    token = normalize_token(
        credentials.get(
            "OIC_BEARER_TOKEN",
            "",
        )
    )

    if not cookie:
        log(
            "OIC_COOKIE is missing from .env",
            "ERROR",
        )
        return 2

    selected_codes = (
        [args.code]
        if args.code
        else list(CATEGORIES)
    )

    selected_types = (
        [args.cube_type]
        if args.cube_type
        else list(CUBE_TYPES)
    )

    targets = [
        (code, cube_type)
        for code in selected_codes
        for cube_type in selected_types
    ]

    log(
        f"CATEGORIES={len(selected_codes)} | "
        f"CUBE TYPES={len(selected_types)} | "
        f"TARGETS={len(targets)} | "
        f"CHECKPOINT EVERY={args.checkpoint_every}"
    )

    client = OICClient(
        cookie=cookie,
        token=token,
        timeout=args.timeout,
    )

    failures = 0
    interrupted = False

    catalog_cache: dict[
        str,
        set[str] | None,
    ] = {}

    try:
        for code, cube_type in targets:

            # Discover the Cube list once per category.

            if code not in catalog_cache:
                try:
                    cubes = client.discover(
                        code,
                        "MDSCHEMA_CUBES",
                    )

                    catalog_cache[code] = {
                        clean(
                            item.get("CUBE_NAME")
                        ).upper()
                        for item in cubes
                        if item.get("CUBE_NAME")
                    }

                    log(
                        f"{code}: discovered "
                        f"{len(catalog_cache[code])} cubes"
                    )

                except Exception as exc:
                    catalog_cache[code] = None

                    log(
                        f"{code}: cube listing unavailable: "
                        + safe_error(
                            exc,
                            (
                                client.cookie,
                                client.token,
                            ),
                        ),
                        "WARN",
                    )

            try:
                result = verify_cube(
                    client,
                    code,
                    cube_type,
                    args,
                    known_cubes=catalog_cache[code],
                )

                print_cube_result(result)

                if result["status"] not in (
                    "AVAILABLE",
                    "NOT_AVAILABLE",
                ):
                    failures += 1

            except KeyboardInterrupt:
                interrupted = True
                break

            except Exception as exc:
                failures += 1

                error = safe_error(
                    exc,
                    (
                        client.cookie,
                        client.token,
                    ),
                )

                log(
                    f"{code}/{cube_type}: "
                    f"ERROR {error}",
                    "ERROR",
                )

                try:
                    preserve_cube_error(
                        code,
                        cube_type,
                        error,
                    )

                except OSError as save_exc:
                    log(
                        f"{code}/{cube_type}: "
                        f"COULD NOT SAVE ERROR STATUS: "
                        f"{type(save_exc).__name__}",
                        "ERROR",
                    )

    except KeyboardInterrupt:
        interrupted = True

    finally:
        client.close()

        try:
            summary, json_path, csv_path = save_summary()

            print_summary(summary)

            log(
                f"REPORT: {json_path}"
            )

            log(
                f"FIELDS: {csv_path}"
            )

        except OSError as exc:
            log(
                "FINAL SUMMARY SAVE FAILED: "
                + safe_error(exc),
                "ERROR",
            )

            failures += 1

    if interrupted:
        log(
            "Interrupted. Saved verification results "
            "can be resumed.",
            "WARN",
        )
        return 130

    if failures:
        log(
            f"Finished with {failures} "
            "Cube processing errors.",
            "WARN",
        )
        return 1

    log(
        "Field verification finished.",
        "OK",
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
