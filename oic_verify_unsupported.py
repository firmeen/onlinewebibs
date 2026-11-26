
#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC UNSUPPORTED FIELD VERIFIER
==============================
File: oic_verify_unsupported.py

Purpose:
    Resolve UNSUPPORTED fields from previous OIC verification reports.

    1. Discover native and calculated measures.
    2. Analyze GetStructure metadata.
    3. Analyze OIC templates.
    4. Resolve actual MDX identifiers.
    5. Resolve documented arithmetic formulas.
    6. Execute and validate each resolved field.
    7. Preserve the original verification reports.

Authentication:
    .env
        OIC_COOKIE=...
        OIC_BEARER_TOKEN=...

Installation:
    pip install requests rich

Usage:
    python oic_verify_unsupported.py
    python oic_verify_unsupported.py --code AH
    python oic_verify_unsupported.py --code MV --cube-type UY
    python oic_verify_unsupported.py --force
    python oic_verify_unsupported.py --summary

Input:
    output/field_verification/*.json

Output:
    output/unsupported_verification/
        summary.json
        resolved_fields.json
        unresolved_fields.json
        all_fields.csv
        checkpoint.json
        metadata/
            AH_UY.json
            MV_UY.json

Statuses:
    VERIFIED_NATIVE
    VERIFIED_CALCULATED
    UNVERIFIED
    UNSUPPORTED
    ERROR

Important:
    - Standalone script.
    - Does not import or modify other OIC Python files.
    - Does not overwrite original verification reports.
    - Only processes fields previously marked UNSUPPORTED.
    - Does not invent financial or actuarial formulas.
    - Only evaluates formulas supported by actual metadata.
    - Successful Execute confirms queryability, not mathematical
      equivalence with every calculation displayed by the OIC UI.
    - Live verification requires a valid OIC session.
    - This implementation has not been live-tested against OIC.
"""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import json
import os
import re
import sys
import time
import uuid

from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse
from xml.etree import ElementTree as ET

try:
    import requests
except ImportError:
    print("Missing dependency: requests")
    print("Install: python -m pip install requests rich")
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

SOURCE_DIR = (
    ROOT
    / "output"
    / "field_verification"
)

OUTPUT_DIR = (
    ROOT
    / "output"
    / "unsupported_verification"
)

METADATA_DIR = OUTPUT_DIR / "metadata"

CHECKPOINT_FILE = OUTPUT_DIR / "checkpoint.json"

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

SOAP_NS = (
    "http://schemas.xmlsoap.org/soap/envelope/"
)

XMLA_NS = (
    "urn:schemas-microsoft-com:xml-analysis"
)

MD_NS = (
    "urn:schemas-microsoft-com:"
    "xml-analysis:mddataset"
)

ROWSET_NS = (
    "urn:schemas-microsoft-com:"
    "xml-analysis:rowset"
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

RESOLVED_STATUSES = {
    "VERIFIED_NATIVE",
    "VERIFIED_CALCULATED",
}

STATUSES = (
    "VERIFIED_NATIVE",
    "VERIFIED_CALCULATED",
    "UNVERIFIED",
    "UNSUPPORTED",
    "ERROR",
)

FILE_RETRY_DELAYS = (
    0.10,
    0.25,
    0.50,
    0.75,
    1.00,
    1.50,
    2.00,
)

MEASURE_PATTERN = re.compile(
    r"\[Measures\]\.\[(?:[^\]\r\n]|\]\])+\]",
    re.IGNORECASE,
)

HIERARCHY_PATTERN = re.compile(
    r"^\[(?:[^\]\r\n]|\]\])+\]"
    r"\.\[(?:[^\]\r\n]|\]\])+\]$"
)

FORMULA_KEYS = {
    "CALCULATE",
    "CALCULATION",
    "FORMULA",
    "EXPRESSION",
    "MDXEXPRESSION",
    "MDXFORMULA",
    "CALCULATEDMEMBEREXPRESSION",
    "MEASUREEXPRESSION",
    "CALCULATEDFIELD",
}

IDENTIFIER_KEYS = {
    "MDX",
    "MDXNAME",
    "MDXIDENTIFIER",
    "UNIQUENAME",
    "MEASUREUNIQUENAME",
    "DATAFIELD",
    "MEASURENAME",
}

SENSITIVE_KEYS = {
    "COOKIE",
    "SETCOOKIE",
    "AUTHORIZATION",
    "TOKEN",
    "ACCESSTOKEN",
    "REFRESHTOKEN",
    "BEARERTOKEN",
    "PASSWORD",
    "SESSIONID",
    "OICCOOKIE",
    "OICBEARERTOKEN",
}

CSV_COLUMNS = [
    "insurance_code",
    "cube_type",
    "catalog",
    "cube",
    "field_id",
    "field_caption",
    "field_type",
    "data_table",
    "data_field",
    "summary_type",
    "original_status",
    "status",
    "resolved_mdx",
    "resolved_expression",
    "resolution_method",
    "mapping_source",
    "verification_context",
    "sample_cells",
    "nonempty_cells",
    "verified_at",
    "reason",
    "error",
]

if RICH_AVAILABLE:
    console = Console(highlight=False)
else:
    console = None


# ============================================================
# EXCEPTIONS
# ============================================================

class AuthenticationError(RuntimeError):
    pass


class XMLAError(RuntimeError):
    pass


class ResolutionError(RuntimeError):
    pass


# ============================================================
# TERMINAL
# ============================================================

def log(
    message: str,
    level: str = "INFO",
) -> None:
    now = datetime.now().strftime("%H:%M:%S")

    if not RICH_AVAILABLE:
        print(
            f"{now} | {level:<7} | {message}",
            flush=True,
        )
        return

    colors = {
        "INFO": "cyan",
        "OK": "green",
        "WARN": "yellow",
        "ERROR": "red",
    }

    line = Text()

    line.append(now, style="dim")
    line.append("  ")

    line.append(
        f"{level:<7}",
        style=colors.get(level, "white"),
    )

    line.append("  ")
    line.append(str(message))

    console.print(line)


def print_header() -> None:
    if not RICH_AVAILABLE:
        print("=" * 72)
        print("OIC UNSUPPORTED FIELD VERIFIER")
        print("NATIVE MEASURE / CALCULATED MEASURE RESOLUTION")
        print("=" * 72)
        return

    console.print(
        Panel(
            "[bold cyan]OIC UNSUPPORTED FIELD VERIFIER[/]\n"
            "[white]NATIVE + CALCULATED MEASURE RESOLUTION[/]\n"
            "[dim]UNSUPPORTED ONLY | LIVE XMLA | AUTO RESUME[/]",
            border_style="cyan",
        )
    )


# ============================================================
# GENERAL UTILITIES
# ============================================================

def utc_now() -> str:
    return datetime.now(
        timezone.utc
    ).isoformat(timespec="seconds")


def clean(value) -> str:
    return str(
        "" if value is None else value
    ).strip()


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def normalized(value) -> str:
    return " ".join(
        clean(value)
        .replace("\u00a0", " ")
        .casefold()
        .split()
    )


def normalized_key(value) -> str:
    return re.sub(
        r"[^A-Z0-9]",
        "",
        clean(value).upper(),
    )


def normalized_identifier(value) -> str:
    return " ".join(
        clean(value)
        .replace("_", " ")
        .casefold()
        .split()
    )


def valid_measure(value: str) -> bool:
    return bool(
        MEASURE_PATTERN.fullmatch(
            clean(value)
        )
    )


def valid_hierarchy(value: str) -> bool:
    return bool(
        HIERARCHY_PATTERN.fullmatch(
            clean(value)
        )
    )


def mdx_quote(value: str) -> str:
    return "[" + value.replace("]", "]]") + "]"


def digest(value) -> str:
    payload = json.dumps(
        value,
        ensure_ascii=False,
        sort_keys=True,
        default=str,
    ).encode("utf-8")

    return hashlib.sha256(
        payload
    ).hexdigest()


def safe_error(
    exception: Exception,
    secrets: tuple[str, ...] = (),
) -> str:
    message = (
        f"{type(exception).__name__}: {exception}"
    )

    for secret in secrets:
        if secret and len(secret) >= 6:
            message = message.replace(
                secret,
                "[REDACTED]",
            )

    message = re.sub(
        r"(?i)Bearer\s+[A-Za-z0-9_.\-]+",
        "Bearer [REDACTED]",
        message,
    )

    message = re.sub(
        r"eyJ[A-Za-z0-9_-]+\."
        r"[A-Za-z0-9_-]+\."
        r"[A-Za-z0-9_-]+",
        "[REDACTED_JWT]",
        message,
    )

    return message[:1000]


def sanitize_metadata(value):
    if isinstance(value, dict):
        return {
            key: sanitize_metadata(item)
            for key, item in value.items()
            if normalized_key(key) not in SENSITIVE_KEYS
        }

    if isinstance(value, list):
        return [
            sanitize_metadata(item)
            for item in value
        ]

    if isinstance(value, str):
        return re.sub(
            r"eyJ[A-Za-z0-9_-]+\."
            r"[A-Za-z0-9_-]+\."
            r"[A-Za-z0-9_-]+",
            "[REDACTED_JWT]",
            value,
        )

    return value


def decode_json_value(value):
    for _ in range(4):
        if not isinstance(value, str):
            return value

        stripped = value.strip()

        if not stripped.startswith(
            ("{", "[", '"')
        ):
            return value

        try:
            decoded = json.loads(stripped)
        except (ValueError, TypeError):
            return value

        if decoded == value:
            return value

        value = decoded

    return value


# ============================================================
# ENVIRONMENT / AUTHENTICATION
# ============================================================

def read_env() -> dict[str, str]:
    values = {}

    if ENV_FILE.is_file():
        for original in ENV_FILE.read_text(
            encoding="utf-8-sig"
        ).splitlines():

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

            values[key] = value

    for key in (
        "OIC_COOKIE",
        "OIC_BEARER_TOKEN",
    ):
        if os.environ.get(key):
            values[key] = os.environ[key]

    return values


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

    return value.strip('"')


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


# ============================================================
# WINDOWS-SAFE PERSISTENCE
# ============================================================

def is_file_lock(exc: OSError) -> bool:
    return (
        isinstance(exc, PermissionError)
        or getattr(exc, "winerror", None)
        in (5, 32, 33)
    )


def recovery_files(path: Path) -> list[Path]:
    return list(
        path.parent.glob(
            f"{path.stem}.recovery.*{path.suffix}"
        )
    )


def file_mtime(path: Path) -> int:
    try:
        return path.stat().st_mtime_ns
    except OSError:
        return 0


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
):
    candidates = [
        path,
        *recovery_files(path),
    ]

    candidates.sort(
        key=file_mtime,
        reverse=True,
    )

    for candidate in candidates:
        data = read_json(
            candidate,
            None,
        )

        if isinstance(data, dict):
            return data

    return default


def commit_file(
    temporary: Path,
    destination: Path,
) -> Path:
    for delay in FILE_RETRY_DELAYS:
        try:
            os.replace(
                temporary,
                destination,
            )

            for old in recovery_files(destination):
                try:
                    old.unlink(missing_ok=True)
                except OSError:
                    pass

            return destination

        except OSError as exc:
            if not is_file_lock(exc):
                raise

            time.sleep(delay)

    recovery = destination.with_name(
        f"{destination.stem}.recovery."
        f"{time.time_ns()}."
        f"{uuid.uuid4().hex[:8]}"
        f"{destination.suffix}"
    )

    with temporary.open("rb") as source:
        with recovery.open("xb") as target:
            target.write(source.read())
            target.flush()
            os.fsync(target.fileno())

    log(
        f"FILE LOCK | Recovery saved: {recovery.name}",
        "WARN",
    )

    return recovery


def atomic_json(
    path: Path,
    value: dict,
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

    payload = json.dumps(
        sanitize_metadata(value),
        ensure_ascii=False,
        indent=2,
        default=str,
    ).encode("utf-8")

    try:
        with temporary.open("xb") as stream:
            stream.write(payload)
            stream.flush()
            os.fsync(stream.fileno())

        return commit_file(
            temporary,
            path,
        )

    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


# ============================================================
# ORIGINAL REPORT LOADER
# ============================================================

def load_source_reports() -> list[dict]:
    reports = []

    for code in CATEGORIES:
        for cube_type in CUBE_TYPES:
            path = (
                SOURCE_DIR
                / f"{code}_{cube_type}.json"
            )

            report = read_latest_json(
                path,
                None,
            )

            if not isinstance(report, dict):
                continue

            if report.get("insurance_code") != code:
                continue

            if report.get("cube_type") != cube_type:
                continue

            reports.append(report)

    return reports


def load_unsupported_fields(
    selected_code: str | None = None,
    selected_cube_type: str | None = None,
) -> list[dict]:
    targets = []

    for report in load_source_reports():
        code = report["insurance_code"]
        cube_type = report["cube_type"]

        if selected_code and code != selected_code:
            continue

        if (
            selected_cube_type
            and cube_type != selected_cube_type
        ):
            continue

        for field in report.get("fields", []):
            if not isinstance(field, dict):
                continue

            if field.get("status") != "UNSUPPORTED":
                continue

            details = field.get("details") or {}

            if not isinstance(details, dict):
                details = {}

            field_id = clean(
                field.get("id")
            )

            if not field_id:
                field_id = digest(field)[:24]

            target = {
                "insurance_code": code,
                "cube_type": cube_type,
                "catalog": report.get(
                    "catalog",
                    CATEGORIES[code]["catalog"],
                ),
                "cube": report.get(
                    "cube",
                    f"{code}_AGGREGATE_{cube_type}",
                ),
                "field_id": field_id,
                "field_caption": clean(
                    field.get("caption")
                ),
                "field_type": clean(
                    field.get("type")
                ),
                "field_folder": clean(
                    field.get("folder")
                ),
                "data_table": clean(
                    field.get("data_table")
                    or details.get("DATATABLE")
                ),
                "data_field": clean(
                    field.get("data_field")
                    or details.get("DATAFIELD")
                ),
                "summary_type": clean(
                    details.get("SUMMARYTYPE")
                ),
                "calculate": details.get(
                    "CALCULATE"
                ),
                "original_status": "UNSUPPORTED",
                "original_mapping_reason": field.get(
                    "mapping_reason"
                ),
                "original_metadata": details,
            }

            target["target_key"] = (
                f"{code}:{cube_type}:{field_id}"
            )

            target["source_fingerprint"] = digest(
                {
                    "catalog": target["catalog"],
                    "cube": target["cube"],
                    "field": target["original_metadata"],
                    "caption": target["field_caption"],
                    "data_field": target["data_field"],
                }
            )

            targets.append(target)

    return targets


# ============================================================
# XMLA REQUEST BUILDERS
# ============================================================

def make_discover_request(
    catalog: str,
    request_type: str,
    cube: str,
) -> bytes:
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
    statement: str,
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
    ).text = statement

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
        f"{{{XMLA_NS}}}LocaleIdentifier",
    ).text = "9"

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}ShowHiddenCubes",
    ).text = "True"

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}Format",
    ).text = "Multidimensional"

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}Timeout",
    ).text = "60"

    return ET.tostring(
        envelope,
        encoding="utf-8",
    )


# ============================================================
# XMLA RESPONSE PARSERS
# ============================================================

def parse_soap(content: bytes) -> ET.Element:
    preview = content[:500].decode(
        "utf-8",
        errors="ignore",
    )

    upper_preview = preview.upper()

    if "SESSION_EXPIRED" in upper_preview:
        raise AuthenticationError(
            "OIC_COOKIE session expired"
        )

    if (
        preview.lstrip().lower().startswith("<html")
        or preview.lstrip().lower().startswith(
            "<!doctype html"
        )
    ):
        raise AuthenticationError(
            "Unexpected HTML response; "
            "OIC session may require renewal"
        )

    try:
        root = ET.fromstring(content)

    except ET.ParseError as exc:
        raise XMLAError(
            "Invalid XMLA response"
        ) from exc

    faults = [
        item
        for item in root.iter()
        if local_name(item.tag) == "Fault"
    ]

    if faults:
        message = " ".join(
            clean(node.text)
            for fault in faults
            for node in fault.iter()
            if node.text
        )

        raise XMLAError(
            f"XMLA SOAP Fault: {message[:600]}"
        )

    return root


def parse_discover_rows(
    content: bytes,
    catalog: str,
    cube: str,
) -> list[dict]:
    root = parse_soap(content)

    rowset = next(
        (
            item
            for item in root.iter()
            if item.tag == f"{{{ROWSET_NS}}}root"
        ),
        None,
    )

    if rowset is None:
        raise XMLAError(
            "XMLA Discover rowset not found"
        )

    rows = []

    for element in rowset.iter():
        if local_name(element.tag).lower() != "row":
            continue

        row = {
            local_name(child.tag).upper(): (
                child.text or ""
            )
            for child in element
        }

        row_catalog = clean(
            row.get("CATALOG_NAME")
        )

        row_cube = clean(
            row.get("CUBE_NAME")
        )

        if (
            row_catalog
            and row_catalog.casefold()
            != catalog.casefold()
        ):
            continue

        if (
            row_cube
            and row_cube.casefold()
            != cube.casefold()
        ):
            continue

        if row:
            rows.append(row)

    return rows


def parse_execute_response(
    content: bytes,
    expected_measure: str,
    expected_hierarchy: str | None = None,
) -> dict:
    root = parse_soap(content)

    dataset = next(
        (
            node
            for node in root.iter()
            if node.tag == f"{{{MD_NS}}}root"
        ),
        None,
    )

    if dataset is None:
        raise XMLAError(
            "Multidimensional dataset not found"
        )

    axes = {
        node.attrib.get("name"): node
        for node in dataset.iter()
        if local_name(node.tag) == "Axis"
    }

    axis0 = axes.get("Axis0")

    if axis0 is None:
        raise XMLAError(
            "Execute response missing Axis0"
        )

    measures = set()

    for member in axis0.iter():
        if local_name(member.tag) != "Member":
            continue

        for child in member:
            if local_name(child.tag) in (
                "UName",
                "MEMBER_UNIQUE_NAME",
            ):
                if child.text:
                    measures.add(
                        child.text.strip()
                    )

    if expected_measure.casefold() not in {
        item.casefold()
        for item in measures
    }:
        raise XMLAError(
            "Expected Measure was not returned on Axis0"
        )

    year_members = []

    if expected_hierarchy:
        axis1 = axes.get("Axis1")

        if axis1 is None:
            raise XMLAError(
                "Execute response missing Axis1"
            )

        hierarchy_found = False

        for member in axis1.iter():
            if local_name(member.tag) != "Member":
                continue

            hierarchy = clean(
                member.attrib.get("Hierarchy")
            )

            member_name = ""
            caption = ""

            for child in member:
                name = local_name(child.tag)

                if name in (
                    "UName",
                    "MEMBER_UNIQUE_NAME",
                ):
                    member_name = clean(child.text)

                elif name in (
                    "Caption",
                    "MEMBER_CAPTION",
                ):
                    caption = clean(child.text)

            if (
                hierarchy.casefold()
                == expected_hierarchy.casefold()
                or member_name.casefold().startswith(
                    expected_hierarchy.casefold() + "."
                )
            ):
                hierarchy_found = True

                year_members.append({
                    "unique_name": member_name,
                    "caption": caption,
                })

        if not hierarchy_found:
            raise XMLAError(
                "Expected year hierarchy was not returned "
                "on Axis1"
            )

    cell_data = next(
        (
            node
            for node in dataset.iter()
            if local_name(node.tag) == "CellData"
        ),
        None,
    )

    if cell_data is None:
        raise XMLAError(
            "Execute response missing CellData"
        )

    cells = []
    nonempty = 0
    errors = []

    for cell in cell_data:
        if local_name(cell.tag) != "Cell":
            continue

        value = None
        formatted = None

        for child in cell:
            name = local_name(child.tag)

            if name == "Value":
                value = child.text

            elif name == "FmtValue":
                formatted = child.text

        for child in cell.iter():
            if local_name(child.tag) == "Error":
                errors.append(
                    clean(
                        child.attrib.get("Description")
                        or child.text
                        or "Cell calculation error"
                    )
                )

            for attr, attr_value in child.attrib.items():
                if (
                    local_name(attr) == "type"
                    and clean(attr_value)
                    .casefold()
                    .endswith(":error")
                ):
                    errors.append(
                        clean(child.text)
                        or "XMLA error value"
                    )

        if value is not None and clean(value):
            nonempty += 1

        cells.append({
            "ordinal": cell.attrib.get("CellOrdinal"),
            "value": value,
            "formatted_value": formatted,
        })

    if errors:
        raise XMLAError(
            "Cell calculation failed: "
            + "; ".join(errors[:3])
        )

    return {
        "expected_measure": expected_measure,
        "returned_measures": sorted(measures),
        "axis_verified": True,
        "year_hierarchy": expected_hierarchy,
        "year_members": year_members,
        "sample_cells": len(cells),
        "nonempty_cells": nonempty,
        "cells": cells[:10],
        "has_data": nonempty > 0,
    }


# ============================================================
# WEBFORM AUTHORIZATION DISCOVERY
# ============================================================

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

AUTH_PATTERN = re.compile(
    r"""
    ['"]?Authorization['"]?
    \s*:\s*
    ['"](?P<token>[A-Fa-f0-9]{32,128})['"]
    """,
    re.IGNORECASE | re.VERBOSE,
)

SCRIPT_PATTERN = re.compile(
    r"""<script\b[^>]*\bsrc\s*=\s*["']([^"']+)["']""",
    re.IGNORECASE,
)


def extract_webform_authorizations(
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

        end = min(
            match.end() + 2500,
            (
                matches[index + 1].start()
                if index + 1 < len(matches)
                else len(source)
            ),
        )

        fragment = source[
            match.end():end
        ]

        for item in AUTH_PATTERN.finditer(fragment):
            value = item.group("token")

            if value not in result[method]:
                result[method].append(value)

    return result


def decode_webform_response(
    response: requests.Response,
    method: str,
):
    if response.status_code in (401, 403):
        raise AuthenticationError(
            f"{method}: HTTP {response.status_code}"
        )

    if response.status_code != 200:
        raise ValueError(
            f"{method}: HTTP {response.status_code}"
        )

    try:
        result = response.json()

    except ValueError as exc:
        raise ValueError(
            f"{method}: invalid JSON"
        ) from exc

    if isinstance(result, dict) and "d" in result:
        result = result["d"]

    result = decode_json_value(result)

    if (
        isinstance(result, str)
        and "SESSION_EXPIRED" in result.upper()
    ):
        raise AuthenticationError(
            "OIC_COOKIE session expired"
        )

    if not isinstance(result, (list, dict)):
        raise ValueError(
            f"{method}: unexpected response format"
        )

    return result


# ============================================================
# OIC CLIENT
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
            raise AuthenticationError(
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

    def close(self):
        self.session.close()

    def secrets(self) -> tuple[str, ...]:
        values = [
            self.cookie,
            self.token,
        ]

        values.extend(
            cookie.value
            for cookie in self.session.cookies
        )

        for authorization in self.webform_auth.values():
            for tokens in authorization.values():
                values.extend(tokens)

        return tuple(values)

    def referer(self, code: str) -> str:
        return (
            WEBFORM_URL
            + "?Class="
            + CATEGORIES[code]["web_class"]
        )

    def refresh_token(
        self,
        code: str,
    ) -> None:
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
            raise AuthenticationError(
                f"GenToken HTTP {response.status_code}"
            )

        try:
            payload = response.json()

        except ValueError as exc:
            raise AuthenticationError(
                "GenToken returned invalid JSON"
            ) from exc

        if isinstance(payload, dict):
            payload = payload.get("d")

        payload = decode_json_value(payload)

        if not isinstance(payload, str):
            raise AuthenticationError(
                "GenToken returned unexpected data"
            )

        if "SESSION_EXPIRED" in payload.upper():
            raise AuthenticationError(
                "OIC_COOKIE session expired"
            )

        token = normalize_token(payload)

        if len(token.split(".")) != 3:
            raise AuthenticationError(
                "GenToken did not return a JWT"
            )

        expiry = jwt_exp(token)

        if expiry and expiry <= time.time() + 30:
            raise AuthenticationError(
                "GenToken returned an expired JWT"
            )

        self.token = token

        log(
            f"{code}: JWT refreshed",
            "OK",
        )

    def bearer(
        self,
        code: str,
    ) -> str:
        expiry = jwt_exp(self.token)

        if (
            not self.token
            or (
                expiry
                and expiry <= time.time() + 60
            )
        ):
            self.refresh_token(code)

        return self.token

    def xmla(
        self,
        code: str,
        body: bytes,
    ) -> bytes:
        refreshed = False

        for attempt in range(3):
            token = self.bearer(code)

            try:
                response = self.session.post(
                    XMLA_URL,
                    data=body,
                    headers={
                        "Authorization": f"Bearer {token}",
                        "Content-Type": "text/xml",
                        "Accept": "text/plain, */*; q=0.01",
                        "X-Requested-With": "XMLHttpRequest",
                        "Origin": BASE_URL,
                        "Referer": self.referer(code),
                    },
                    timeout=(10, self.timeout),
                    allow_redirects=False,
                )

            except requests.RequestException:
                if attempt == 2:
                    raise

                time.sleep(attempt + 1)
                continue

            if response.status_code in (401, 403):
                if not refreshed:
                    self.refresh_token(code)
                    refreshed = True
                    continue

                raise AuthenticationError(
                    f"XMLA HTTP {response.status_code}"
                )

            if response.status_code in (
                301, 302, 303, 307, 308
            ):
                raise AuthenticationError(
                    "XMLA request redirected"
                )

            if response.status_code == 200:
                if (
                    b"SESSION_EXPIRED"
                    in response.content[:500].upper()
                ):
                    if not refreshed:
                        self.refresh_token(code)
                        refreshed = True
                        continue

                    raise AuthenticationError(
                        "OIC session expired"
                    )

                return response.content

            if (
                response.status_code
                in (408, 429, 500, 502, 503, 504)
                and attempt < 2
            ):
                try:
                    delay = min(
                        float(
                            response.headers.get(
                                "Retry-After",
                                2 ** attempt,
                            )
                        ),
                        30.0,
                    )

                except (TypeError, ValueError):
                    delay = 2 ** attempt

                time.sleep(max(0.5, delay))
                continue

            raise XMLAError(
                f"XMLA HTTP {response.status_code}"
            )

        raise XMLAError(
            "XMLA retry limit reached"
        )

    def discover(
        self,
        code: str,
        catalog: str,
        cube: str,
        request_type: str,
    ) -> list[dict]:
        payload = make_discover_request(
            catalog,
            request_type,
            cube,
        )

        response = self.xmla(
            code,
            payload,
        )

        return parse_discover_rows(
            response,
            catalog,
            cube,
        )

    def execute(
        self,
        code: str,
        catalog: str,
        statement: str,
        expected_measure: str,
        expected_hierarchy: str | None = None,
    ) -> dict:
        payload = make_execute_request(
            catalog,
            statement,
        )

        response = self.xmla(
            code,
            payload,
        )

        return parse_execute_response(
            response,
            expected_measure,
            expected_hierarchy,
        )

    def load_webform_auth(
        self,
        code: str,
        *,
        refresh: bool = False,
    ) -> dict:
        if code in self.webform_auth and not refresh:
            return self.webform_auth[code]

        referer = self.referer(code)

        response = self.session.get(
            referer,
            headers={
                "Referer": BASE_URL + "/",
                "Accept": (
                    "text/html,application/xhtml+xml,*/*"
                ),
            },
            timeout=(10, self.timeout),
            allow_redirects=False,
        )

        if response.status_code in (
            301, 302, 303, 307, 308, 401, 403
        ):
            raise AuthenticationError(
                f"WebForm HTTP {response.status_code}"
            )

        if response.status_code != 200:
            raise ValueError(
                f"WebForm HTTP {response.status_code}"
            )

        auth = extract_webform_authorizations(
            response.text
        )

        urls = []
        seen = set()

        for match in SCRIPT_PATTERN.finditer(
            response.text
        ):
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

        for url in urls[:20]:
            if all(auth.values()):
                break

            try:
                script = self.session.get(
                    url,
                    headers={
                        "Referer": referer,
                    },
                    timeout=(10, self.timeout),
                    allow_redirects=False,
                )

                if script.status_code != 200:
                    continue

                extra = (
                    extract_webform_authorizations(
                        script.text
                    )
                )

                for method, values in extra.items():
                    for value in values:
                        if value not in auth[method]:
                            auth[method].append(value)

            except requests.RequestException:
                continue

        self.webform_auth[code] = auth

        return auth

    def webform(
        self,
        code: str,
        method: str,
        payload: dict,
    ):
        if method not in WEBFORM_METHODS:
            raise ValueError(
                f"Invalid method: {method}"
            )

        last_error = None

        for refresh in (False, True):
            authorizations = self.load_webform_auth(
                code,
                refresh=refresh,
            ).get(method, [])

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
                            "application/json, "
                            "text/javascript, */*; q=0.01"
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

                except AuthenticationError as exc:
                    last_error = exc

                    if "SESSION_EXPIRED" in str(exc).upper():
                        raise

                except ValueError as exc:
                    last_error = exc

                    if response.status_code != 200:
                        raise

        if isinstance(
            last_error,
            AuthenticationError,
        ):
            raise last_error

        raise ValueError(
            f"{method}: Authorization or response unavailable"
        )


# ============================================================
# METADATA EXTRACTION
# ============================================================

def flatten_records(
    value,
    *,
    depth: int = 0,
    max_depth: int = 12,
):
    if depth > max_depth:
        return

    value = decode_json_value(value)

    if isinstance(value, dict):
        yield value

        for item in value.values():
            yield from flatten_records(
                item,
                depth=depth + 1,
                max_depth=max_depth,
            )

    elif isinstance(value, list):
        for item in value:
            yield from flatten_records(
                item,
                depth=depth + 1,
                max_depth=max_depth,
            )


def get_casefold(
    record: dict,
    *keys: str,
):
    indexed = {
        normalized_key(key): value
        for key, value in record.items()
    }

    for key in keys:
        lookup = normalized_key(key)

        if lookup in indexed:
            return indexed[lookup]

    return None


def get_record_caption(
    record: dict,
) -> str:
    return clean(
        get_casefold(
            record,
            "CAPTION",
            "MEASURE_CAPTION",
            "caption",
        )
    )


def get_record_folder(
    record: dict,
) -> str:
    return clean(
        get_casefold(
            record,
            "DIMENSION",
            "displayFolder",
            "MEASURE_DISPLAY_FOLDER",
        )
    )


def find_live_structure(
    target: dict,
    structure,
) -> list[dict]:
    original = target["original_metadata"]

    original_seq = clean(
        original.get("SEQ")
    )

    matches = []

    for record in flatten_records(structure):
        caption = get_record_caption(record)

        if not caption:
            continue

        if (
            normalized(caption)
            != normalized(target["field_caption"])
        ):
            continue

        data_field = clean(
            get_casefold(
                record,
                "DATAFIELD",
            )
        )

        data_table = clean(
            get_casefold(
                record,
                "DATATABLE",
            )
        )

        folder = get_record_folder(record)

        seq = clean(
            get_casefold(record, "SEQ")
        )

        if (
            target["data_field"]
            and data_field
            and normalized_identifier(data_field)
            != normalized_identifier(
                target["data_field"]
            )
        ):
            continue

        if (
            target["data_table"]
            and data_table
            and normalized_identifier(data_table)
            != normalized_identifier(
                target["data_table"]
            )
        ):
            continue

        if (
            target["field_folder"]
            and folder
            and normalized(folder)
            != normalized(
                target["field_folder"]
            )
        ):
            continue

        if original_seq and seq and original_seq != seq:
            continue

        matches.append(record)

    return matches


def extract_identifiers(
    record: dict,
) -> list[str]:
    values = set()

    for key, value in record.items():
        if (
            normalized_key(key)
            not in IDENTIFIER_KEYS
        ):
            continue

        if not isinstance(value, str):
            continue

        value = value.strip()

        if valid_measure(value):
            values.add(value)

    return sorted(values)


def extract_formulas(
    record: dict,
) -> list[dict]:
    formulas = []

    for key, value in record.items():
        if (
            normalized_key(key)
            not in FORMULA_KEYS
        ):
            continue

        if not isinstance(value, str):
            continue

        value = value.strip()

        if normalized(value) in {
            "",
            "0",
            "1",
            "true",
            "false",
            "yes",
            "no",
            "null",
            "none",
        }:
            continue

        formulas.append({
            "key": key,
            "expression": value,
        })

    return formulas


def extract_template_items(
    value,
    code: str,
    cube: str,
    *,
    source: str,
    inherited_cube: str | None = None,
    depth: int = 0,
) -> list[dict]:
    if depth > 12:
        return []

    value = decode_json_value(value)

    items = []

    if isinstance(value, list):
        for item in value:
            items.extend(
                extract_template_items(
                    item,
                    code,
                    cube,
                    source=source,
                    inherited_cube=inherited_cube,
                    depth=depth + 1,
                )
            )

        return items

    if not isinstance(value, dict):
        return items

    record_class = clean(
        get_casefold(
            value,
            "CLASS",
            "INSURANCE_CODE",
        )
    )

    if record_class and record_class.upper() not in {
        code,
        CATEGORIES[code]["web_class"],
    }:
        return items

    record_cube = clean(
        get_casefold(
            value,
            "CUBE_NAME",
            "cubeName",
        )
    ) or inherited_cube

    if (
        record_cube
        and record_cube.casefold()
        != cube.casefold()
    ):
        return items

    caption = get_record_caption(value)

    area = clean(
        get_casefold(value, "AREA")
    ).casefold()

    identifiers = extract_identifiers(value)
    formulas = extract_formulas(value)

    # Unscoped menu items are retained for diagnostics,
    # but cannot establish a mapping by themselves.

    scoped = bool(
        record_cube
        and record_cube.casefold()
        == cube.casefold()
    )

    if (
        caption
        and area in ("data", "measure", "measures", "")
        and (identifiers or formulas)
    ):
        items.append({
            "caption": caption,
            "folder": get_record_folder(value),
            "area": area,
            "identifiers": identifiers,
            "formulas": formulas,
            "source": source,
            "cube_scoped": scoped,
            "metadata": sanitize_metadata(value),
        })

    for child in value.values():
        items.extend(
            extract_template_items(
                child,
                code,
                cube,
                source=source,
                inherited_cube=record_cube,
                depth=depth + 1,
            )
        )

    return items


# ============================================================
# LIVE CUBE METADATA
# ============================================================

def discover_cube_metadata(
    client: OICClient,
    target: dict,
) -> dict:
    code = target["insurance_code"]
    catalog = target["catalog"]
    cube = target["cube"]
    cube_type = target["cube_type"]

    metadata = {
        "insurance_code": code,
        "cube_type": cube_type,
        "catalog": catalog,
        "cube": cube,
        "discovered_at": utc_now(),
        "measures": [],
        "hierarchies": [],
        "structure": [],
        "templates": [],
        "errors": [],
        "sources": [],
    }

    log(
        f"DISCOVER {code}/{cube_type} | {cube}"
    )

    for request_type, destination in (
        ("MDSCHEMA_MEASURES", "measures"),
        ("MDSCHEMA_HIERARCHIES", "hierarchies"),
    ):
        try:
            records = client.discover(
                code,
                catalog,
                cube,
                request_type,
            )

            metadata[destination] = records

            metadata["sources"].append(
                request_type
            )

            log(
                f"{code}/{cube_type}: "
                f"{destination.upper()}={len(records)}"
            )

        except AuthenticationError:
            raise

        except Exception as exc:
            metadata["errors"].append({
                "source": request_type,
                "error": safe_error(
                    exc,
                    client.secrets(),
                ),
            })

    if "MDSCHEMA_MEASURES" not in metadata["sources"]:
        raise ResolutionError(
            "Native Measure discovery failed; "
            "cannot establish reliable mappings"
        )

    webform_targets = (
        (
            "GetStructure",
            {
                "Template": [
                    code,
                    cube,
                ],
            },
        ),
        (
            "GET_MENUTEMPLATE_LIST",
            {
                "Template": [
                    CATEGORIES[code]["web_class"],
                ],
            },
        ),
        (
            "BindChangeTemplate",
            {
                "Template": [
                    code,
                    cube,
                ],
            },
        ),
    )

    for method, payload in webform_targets:
        try:
            records = client.webform(
                code,
                method,
                payload,
            )

            metadata["sources"].append(method)

            if method == "GetStructure":
                metadata["structure"] = records

            else:
                metadata["templates"].extend(
                    extract_template_items(
                        records,
                        code,
                        cube,
                        source=method,
                        inherited_cube=(
                            cube
                            if method == "BindChangeTemplate"
                            else None
                        ),
                    )
                )

        except AuthenticationError as exc:
            if "SESSION_EXPIRED" in str(exc).upper():
                raise

            metadata["errors"].append({
                "source": method,
                "error": safe_error(
                    exc,
                    client.secrets(),
                ),
            })

        except Exception as exc:
            metadata["errors"].append({
                "source": method,
                "error": safe_error(
                    exc,
                    client.secrets(),
                ),
            })

    path = (
        METADATA_DIR
        / f"{code}_{cube_type}.json"
    )

    atomic_json(
        path,
        metadata,
    )

    log(
        f"{code}/{cube_type}: "
        f"TEMPLATES={len(metadata['templates'])} | "
        f"METADATA ERRORS={len(metadata['errors'])}"
    )

    return metadata


# ============================================================
# MEASURE INDEX
# ============================================================

def build_measure_index(
    records: list[dict],
) -> dict:
    by_mdx = {}
    by_caption = defaultdict(set)
    by_name = defaultdict(set)

    for record in records:
        mdx = clean(
            record.get("MEASURE_UNIQUE_NAME")
        )

        if not valid_measure(mdx):
            continue

        caption = clean(
            record.get("MEASURE_CAPTION")
        )

        measure_name = clean(
            record.get("MEASURE_NAME")
        )

        mdx_key = mdx.casefold()

        by_mdx[mdx_key] = {
            "mdx": mdx,
            "caption": caption,
            "name": measure_name,
            "folder": clean(
                record.get(
                    "MEASURE_DISPLAY_FOLDER"
                )
            ),
            "aggregator": clean(
                record.get(
                    "MEASURE_AGGREGATOR"
                )
            ),
            "expression": clean(
                record.get(
                    "EXPRESSION"
                )
            ),
            "metadata": record,
        }

        if caption:
            by_caption[
                normalized(caption)
            ].add(mdx)

        if measure_name:
            by_name[
                normalized_identifier(
                    measure_name
                )
            ].add(mdx)

    return {
        "by_mdx": by_mdx,
        "by_caption": by_caption,
        "by_name": by_name,
    }


def matching_template_items(
    target: dict,
    metadata: dict,
) -> list[dict]:
    matches = []

    for item in metadata.get(
        "templates",
        [],
    ):
        if not item.get("cube_scoped"):
            continue

        if normalized(
            item.get("caption")
        ) != normalized(
            target["field_caption"]
        ):
            continue

        folder = clean(
            item.get("folder")
        )

        if (
            folder
            and target["field_folder"]
            and normalized(folder)
            != normalized(
                target["field_folder"]
            )
        ):
            continue

        matches.append(item)

    return matches


# ============================================================
# NATIVE MEASURE RESOLUTION
# ============================================================

def resolve_native_measure(
    target: dict,
    metadata: dict,
    measure_index: dict,
) -> dict:
    by_mdx = measure_index["by_mdx"]

    templates = matching_template_items(
        target,
        metadata,
    )

    live_structure = find_live_structure(
        target,
        metadata.get("structure", []),
    )

    evidence = []

    def result(
        mdx: str,
        source: str,
        reason: str,
    ) -> dict:
        actual = by_mdx.get(
            mdx.casefold()
        )

        return {
            "mdx": (
                actual["mdx"]
                if actual
                else mdx
            ),
            "mapping_source": source,
            "reason": reason,
            "metadata_confirmed": actual is not None,
            "measure_metadata": (
                actual["metadata"]
                if actual
                else None
            ),
            "evidence": evidence,
            "ambiguous": False,
        }

    # --------------------------------------------------------
    # 1. Explicit MDX from the original/live metadata.
    # --------------------------------------------------------

    explicit = set()

    for row in (
        [target["original_metadata"]]
        + live_structure
    ):
        explicit.update(
            extract_identifiers(row)
        )

    explicit_index = {
        value.casefold(): value
        for value in explicit
    }

    if len(explicit_index) == 1:
        mdx = next(
            iter(explicit_index.values())
        )

        evidence.append({
            "source": "GETSTRUCTURE",
            "candidate": mdx,
        })

        return result(
            mdx,
            "EXPLICIT_FIELD_MDX",
            "Explicit Measure identifier in OIC field metadata",
        )

    if len(explicit_index) > 1:
        return {
            "mdx": None,
            "ambiguous": True,
            "reason": "Conflicting explicit MDX identifiers",
            "candidates": sorted(explicit),
        }

    # --------------------------------------------------------
    # 2. Exact field mapping from cube-scoped OIC templates.
    # --------------------------------------------------------

    template_candidates = {
        mdx.casefold(): mdx
        for item in templates
        for mdx in item.get("identifiers", [])
        if valid_measure(mdx)
    }

    if len(template_candidates) == 1:
        mdx = next(
            iter(template_candidates.values())
        )

        evidence.append({
            "source": "OIC_TEMPLATE",
            "candidate": mdx,
        })

        return result(
            mdx,
            "OIC_TEMPLATE",
            "Exact caption/folder in a cube-scoped OIC template",
        )

    if len(template_candidates) > 1:
        return {
            "mdx": None,
            "ambiguous": True,
            "reason": "Multiple conflicting template identifiers",
            "candidates": sorted(
                template_candidates.values()
            ),
        }

    # --------------------------------------------------------
    # 3. Exact caption match in XMLA Measure metadata.
    # --------------------------------------------------------

    caption_matches = set(
        measure_index["by_caption"].get(
            normalized(
                target["field_caption"]
            ),
            set(),
        )
    )

    if len(caption_matches) > 1:
        folder = normalized(
            target["field_folder"]
        )

        if folder:
            folder_matches = {
                mdx
                for mdx in caption_matches
                if normalized(
                    by_mdx[mdx.casefold()]["folder"]
                ) == folder
            }

            if folder_matches:
                caption_matches = folder_matches

    if len(caption_matches) == 1:
        mdx = next(iter(caption_matches))

        evidence.append({
            "source": "XMLA_CAPTION",
            "candidate": mdx,
        })

        return result(
            mdx,
            "XMLA_EXACT_CAPTION",
            "Exact unique Measure caption returned by XMLA",
        )

    if len(caption_matches) > 1:
        return {
            "mdx": None,
            "ambiguous": True,
            "reason": "Caption matches multiple Measures",
            "candidates": sorted(caption_matches),
        }

    # --------------------------------------------------------
    # 4. Source-name matching.
    #
    # A calculated field must not be incorrectly mapped
    # to its underlying base Measure solely by DATAFIELD.
    # --------------------------------------------------------

    calculation = normalized(
        target.get("calculate")
    )

    summary_type = normalized(
        target.get("summary_type")
    )

    derived = (
        calculation not in (
            "",
            "0",
            "false",
            "none",
            "null",
            "no",
        )
        or summary_type not in ("", "sum")
    )

    source_name = normalized_identifier(
        target["data_field"]
    )

    name_matches = set(
        measure_index["by_name"].get(
            source_name,
            set(),
        )
    )

    if len(name_matches) == 1 and not derived:
        mdx = next(iter(name_matches))

        return result(
            mdx,
            "XMLA_EXACT_MEASURE_NAME",
            "Exact native source Measure name",
        )

    if name_matches:
        evidence.append({
            "source": "SOURCE_NAME",
            "candidates": sorted(name_matches),
            "accepted": False,
            "reason": (
                "Source name alone cannot establish "
                "calculated-field semantics"
            ),
        })

    return {
        "mdx": None,
        "ambiguous": False,
        "reason": (
            "No reliable native Measure mapping found"
        ),
        "evidence": evidence,
    }


# ============================================================
# DOCUMENTED FORMULA RESOLUTION
# ============================================================

FORMULA_TOKEN = re.compile(
    r"""
    (?P<SPACE>\s+)
    |
    (?P<MEASURE>
        \[Measures\]\.\[(?:[^\]\r\n]|\]\])+\]
    )
    |
    (?P<NUMBER>
        (?:\d+(?:\.\d*)?|\.\d+)
        (?:[eE][+-]?\d+)?
    )
    |
    (?P<OPERATOR>[+\-*/()])
    """,
    re.IGNORECASE | re.VERBOSE,
)


def tokenize_formula(
    expression: str,
) -> list[tuple[str, str]]:
    expression = clean(expression)

    if expression.startswith("="):
        expression = expression[1:].strip()

    if not expression or len(expression) > 8192:
        raise ResolutionError(
            "Empty or excessively large formula"
        )

    tokens = []
    offset = 0

    while offset < len(expression):
        match = FORMULA_TOKEN.match(
            expression,
            offset,
        )

        if not match:
            raise ResolutionError(
                "Formula contains unsupported syntax "
                f"near character {offset}"
            )

        offset = match.end()

        kind = match.lastgroup
        value = match.group()

        if kind != "SPACE":
            tokens.append((kind, value))

    return tokens


class ArithmeticFormulaParser:

    """
    Supported expressions:

        [Measures].[A] + [Measures].[B]
        [Measures].[A] - [Measures].[B]
        [Measures].[A] * 100
        [Measures].[A] / [Measures].[B]
        ([Measures].[A] + [Measures].[B]) / 1000000

    Unsupported expressions are retained for review.

    No Python eval().
    No JavaScript execution.
    No arbitrary server-side commands.
    No guessed actuarial formulas.
    """

    def __init__(
        self,
        tokens: list[tuple[str, str]],
        measure_index: dict,
    ):
        self.tokens = tokens
        self.position = 0
        self.measure_index = measure_index
        self.dependencies = set()

    def peek(self):
        if self.position >= len(self.tokens):
            return None

        return self.tokens[self.position]

    def consume(self):
        current = self.peek()

        if current is None:
            raise ResolutionError(
                "Unexpected end of formula"
            )

        self.position += 1

        return current

    def parse(self) -> dict:
        if len(self.tokens) > 512:
            raise ResolutionError(
                "Formula is too complex"
            )

        expression = self.parse_expression()

        if self.peek() is not None:
            raise ResolutionError(
                "Unexpected trailing formula content"
            )

        if not self.dependencies:
            raise ResolutionError(
                "Formula references no confirmed native Measures"
            )

        return {
            "expression": expression,
            "dependencies": sorted(
                self.dependencies
            ),
        }

    def parse_expression(self) -> str:
        left = self.parse_term()

        while self.peek() in (
            ("OPERATOR", "+"),
            ("OPERATOR", "-"),
        ):
            _, operator = self.consume()

            right = self.parse_term()

            left = (
                f"({left} {operator} {right})"
            )

        return left

    def parse_term(self) -> str:
        left = self.parse_factor()

        while self.peek() in (
            ("OPERATOR", "*"),
            ("OPERATOR", "/"),
        ):
            _, operator = self.consume()

            right = self.parse_factor()

            left = (
                f"({left} {operator} {right})"
            )

        return left

    def parse_factor(self) -> str:
        kind, value = self.consume()

        if kind == "NUMBER":
            return value

        if kind == "MEASURE":
            actual = self.measure_index[
                "by_mdx"
            ].get(
                value.casefold()
            )

            if actual is None:
                raise ResolutionError(
                    f"Formula dependency not found: {value}"
                )

            mdx = actual["mdx"]

            self.dependencies.add(mdx)

            return mdx

        if kind == "OPERATOR" and value in (
            "+",
            "-",
        ):
            return (
                f"({value}{self.parse_factor()})"
            )

        if kind == "OPERATOR" and value == "(":
            result = self.parse_expression()

            if self.consume() != (
                "OPERATOR",
                ")",
            ):
                raise ResolutionError(
                    "Unbalanced formula parentheses"
                )

            return f"({result})"

        raise ResolutionError(
            f"Invalid formula token: {value}"
        )


def validate_formula(
    expression: str,
    measure_index: dict,
) -> dict:
    tokens = tokenize_formula(
        expression
    )

    parser = ArithmeticFormulaParser(
        tokens,
        measure_index,
    )

    return parser.parse()


def resolve_calculated_measure(
    target: dict,
    metadata: dict,
    measure_index: dict,
) -> dict:
    evidence = []

    structure_matches = find_live_structure(
        target,
        metadata.get("structure", []),
    )

    for record in (
        [target["original_metadata"]]
        + structure_matches
    ):
        for item in extract_formulas(record):
            evidence.append({
                **item,
                "source": "GETSTRUCTURE",
            })

    for template in matching_template_items(
        target,
        metadata,
    ):
        for item in template.get(
            "formulas",
            [],
        ):
            evidence.append({
                **item,
                "source": template["source"],
            })

    accepted = {}
    rejected = []

    for item in evidence:
        expression = item["expression"]

        try:
            parsed = validate_formula(
                expression,
                measure_index,
            )

            key = parsed["expression"].casefold()

            accepted[key] = {
                "expression": parsed["expression"],
                "original_expression": expression,
                "dependencies": parsed["dependencies"],
                "mapping_source": item["source"],
                "formula_key": item["key"],
            }

        except (
            ResolutionError,
            RecursionError,
        ) as exc:
            rejected.append({
                "source": item["source"],
                "key": item["key"],
                "expression": expression,
                "reason": str(exc),
            })

    if len(accepted) == 1:
        result = next(
            iter(accepted.values())
        )

        return {
            "resolved": True,
            **result,
            "rejected_formulas": rejected,
        }

    if len(accepted) > 1:
        return {
            "resolved": False,
            "ambiguous": True,
            "reason": "Conflicting documented formulas",
            "candidates": list(
                accepted.values()
            ),
            "rejected_formulas": rejected,
        }

    return {
        "resolved": False,
        "ambiguous": False,
        "reason": (
            "Formula found but not safely convertible to MDX"
            if evidence
            else "No documented calculation formula found"
        ),
        "rejected_formulas": rejected,
    }


# ============================================================
# YEAR HIERARCHY RESOLUTION
# ============================================================

def resolve_year_hierarchy(
    target: dict,
    metadata: dict,
) -> str | None:
    code = target["insurance_code"]
    cube_type = target["cube_type"]

    expected = {
        "FY": {
            "year",
            "fy year",
            "ปีบัญชี",
        },
        "UY": {
            "uy year",
            "ปีรับประกัน",
            "ปีรับประกันภัย",
        },
        "AY": {
            "ay year",
            "ปีอุบัติเหตุ",
        },
    }[cube_type]

    matches = []

    for record in metadata.get(
        "hierarchies",
        [],
    ):
        hierarchy = clean(
            record.get(
                "HIERARCHY_UNIQUE_NAME"
            )
        )

        if not valid_hierarchy(hierarchy):
            continue

        name = normalized_identifier(
            record.get("HIERARCHY_NAME")
        )

        caption = normalized_identifier(
            record.get("HIERARCHY_CAPTION")
        )

        if (
            name not in expected
            and caption not in expected
        ):
            continue

        score = 0

        hierarchy_upper = hierarchy.upper()

        if (
            f"VW PV {code} AGGREGATE {cube_type}"
            in hierarchy_upper
        ):
            score += 10

        if f"[{cube_type} YEAR]" in hierarchy_upper:
            score += 5

        matches.append(
            (score, hierarchy)
        )

    if not matches:
        return None

    best_score = max(
        score
        for score, _ in matches
    )

    best = {
        hierarchy
        for score, hierarchy in matches
        if score == best_score
    }

    if len(best) != 1:
        return None

    return next(iter(best))


# ============================================================
# MDX EXECUTION
# ============================================================

def build_native_mdx(
    cube: str,
    measure: str,
    year: str | None = None,
) -> str:
    if not valid_measure(measure):
        raise ResolutionError(
            "Invalid native Measure identifier"
        )

    if year:
        if not valid_hierarchy(year):
            raise ResolutionError(
                "Invalid year hierarchy"
            )

        return (
            f"SELECT {{{measure}}} ON COLUMNS, "
            f"Subset({year}.AllMembers, 0, 4) ON ROWS "
            f"FROM {mdx_quote(cube)} "
            "CELL PROPERTIES VALUE, FORMATTED_VALUE"
        )

    return (
        f"SELECT {{{measure}}} ON COLUMNS "
        f"FROM {mdx_quote(cube)} "
        "CELL PROPERTIES VALUE, FORMATTED_VALUE"
    )


def build_calculated_mdx(
    target: dict,
    expression: str,
    year: str | None = None,
) -> tuple[str, str]:
    identifier = (
        "[Measures].[OIC_VERIFY_"
        + digest(target["target_key"])[:16]
        + "]"
    )

    statement = (
        f"WITH MEMBER {identifier} AS {expression}\n"
        f"SELECT {{{identifier}}} ON COLUMNS"
    )

    if year:
        if not valid_hierarchy(year):
            raise ResolutionError(
                "Invalid year hierarchy"
            )

        statement += (
            f", Subset({year}.AllMembers, 0, 4) "
            "ON ROWS"
        )

    statement += (
        f"\nFROM {mdx_quote(target['cube'])}\n"
        "CELL PROPERTIES VALUE, FORMATTED_VALUE"
    )

    return statement, identifier


def execute_native_measure(
    client: OICClient,
    target: dict,
    native: dict,
    year: str | None,
) -> dict:
    measure = native["mdx"]

    attempts = []

    contexts = (
        [year, None]
        if year
        else [None]
    )

    for context in contexts:
        statement = build_native_mdx(
            target["cube"],
            measure,
            context,
        )

        try:
            execution = client.execute(
                target["insurance_code"],
                target["catalog"],
                statement,
                measure,
                context,
            )

            attempts.append({
                "mdx": statement,
                "status": "SUCCESS",
                "context": (
                    "YEAR_SAMPLE"
                    if context
                    else "DEFAULT_CONTEXT"
                ),
            })

            return {
                "success": True,
                "statement": statement,
                "verification_context": (
                    "YEAR_SAMPLE"
                    if context
                    else "DEFAULT_CONTEXT"
                ),
                "execution": execution,
                "attempts": attempts,
            }

        except AuthenticationError:
            raise

        except (
            XMLAError,
            requests.RequestException,
        ) as exc:
            attempts.append({
                "mdx": statement,
                "status": "FAILED",
                "error": safe_error(
                    exc,
                    client.secrets(),
                ),
            })

    return {
        "success": False,
        "attempts": attempts,
    }


def execute_calculated_measure(
    client: OICClient,
    target: dict,
    calculated: dict,
    year: str | None,
) -> dict:
    attempts = []

    contexts = (
        [year, None]
        if year
        else [None]
    )

    for context in contexts:
        statement, measure = build_calculated_mdx(
            target,
            calculated["expression"],
            context,
        )

        try:
            execution = client.execute(
                target["insurance_code"],
                target["catalog"],
                statement,
                measure,
                context,
            )

            attempts.append({
                "mdx": statement,
                "status": "SUCCESS",
                "context": (
                    "YEAR_SAMPLE"
                    if context
                    else "DEFAULT_CONTEXT"
                ),
            })

            # A calculated formula must return at least
            # one nonempty cell before it is considered
            # successfully verified.

            if not execution["has_data"]:
                attempts[-1]["status"] = "EMPTY_RESULT"

                continue

            return {
                "success": True,
                "measure": measure,
                "statement": statement,
                "verification_context": (
                    "YEAR_SAMPLE"
                    if context
                    else "DEFAULT_CONTEXT"
                ),
                "execution": execution,
                "attempts": attempts,
            }

        except AuthenticationError:
            raise

        except (
            XMLAError,
            requests.RequestException,
        ) as exc:
            attempts.append({
                "mdx": statement,
                "status": "FAILED",
                "error": safe_error(
                    exc,
                    client.secrets(),
                ),
            })

    return {
        "success": False,
        "attempts": attempts,
    }


# ============================================================
# FIELD RESOLUTION ENGINE
# ============================================================

def create_result(
    target: dict,
) -> dict:
    return {
        **target,
        "status": "UNSUPPORTED",
        "resolved_mdx": None,
        "resolved_expression": None,
        "resolution_method": None,
        "mapping_source": None,
        "dependencies": [],
        "verification_context": None,
        "verification_query": None,
        "sample_cells": 0,
        "nonempty_cells": 0,
        "has_data": False,
        "verified_at": None,
        "checked_at": utc_now(),
        "reason": None,
        "error": None,
        "attempts": [],
        "diagnostics": {},
    }


def resolve_field(
    client: OICClient,
    target: dict,
    metadata: dict,
    measure_index: dict,
) -> dict:
    result = create_result(target)

    result["diagnostics"][
        "metadata_sources"
    ] = metadata["sources"]

    result["diagnostics"][
        "metadata_errors"
    ] = metadata["errors"]

    if target["field_type"] != "MEASURE":
        result["status"] = "UNVERIFIED"

        result["reason"] = (
            "Unsupported field is not a Measure; "
            "this resolver does not remap Dimensions"
        )

        return result

    year = resolve_year_hierarchy(
        target,
        metadata,
    )

    # --------------------------------------------------------
    # STRATEGY 1: EXISTING NATIVE / CUBE-CALCULATED MEASURE
    # --------------------------------------------------------

    native = resolve_native_measure(
        target,
        metadata,
        measure_index,
    )

    result["diagnostics"][
        "native_resolution"
    ] = native

    if native.get("mdx"):
        verified = execute_native_measure(
            client,
            target,
            native,
            year,
        )

        result["attempts"].extend(
            verified.get("attempts", [])
        )

        if verified["success"]:
            execution = verified["execution"]

            measure_metadata = (
                native.get("measure_metadata")
                or {}
            )

            is_cube_calculated = (
                clean(
                    measure_metadata.get(
                        "MEASURE_AGGREGATOR"
                    )
                ) == "127"
                or bool(
                    clean(
                        measure_metadata.get(
                            "EXPRESSION"
                        )
                    )
                )
            )

            result.update({
                "status": "VERIFIED_NATIVE",
                "resolved_mdx": native["mdx"],
                "resolution_method": (
                    "EXISTING_CUBE_CALCULATED_MEASURE"
                    if is_cube_calculated
                    else "NATIVE_MEASURE"
                ),
                "mapping_source": native[
                    "mapping_source"
                ],
                "verification_context": verified[
                    "verification_context"
                ],
                "verification_query": verified[
                    "statement"
                ],
                "sample_cells": execution[
                    "sample_cells"
                ],
                "nonempty_cells": execution[
                    "nonempty_cells"
                ],
                "has_data": execution[
                    "has_data"
                ],
                "verified_at": utc_now(),
                "reason": (
                    "Confirmed Measure identifier returned "
                    "on Axis0 with successful XMLA Execute"
                    if execution["has_data"]
                    else
                    "Measure identifier verified; "
                    "sample context returned no nonempty values"
                ),
                "error": None,
            })

            result["diagnostics"][
                "execution"
            ] = execution

            return result

    # --------------------------------------------------------
    # STRATEGY 2: DOCUMENTED CALCULATION EXPRESSION
    # --------------------------------------------------------

    calculated = resolve_calculated_measure(
        target,
        metadata,
        measure_index,
    )

    result["diagnostics"][
        "calculated_resolution"
    ] = calculated

    if calculated.get("resolved"):
        verified = execute_calculated_measure(
            client,
            target,
            calculated,
            year,
        )

        result["attempts"].extend(
            verified.get("attempts", [])
        )

        if verified["success"]:
            execution = verified["execution"]

            result.update({
                "status": "VERIFIED_CALCULATED",
                "resolved_mdx": verified["measure"],
                "resolved_expression": calculated[
                    "expression"
                ],
                "resolution_method": "WITH_MEMBER",
                "mapping_source": calculated[
                    "mapping_source"
                ],
                "dependencies": calculated[
                    "dependencies"
                ],
                "verification_context": verified[
                    "verification_context"
                ],
                "verification_query": verified[
                    "statement"
                ],
                "sample_cells": execution[
                    "sample_cells"
                ],
                "nonempty_cells": execution[
                    "nonempty_cells"
                ],
                "has_data": execution[
                    "has_data"
                ],
                "verified_at": utc_now(),
                "reason": (
                    "Documented formula successfully "
                    "executed using WITH MEMBER"
                ),
                "error": None,
            })

            result["diagnostics"][
                "execution"
            ] = execution

            return result

    # --------------------------------------------------------
    # STRATEGY 3: RETAIN UNRESOLVED EVIDENCE
    # --------------------------------------------------------

    if (
        native.get("ambiguous")
        or calculated.get("ambiguous")
    ):
        result["status"] = "UNVERIFIED"

        result["reason"] = (
            "Multiple possible mappings or formulas; "
            "manual confirmation required"
        )

    elif result["attempts"]:
        result["status"] = "UNVERIFIED"

        result["reason"] = (
            "Mapping or formula identified, "
            "but successful verification was not established"
        )

    elif metadata["errors"]:
        result["status"] = "UNVERIFIED"

        result["reason"] = (
            "Metadata discovery was incomplete; "
            "support cannot be conclusively determined"
        )

    else:
        result["status"] = "UNSUPPORTED"

        result["reason"] = (
            "No confirmed native Measure or safely "
            "convertible documented formula was found; "
            "this does not prove OIC cannot support the field"
        )

    failed_attempts = [
        item.get("error")
        for item in result["attempts"]
        if item.get("error")
    ]

    if failed_attempts:
        result["error"] = failed_attempts[-1]

    return result


# ============================================================
# CHECKPOINT / RESUME
# ============================================================

def load_checkpoint() -> dict:
    checkpoint = read_latest_json(
        CHECKPOINT_FILE,
        {},
    )

    if not isinstance(checkpoint, dict):
        checkpoint = {}

    if not isinstance(
        checkpoint.get("results"),
        dict,
    ):
        checkpoint["results"] = {}

    return checkpoint


def save_checkpoint(
    checkpoint: dict,
) -> Path:
    checkpoint["updated_at"] = utc_now()

    return atomic_json(
        CHECKPOINT_FILE,
        checkpoint,
    )


def can_reuse_result(
    target: dict,
    result: dict | None,
    *,
    force: bool = False,
) -> bool:
    if force:
        return False

    if not isinstance(result, dict):
        return False

    if result.get("status") not in RESOLVED_STATUSES:
        return False

    if (
        result.get("source_fingerprint")
        != target.get("source_fingerprint")
    ):
        return False

    if not result.get("verified_at"):
        return False

    return True


# ============================================================
# REPORT EXPORT
# ============================================================

def collect_results(
    targets: list[dict],
    checkpoint: dict,
) -> list[dict]:
    results = []

    saved = checkpoint.get(
        "results",
        {},
    )

    for target in targets:
        result = saved.get(
            target["target_key"]
        )

        if (
            isinstance(result, dict)
            and result.get("source_fingerprint")
            == target["source_fingerprint"]
        ):
            results.append(result)

        else:
            pending = create_result(
                target
            )

            pending["status"] = "PENDING"

            pending["reason"] = (
                "Field has not been processed"
            )

            results.append(pending)

    return results


def export_csv(
    results: list[dict],
) -> Path:
    path = OUTPUT_DIR / "all_fields.csv"

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

            for item in results:
                row = {
                    key: item.get(key)
                    for key in CSV_COLUMNS
                }

                for key, value in row.items():
                    if (
                        isinstance(value, str)
                        and value.lstrip().startswith(
                            ("=", "+", "-", "@")
                        )
                    ):
                        row[key] = "'" + value

                writer.writerow(row)

            stream.flush()
            os.fsync(stream.fileno())

        return commit_file(
            temporary,
            path,
        )

    finally:
        try:
            temporary.unlink(missing_ok=True)
        except OSError:
            pass


def build_summary(
    results: list[dict],
) -> dict:
    totals = Counter(
        item.get("status", "PENDING")
        for item in results
    )

    grouped = defaultdict(Counter)

    for result in results:
        key = (
            result["insurance_code"],
            result["cube_type"],
        )

        grouped[key][
            result["status"]
        ] += 1

    cubes = []

    for code, cube_type in sorted(grouped):
        counts = grouped[
            (code, cube_type)
        ]

        total = sum(counts.values())

        resolved = sum(
            counts[status]
            for status in RESOLVED_STATUSES
        )

        cubes.append({
            "insurance_code": code,
            "cube_type": cube_type,
            "total": total,
            "verified_native": counts[
                "VERIFIED_NATIVE"
            ],
            "verified_calculated": counts[
                "VERIFIED_CALCULATED"
            ],
            "unverified": counts[
                "UNVERIFIED"
            ],
            "unsupported": counts[
                "UNSUPPORTED"
            ],
            "error": counts["ERROR"],
            "pending": counts["PENDING"],
            "resolved": resolved,
            "resolution_percentage": round(
                100.0 * resolved / total,
                2,
            ) if total else 0.0,
        })

    resolved = sum(
        totals[status]
        for status in RESOLVED_STATUSES
    )

    total = len(results)

    return {
        "generated_at": utc_now(),
        "total_unsupported_targets": total,
        "resolved": resolved,
        "verified_native": totals[
            "VERIFIED_NATIVE"
        ],
        "verified_calculated": totals[
            "VERIFIED_CALCULATED"
        ],
        "unverified": totals[
            "UNVERIFIED"
        ],
        "unsupported": totals[
            "UNSUPPORTED"
        ],
        "error": totals["ERROR"],
        "pending": totals["PENDING"],
        "resolution_percentage": round(
            100.0 * resolved / total,
            2,
        ) if total else 0.0,
        "cubes": cubes,
        "original_reports_modified": False,
    }


def export_reports(
    targets: list[dict],
    checkpoint: dict,
) -> dict:
    results = collect_results(
        targets,
        checkpoint,
    )

    resolved = [
        item
        for item in results
        if item["status"] in RESOLVED_STATUSES
    ]

    unresolved = [
        item
        for item in results
        if item["status"] not in RESOLVED_STATUSES
    ]

    summary = build_summary(
        results
    )

    paths = {}

    paths["resolved"] = str(
        atomic_json(
            OUTPUT_DIR / "resolved_fields.json",
            {
                "generated_at": utc_now(),
                "total": len(resolved),
                "fields": resolved,
            },
        )
    )

    paths["unresolved"] = str(
        atomic_json(
            OUTPUT_DIR / "unresolved_fields.json",
            {
                "generated_at": utc_now(),
                "total": len(unresolved),
                "fields": unresolved,
            },
        )
    )

    paths["csv"] = str(
        export_csv(results)
    )

    summary["output_files"] = paths

    paths["summary"] = str(
        atomic_json(
            OUTPUT_DIR / "summary.json",
            summary,
        )
    )

    summary["output_files"] = paths

    return summary


# ============================================================
# SUMMARY TERMINAL
# ============================================================

def print_summary(
    summary: dict,
) -> None:
    if not RICH_AVAILABLE:
        print(
            json.dumps(
                summary,
                ensure_ascii=False,
                indent=2,
            )
        )

        return

    table = Table(
        title="OIC UNSUPPORTED VERIFICATION SUMMARY",
        expand=True,
    )

    table.add_column(
        "CUBE",
        style="cyan",
    )

    table.add_column(
        "TARGETS",
        justify="right",
    )

    table.add_column(
        "NATIVE",
        justify="right",
        style="green",
    )

    table.add_column(
        "CALCULATED",
        justify="right",
        style="green",
    )

    table.add_column(
        "UNVERIFIED",
        justify="right",
    )

    table.add_column(
        "UNSUPPORTED",
        justify="right",
        style="yellow",
    )

    table.add_column(
        "ERROR",
        justify="right",
        style="red",
    )

    table.add_column(
        "PENDING",
        justify="right",
    )

    table.add_column(
        "RESOLVED %",
        justify="right",
    )

    for item in summary["cubes"]:
        table.add_row(
            f"{item['insurance_code']}/{item['cube_type']}",
            str(item["total"]),
            str(item["verified_native"]),
            str(item["verified_calculated"]),
            str(item["unverified"]),
            str(item["unsupported"]),
            str(item["error"]),
            str(item["pending"]),
            f"{item['resolution_percentage']:.2f}%",
        )

    console.print(table)

    console.print(
        Panel(
            f"TARGETS: {summary['total_unsupported_targets']:,}\n"
            f"[green]RESOLVED: {summary['resolved']:,}[/]\n"
            f"VERIFIED NATIVE: "
            f"{summary['verified_native']:,}\n"
            f"VERIFIED CALCULATED: "
            f"{summary['verified_calculated']:,}\n"
            f"UNVERIFIED: "
            f"{summary['unverified']:,}\n"
            f"UNSUPPORTED: "
            f"{summary['unsupported']:,}\n"
            f"ERROR: {summary['error']:,}\n"
            f"PENDING: {summary['pending']:,}\n"
            f"\n[bold]RESOLUTION RATE: "
            f"{summary['resolution_percentage']:.2f}%[/]",
            title="RESULT",
            border_style="cyan",
        )
    )


# ============================================================
# COMMAND LINE
# ============================================================

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="oic_verify_unsupported.py",
        description=(
            "Resolve previously unsupported "
            "OIC Measures"
        ),
    )

    parser.add_argument(
        "--code",
        type=str.upper,
        choices=list(CATEGORIES),
        help="Select insurance category",
    )

    parser.add_argument(
        "--cube-type",
        type=str.upper,
        choices=list(CUBE_TYPES),
        help="Select FY, UY or AY",
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help="Re-verify previously resolved fields",
    )

    parser.add_argument(
        "--summary",
        action="store_true",
        help="Display saved results without network requests",
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=30,
        help="HTTP timeout in seconds",
    )

    parser.add_argument(
        "--delay",
        type=float,
        default=0.3,
        help="Delay between target fields",
    )

    args = parser.parse_args()

    if args.timeout < 1:
        parser.error(
            "--timeout must be greater than zero"
        )

    if args.delay < 0:
        parser.error(
            "--delay cannot be negative"
        )

    return args


# ============================================================
# APPLICATION
# ============================================================

def main() -> int:
    args = parse_args()

    print_header()

    if not SOURCE_DIR.is_dir():
        log(
            f"Source report directory not found: {SOURCE_DIR}",
            "ERROR",
        )

        log(
            "Run oic_verify_fields.py first.",
            "INFO",
        )

        return 2

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    METADATA_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    all_targets = load_unsupported_fields()

    targets = load_unsupported_fields(
        selected_code=args.code,
        selected_cube_type=args.cube_type,
    )

    checkpoint = load_checkpoint()

    if args.summary:
        results = collect_results(
            targets,
            checkpoint,
        )

        print_summary(
            build_summary(results)
        )

        return 0

    if not targets:
        log(
            "No UNSUPPORTED fields found "
            "for the selected categories.",
            "OK",
        )

        return 0

    pending = [
        target
        for target in targets
        if not can_reuse_result(
            target,
            checkpoint["results"].get(
                target["target_key"]
            ),
            force=args.force,
        )
    ]

    grouped = defaultdict(list)

    for target in pending:
        key = (
            target["insurance_code"],
            target["cube_type"],
            target["catalog"],
            target["cube"],
        )

        grouped[key].append(target)

    log(
        f"UNSUPPORTED TARGETS={len(targets)} | "
        f"PREVIOUSLY RESOLVED={len(targets) - len(pending)} | "
        f"PENDING={len(pending)} | "
        f"CUBES={len(grouped)}"
    )

    if not pending:
        summary = export_reports(
            all_targets,
            checkpoint,
        )

        print_summary(summary)

        return 0

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

    client = OICClient(
        cookie=cookie,
        token=token,
        timeout=args.timeout,
    )

    interrupted = False
    authentication_failed = False
    fatal_error = False

    metadata_cache = {}

    try:
        # ----------------------------------------------------
        # AUTHENTICATION PREFLIGHT
        # ----------------------------------------------------

        first = pending[0]

        first_key = (
            first["insurance_code"],
            first["cube_type"],
            first["catalog"],
            first["cube"],
        )

        log(
            "Checking OIC authentication..."
        )

        try:
            probe = client.discover(
                first["insurance_code"],
                first["catalog"],
                first["cube"],
                "MDSCHEMA_MEASURES",
            )

            if not probe:
                log(
                    "Authentication request succeeded, "
                    "but no Measures were returned.",
                    "WARN",
                )
            else:
                log(
                    f"AUTHENTICATION OK | "
                    f"MEASURES={len(probe)}",
                    "OK",
                )

        except AuthenticationError:
            raise

        except Exception as exc:
            log(
                "Authentication preflight failed: "
                + safe_error(
                    exc,
                    client.secrets(),
                ),
                "ERROR",
            )

            fatal_error = True

            raise ResolutionError(
                "Cannot establish a valid XMLA connection"
            ) from exc

        # ----------------------------------------------------
        # PROCESS TARGET GROUPS
        # ----------------------------------------------------

        def process_group(
            key: tuple,
            items: list[dict],
            progress=None,
            task_id=None,
        ) -> None:
            code, cube_type, _, _ = key

            try:
                metadata = metadata_cache.get(key)

                if metadata is None:
                    metadata = discover_cube_metadata(
                        client,
                        items[0],
                    )

                    metadata_cache[key] = metadata

                measure_index = build_measure_index(
                    metadata["measures"]
                )

                log(
                    f"{code}/{cube_type} | "
                    f"TARGETS={len(items)} | "
                    f"NATIVE MEASURES="
                    f"{len(measure_index['by_mdx'])}"
                )

            except AuthenticationError:
                raise

            except Exception as exc:
                error = safe_error(
                    exc,
                    client.secrets(),
                )

                log(
                    f"{code}/{cube_type}: "
                    f"METADATA ERROR | {error}",
                    "ERROR",
                )

                for target in items:
                    previous = checkpoint[
                        "results"
                    ].get(
                        target["target_key"]
                    )

                    # Preserve successful prior results if
                    # a forced refresh encounters an error.

                    if can_reuse_result(
                        target,
                        previous,
                    ):
                        continue

                    result = create_result(
                        target
                    )

                    result["status"] = "ERROR"
                    result["error"] = error
                    result["reason"] = (
                        "Live metadata discovery failed"
                    )

                    checkpoint["results"][
                        target["target_key"]
                    ] = result

                save_checkpoint(checkpoint)

                if progress is not None:
                    progress.advance(
                        task_id,
                        len(items),
                    )

                return

            for target in items:
                caption = target[
                    "field_caption"
                ]

                log(
                    f"{code}/{cube_type}: "
                    f"RESOLVING {caption}"
                )

                try:
                    result = resolve_field(
                        client,
                        target,
                        metadata,
                        measure_index,
                    )

                except AuthenticationError:
                    raise

                except Exception as exc:
                    result = create_result(
                        target
                    )

                    result["status"] = "ERROR"

                    result["error"] = safe_error(
                        exc,
                        client.secrets(),
                    )

                    result["reason"] = (
                        "Unexpected resolution error"
                    )

                previous = checkpoint[
                    "results"
                ].get(
                    target["target_key"]
                )

                # Do not discard a successful result
                # because a later forced attempt failed.

                if (
                    result["status"]
                    not in RESOLVED_STATUSES
                    and can_reuse_result(
                        target,
                        previous,
                    )
                ):
                    previous["last_recheck"] = {
                        "checked_at": utc_now(),
                        "status": result["status"],
                        "reason": result["reason"],
                        "error": result["error"],
                    }

                else:
                    checkpoint["results"][
                        target["target_key"]
                    ] = result

                save_checkpoint(
                    checkpoint
                )

                level = (
                    "OK"
                    if result["status"] in RESOLVED_STATUSES
                    else (
                        "ERROR"
                        if result["status"] == "ERROR"
                        else "WARN"
                    )
                )

                log(
                    f"{code}/{cube_type}: "
                    f"{result['status']} | "
                    f"{caption}",
                    level,
                )

                if result.get("resolved_mdx"):
                    log(
                        f"MDX: {result['resolved_mdx']}"
                    )

                if result.get("resolved_expression"):
                    log(
                        "FORMULA: "
                        + result["resolved_expression"]
                    )

                if result.get("reason"):
                    log(
                        f"REASON: {result['reason']}"
                    )

                if result.get("error"):
                    log(
                        result["error"],
                        "WARN",
                    )

                if progress is not None:
                    progress.advance(
                        task_id,
                        1,
                    )

                if args.delay > 0:
                    time.sleep(args.delay)

        # ----------------------------------------------------
        # TERMINAL PROGRESS
        # ----------------------------------------------------

        if RICH_AVAILABLE:
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

                task_id = progress.add_task(
                    "RESOLVING UNSUPPORTED FIELDS",
                    total=len(pending),
                )

                for key, items in grouped.items():
                    process_group(
                        key,
                        items,
                        progress=progress,
                        task_id=task_id,
                    )

        else:
            for key, items in grouped.items():
                process_group(
                    key,
                    items,
                )

    except AuthenticationError as exc:
        authentication_failed = True

        log(
            safe_error(
                exc,
                client.secrets(),
            ),
            "ERROR",
        )

        log(
            "Update OIC_COOKIE and OIC_BEARER_TOKEN "
            "in .env, then rerun this script.",
            "WARN",
        )

        log(
            "Previously saved results are preserved.",
            "INFO",
        )

    except KeyboardInterrupt:
        interrupted = True

        log(
            "STOP requested. Saving current progress.",
            "WARN",
        )

    except Exception as exc:
        fatal_error = True

        log(
            safe_error(
                exc,
                client.secrets(),
            ),
            "ERROR",
        )

    finally:
        client.close()

        try:
            save_checkpoint(
                checkpoint
            )

            summary = export_reports(
                all_targets,
                checkpoint,
            )

            print_summary(summary)

            for name, path in summary.get(
                "output_files",
                {},
            ).items():
                log(
                    f"{name.upper()}: {path}"
                )

            log(
                "Original field_verification reports "
                "were not modified.",
                "OK",
            )

        except OSError as exc:
            fatal_error = True

            log(
                "REPORT SAVE ERROR: "
                + safe_error(exc),
                "ERROR",
            )

    if interrupted:
        return 130

    if authentication_failed:
        return 3

    if fatal_error:
        return 1

    selected_results = collect_results(
        targets,
        checkpoint,
    )

    errors = sum(
        result["status"] == "ERROR"
        for result in selected_results
    )

    remaining = sum(
        result["status"] not in RESOLVED_STATUSES
        for result in selected_results
    )

    if errors:
        log(
            f"Completed with {errors} errors.",
            "WARN",
        )

        return 1

    if remaining:
        log(
            f"Completed | "
            f"UNRESOLVED={remaining} | "
            "See unresolved_fields.json for evidence.",
            "WARN",
        )

        return 0

    log(
        "All selected UNSUPPORTED fields resolved.",
        "OK",
    )

    return 0


if __name__ == "__main__":
    sys.exit(main())
