#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC PARALLEL COLLECTOR
======================

XMLA / MDX / OLAP Aggregate Collector
Rich Terminal Monitor + native PowerShell processes
4 Parallel Workers + Raw XML + Excel + Resume

INSTALL
-------
python -m pip install requests openpyxl rich

COMMANDS
--------
Prepare jobs:
    python oic_parallel_collector.py --prepare

Start 4 background workers with a native PowerShell monitor:
    python oic_parallel_collector.py

Or:
    python oic_parallel_collector.py --launch

The original PowerShell displays a LIVE RICH MONITOR.
A PowerShell monitor opens; worker processes run in the background.

Test one job:
    python oic_parallel_collector.py --worker-id 1 --max-jobs 1 --max-years 1

Run one worker manually:
    python oic_parallel_collector.py --worker-id 1

Check progress:
    python oic_parallel_collector.py --status

Merge manifests:
    python oic_parallel_collector.py --merge

Rebuild Excel:
    python oic_parallel_collector.py --export-only

Resume:
    python oic_parallel_collector.py

EXTRACTION
----------
COLUMNS = YEAR only
ROWS    = Dimension x Measures

Each worker:
    - Own HTTP session
    - Own JWT
    - Own dimensions
    - Own manifest
    - Own raw XML directory

Default partition:
    W01 = 1, 5, 9, ...
    W02 = 2, 6, 10, ...
    W03 = 3, 7, 11, ...
    W04 = 4, 8, 12, ...

Default chunk:
    1 YEAR x 20 Dimension Members

OUTPUT
------
output/oic_parallel/MV/
    job_plan.json
    master_manifest.json

    worker_01/
        manifest.json
        state.json
        dim_001/
            y000_m000000.request.xml
            y000_m000000.response.xml

    worker_02/
    worker_03/
    worker_04/

    excel/
        hierarchy_001.xlsx
        hierarchy_002.xlsx
        ...

REQUIREMENTS
------------
Windows 10/11
PowerShell 5.1 or 7
Python 3.10+
requests
openpyxl
rich

.env:
    OIC_COOKIE=...
    OIC_BEARER_TOKEN=...
"""

from __future__ import annotations

import argparse
import csv
import base64
import json
import math
import os
import re
import shutil
import subprocess
import sys
import time
import uuid
import threading

from collections import OrderedDict
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from xml.etree import ElementTree as ET

import requests

from rich.align import Align
from rich.console import Console, Group
from rich.live import Live
from rich.panel import Panel
from rich.table import Table
from rich.text import Text
from rich import box as rich_box

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None


# ============================================================
# CONFIGURATION
# ============================================================

APP_NAME = "OIC PARALLEL COLLECTOR"
APP_VERSION = "3.0.0"

HERE = Path(__file__).resolve().parent

DEFAULT_SOURCE = (
    HERE
    / "output"
    / "xmla_expand"
    / "MV"
    / "manifest.json"
)

DEFAULT_OUTPUT = (
    HERE
    / "output"
    / "oic_parallel"
    / "MV"
)

BASE_URL = "https://onlinewebibs.oic.or.th"

XMLA_URL = (
    BASE_URL
    + "/xmla-proxy/api/xmla-proxy"
)

TOKEN_URL = (
    BASE_URL
    + "/WebForm/Dashboard/Pivot/"
    "PivotCube.asmx/GenToken"
)

PAGE_PATH = (
    "/WebForm/Dashboard/Pivot/"
    "DynamicPivotCube.aspx"
)

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"

XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"

MD_NS = (
    "urn:schemas-microsoft-com:"
    "xml-analysis:mddataset"
)

ET.register_namespace("soap", SOAP_NS)


# ============================================================
# VERIFIED MV CUBE
# ============================================================

CUBE = {
    "code": "MV",
    "class": "MO",
    "catalog": "PIVOT_202609_MV",
    "cube": "MV_AGGREGATE_FY",
    "year": "[VW PV MV AGGREGATE FY].[YEAR]",
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
}


# ============================================================
# DEFAULTS
# ============================================================

DEFAULT_WORKERS = 4

DEFAULT_YEARS_PER_REQUEST = 1
DEFAULT_MEMBERS_PER_REQUEST = 20

DEFAULT_TIMEOUT = 60
DEFAULT_DELAY = 1.0

DEFAULT_MAX_RESPONSE_MB = 100
DEFAULT_MAX_REQUESTS_PER_JOB = 5000

MONITOR_REFRESH_SECONDS = 1.0

EXCEL_MAX_ROWS = 1_048_576

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)

from oic_query import PROFILES as QUERY_PROFILES, validate_query, build_soap as query_soap, parse_xmla as query_parse, excel_records as query_excel_records, query_hash

console = Console(highlight=False)


# ============================================================
# GENERAL UTILITIES
# ============================================================

def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(
        timespec="seconds"
    )


def local_time() -> str:
    return datetime.now().strftime("%H:%M:%S")


def log(
    message: str,
    worker_id: int | None = None,
    level: str = "INFO",
) -> None:

    prefix = (
        f"W{worker_id:02d}"
        if worker_id is not None
        else "MAIN"
    )

    colors = {
        "INFO": "cyan",
        "OK": "green",
        "WARN": "yellow",
        "ERROR": "red",
    }

    color = colors.get(level, "white")

    console.print(
        f"[dim]{local_time()}[/dim] "
        f"[bold]{prefix}[/bold] "
        f"[{color}]{message}[/{color}]"
    )


def read_json(
    path: Path,
    default: Any = None,
) -> Any:

    if not path.exists():
        return default

    try:
        return json.loads(
            path.read_text(encoding="utf-8")
        )

    except (
        OSError,
        UnicodeError,
        json.JSONDecodeError,
    ):
        return default


def write_json(
    path: Path,
    data: Any,
) -> None:

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temporary.open("w", encoding="utf-8") as stream:
            json.dump(data, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
            stream.flush()
            os.fsync(stream.fileno())
        for attempt in range(8):
            try:
                os.replace(temporary, path)
                break
            except PermissionError:
                if attempt == 7:
                    raise
                time.sleep(0.05 * (attempt + 1))
    finally:
        temporary.unlink(missing_ok=True)


def read_env() -> dict[str, str]:

    values = dict(os.environ)

    path = HERE / ".env"

    if not path.exists():
        return values

    for line in path.read_text(
        encoding="utf-8-sig"
    ).splitlines():

        line = line.strip()

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
            and value[0] in ('"', "'")
        ):
            value = value[1:-1]

        values[key] = value

    return values


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

        return int(data.get("exp", 0))

    except (
        ValueError,
        IndexError,
        TypeError,
    ):
        return 0


def validate_hierarchy(name: str) -> str:

    name = name.strip()

    pattern = (
        r"\[[^\]\r\n]+\]"
        r"\."
        r"\[[^\]\r\n]+\]"
    )

    if not re.fullmatch(pattern, name):
        raise ValueError(
            f"Invalid MDX hierarchy: {name}"
        )

    return name


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def worker_dir(
    output: Path,
    worker_id: int,
) -> Path:

    return (
        output
        / f"worker_{worker_id:02d}"
    )


def dimension_dir(
    output: Path,
    worker_id: int,
    job_id: int,
) -> Path:

    return (
        worker_dir(output, worker_id)
        / f"dim_{job_id:03d}"
    )


def chunk_key(
    job_id: int,
    year_offset: int,
    member_offset: int,
) -> str:

    return (
        f"{job_id}:"
        f"{year_offset}:"
        f"{member_offset}"
    )


def chunk_name(
    year_offset: int,
    member_offset: int,
) -> str:

    return (
        f"y{year_offset:03d}_"
        f"m{member_offset:06d}"
    )


def parse_time(value: str | None) -> datetime | None:

    if not value:
        return None

    try:
        return datetime.fromisoformat(
            value.replace("Z", "+00:00")
        )

    except ValueError:
        return None


def elapsed_seconds(value: str | None) -> int:

    started = parse_time(value)

    if started is None:
        return 0

    return max(
        0,
        int(
            (
                datetime.now(timezone.utc)
                - started
            ).total_seconds()
        ),
    )


def format_duration(seconds: int) -> str:

    hours, remainder = divmod(
        max(0, seconds),
        3600,
    )

    minutes, seconds = divmod(
        remainder,
        60,
    )

    if hours:
        return (
            f"{hours:02d}:"
            f"{minutes:02d}:"
            f"{seconds:02d}"
        )

    return f"{minutes:02d}:{seconds:02d}"


# ============================================================
# HIERARCHY DISCOVERY
# ============================================================

def discover_xml_hierarchies(
    path: Path,
) -> list[str]:

    try:
        root = ET.parse(path).getroot()

    except (
        ET.ParseError,
        OSError,
    ):
        return []

    found = []
    seen = set()

    xpath = (
        f".//{{{MD_NS}}}AxesInfo/"
        f"{{{MD_NS}}}AxisInfo/"
        f"{{{MD_NS}}}HierarchyInfo"
    )

    for element in root.findall(xpath):

        name = element.get("name", "")

        if (
            not name
            or name == "[Measures]"
            or name in seen
        ):
            continue

        try:
            validate_hierarchy(name)

        except ValueError:
            continue

        found.append(name)
        seen.add(name)

    return found


def discover_hierarchies(
    source: Path,
) -> list[str]:

    data = read_json(source, {})

    candidates = []

    if isinstance(data, dict):

        candidates.extend(
            data.get("discovered_hierarchies")
            or []
        )

        for request in data.get(
            "requests",
            [],
        ):

            if not isinstance(request, dict):
                continue

            result = request.get("result") or {}

            candidates.extend(
                result.get("discovered_hierarchies")
                or []
            )

    if not candidates:

        for path in sorted(
            source.parent.glob("*.response.xml")
        ):

            candidates = discover_xml_hierarchies(
                path
            )

            if candidates:
                break

    result = []
    seen = set()

    for name in candidates:

        if not isinstance(name, str):
            continue

        name = name.strip()

        if (
            not name
            or name == "[Measures]"
            or name == CUBE["year"]
            or name in seen
        ):
            continue

        try:
            validate_hierarchy(name)

        except ValueError:
            continue

        result.append(name)
        seen.add(name)

    return result


# ============================================================
# JOB PLANNING
# ============================================================

def prepare_plan(
    output: Path,
    source: Path,
    workers: int,
) -> dict:

    output.mkdir(
        parents=True,
        exist_ok=True,
    )

    path = output / "job_plan.json"

    existing = read_json(path)

    if isinstance(existing, dict):

        if existing.get("workers") != workers:
            raise RuntimeError(
                "Existing job plan uses "
                f"{existing.get('workers')} workers. "
                "Keep the same --workers value."
            )

        return existing

    dimensions = discover_hierarchies(
        source
    )

    if not dimensions:
        raise RuntimeError(
            "No hierarchies found. "
            "Verify the Expand Probe manifest: "
            f"{source}"
        )

    jobs = []

    for index, dimension in enumerate(
        dimensions,
        start=1,
    ):

        worker_id = (
            (index - 1) % workers
        ) + 1

        jobs.append({
            "id": index,
            "worker_id": worker_id,
            "hierarchy": dimension,
        })

    plan = {
        "created_at": utc_now(),
        "code": CUBE["code"],
        "cube": CUBE["cube"],
        "catalog": CUBE["catalog"],
        "year_hierarchy": CUBE["year"],
        "measures": CUBE["measures"],
        "workers": workers,
        "total_hierarchies": len(jobs),
        "source_manifest": str(source),
        "partition_method": "round_robin",
        "jobs": jobs,
    }

    write_json(path, plan)

    return plan


def print_plan(plan: dict) -> None:

    table = Table(
        title="OIC JOB DISTRIBUTION",
        box=rich_box.ROUNDED,
        header_style="bold cyan",
    )

    table.add_column("Worker")
    table.add_column("Jobs", justify="right")
    table.add_column("Job IDs")

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        jobs = [
            job["id"]
            for job in plan["jobs"]
            if job["worker_id"] == worker_id
        ]

        table.add_row(
            f"W{worker_id:02d}",
            str(len(jobs)),
            ", ".join(map(str, jobs)),
        )

    console.print(table)

    console.print(
        f"[bold]Total Hierarchies:[/bold] "
        f"{plan['total_hierarchies']}"
    )


# ============================================================
# MDX BUILDER
# ============================================================

def build_mdx(
    dimension: str,
    year_offset: int,
    member_offset: int,
    years_per_request: int,
    members_per_request: int,
) -> str:

    year = validate_hierarchy(
        CUBE["year"]
    )

    dimension = validate_hierarchy(
        dimension
    )

    measures = ", ".join(
        f"[Measures].[{name}]"
        for name in CUBE["measures"]
    )

    # YEAR: horizontal axis only.
    year_set = (
        f"Subset("
        f"{year}.DefaultMember.Children, "
        f"{year_offset}, "
        f"{years_per_request}"
        f")"
    )

    # Expand one hierarchy at a time.
    member_set = (
        f"Subset("
        f"{dimension}.AllMembers, "
        f"{member_offset}, "
        f"{members_per_request}"
        f")"
    )

    # Measures are placed on ROWS with the dimension.
    row_set = (
        f"CrossJoin("
        f"{member_set}, "
        f"{{{measures}}}"
        f")"
    )

    return (
        "SELECT "
        f"{year_set} ON COLUMNS, "
        f"{row_set} ON ROWS "
        f"FROM [{CUBE['cube']}] "
        "CELL PROPERTIES VALUE, FORMAT_STRING, LANGUAGE"
    )


# ============================================================
# SOAP BUILDER
# ============================================================

def build_soap(mdx: str) -> bytes:

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
        execute,
        f"{{{XMLA_NS}}}Properties",
    )

    property_list = ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}PropertyList",
    )

    values = {
        "Catalog": CUBE["catalog"],
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
# XMLA RESPONSE PARSER
# ============================================================

def parse_xmla(
    path: Path,
    dimension: str,
    include_data: bool = False,
) -> dict:

    result = {
        "valid": False,
        "cube_name": None,
        "year_count": 0,
        "member_count": 0,
        "row_tuple_count": 0,
        "nonempty_cells": 0,
        "soap_fault": None,
        "error": None,
    }

    try:
        root = ET.parse(path).getroot()

    except (
        ET.ParseError,
        OSError,
    ) as exc:

        result["error"] = str(exc)

        if path.exists():

            result["preview"] = (
                path.read_bytes()[:250].decode(
                    "utf-8",
                    errors="replace",
                )
            )

        return result

    for element in root.iter():

        tag = local_name(element.tag)

        if tag == "Fault":

            result["soap_fault"] = " ".join(
                text.strip()
                for text in element.itertext()
                if text.strip()
            )[:1000]

        elif tag == "CubeName" and element.text:

            result["cube_name"] = element.text

    if result["soap_fault"]:

        result["error"] = result["soap_fault"]
        return result

    dataset = root.find(
        f".//{{{MD_NS}}}root"
    )

    if dataset is None:

        result["error"] = (
            "XMLA multidimensional dataset not found"
        )

        return result

    axes = {}

    for axis in dataset.findall(
        f"./{{{MD_NS}}}Axes/{{{MD_NS}}}Axis"
    ):

        name = axis.get("name", "")

        tuples = []

        for item in axis.findall(
            f"{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple"
        ):

            members = []

            for member in item.findall(
                f"{{{MD_NS}}}Member"
            ):

                members.append({
                    "hierarchy": member.get(
                        "Hierarchy",
                        "",
                    ),
                    "unique_name": (
                        member.findtext(
                            f"{{{MD_NS}}}UName"
                        )
                        or ""
                    ),
                    "caption": (
                        member.findtext(
                            f"{{{MD_NS}}}Caption"
                        )
                        or ""
                    ),
                })

            tuples.append(members)
        axes[name] = tuples

    if (
        "Axis0" not in axes
        or "Axis1" not in axes
    ):

        result["error"] = (
            "Expected Axis0 YEAR and Axis1 ROWS"
        )

        return result

    years = []

    for item in axes["Axis0"]:

        year = next(
            (
                member
                for member in item
                if member["hierarchy"] == CUBE["year"]
            ),
            None,
        )

        if year is None:

            result["error"] = (
                "YEAR missing from Axis0"
            )

            return result

        years.append(year)

    row_descriptors = []

    unique_members = OrderedDict()

    for item in axes["Axis1"]:

        member = next(
            (
                entry
                for entry in item
                if entry["hierarchy"] == dimension
            ),
            None,
        )

        measure = next(
            (
                entry
                for entry in item
                if entry["hierarchy"] == "[Measures]"
            ),
            None,
        )

        if member is None or measure is None:

            result["error"] = (
                "Dimension or Measure missing from Axis1"
            )

            return result

        unique_members[
            member["unique_name"]
        ] = member

        row_descriptors.append({
            "member": member,
            "measure": measure,
        })

    cells = {}

    for cell in dataset.findall(
        f"./{{{MD_NS}}}CellData/"
        f"{{{MD_NS}}}Cell"
    ):

        value_element = cell.find(
            f"{{{MD_NS}}}Value"
        )

        if value_element is None:
            continue

        try:
            ordinal = int(
                cell.get("CellOrdinal", "")
            )

        except ValueError:
            continue

        cells[ordinal] = value_element.text

    result.update({
        "valid": True,
        "year_count": len(years),
        "member_count": len(unique_members),
        "row_tuple_count": len(row_descriptors),
        "nonempty_cells": len(cells),
    })

    if include_data:

        result["years"] = years

        result["row_descriptors"] = (
            row_descriptors
        )

        result["cells"] = cells

    return result


# ============================================================
# HTTP CLIENT
# ============================================================

class OICClient:

    def __init__(
        self,
        cookie: str,
        token: str,
        timeout: int,
        max_response_mb: int,
        worker_id: int,
        code: str = "MV",
    ):

        self.profile = QUERY_PROFILES[code]
        self.worker_id = worker_id

        self.token = token.removeprefix(
            "Bearer "
        ).strip()

        self.timeout = timeout

        self.max_bytes = (
            max_response_mb * 1024 * 1024
        )

        self.session = requests.Session()

        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Accept-Language": "th-TH,th;q=0.9",
        })

        hostname = urlparse(
            BASE_URL
        ).hostname

        cookie = cookie.removeprefix(
            "Cookie:"
        ).strip()

        for part in cookie.split(";"):

            name, separator, value = (
                part.strip().partition("=")
            )

            if separator and name:

                self.session.cookies.set(
                    name.strip(),
                    value.strip(),
                    domain=hostname,
                    path="/",
                    secure=True,
                )

    def referer(self) -> str:

        return (
            f"{BASE_URL}{PAGE_PATH}"
            f"?Class={self.profile['cls']}"
        )

    def refresh_token(self) -> None:

        log(
            "Refreshing Bearer JWT via GenToken",
            self.worker_id,
        )

        payload = {
            "_tk": "Init",
            "_pageUrl": (
                f"{PAGE_PATH}"
                f"?Class={self.profile['cls']}"
            ),
        }

        response = self.session.post(
            TOKEN_URL,
            json=payload,
            headers={
                "Accept": (
                    "application/json, text/javascript, "
                    "*/*; q=0.01"
                ),
                "Content-Type": (
                    "application/json; charset=UTF-8"
                ),
                "X-Requested-With": "XMLHttpRequest",
                "Origin": BASE_URL,
                "Referer": self.referer(),
            },
            timeout=(15, 45),
            allow_redirects=False,
        )

        if response.status_code != 200:

            raise RuntimeError(
                f"GenToken HTTP {response.status_code}"
            )

        try:
            token = response.json().get(
                "d",
                "",
            )

        except ValueError as exc:

            raise RuntimeError(
                "Invalid GenToken JSON response"
            ) from exc

        if token == "SESSION_EXPIRED":

            raise RuntimeError(
                "Cookie session expired"
            )

        if not isinstance(token, str):

            raise RuntimeError(
                "GenToken did not return a JWT"
            )

        token = token.removeprefix(
            "Bearer "
        ).strip()

        if (
            jwt_exp(token)
            <= int(time.time()) + 30
        ):

            raise RuntimeError(
                "GenToken returned expired or invalid JWT"
            )

        self.token = token

        log(
            "Bearer JWT refreshed successfully",
            self.worker_id,
            "OK",
        )

    def execute(
        self,
        request_body: bytes,
        destination: Path,
    ) -> dict:

        if (
            jwt_exp(self.token)
            <= int(time.time()) + 60
        ):

            self.refresh_token()

        for attempt in range(2):

            headers = {
                "Accept": "text/plain, */*; q=0.01",
                "Content-Type": "text/xml",
                "X-Requested-With": "XMLHttpRequest",
                "Origin": BASE_URL,
                "Referer": self.referer(),
                "Authorization": (
                    "Bearer " + self.token
                ),
            }

            with self.session.post(
                XMLA_URL,
                data=request_body,
                headers=headers,
                timeout=(15, self.timeout),
                stream=True,
                allow_redirects=False,
            ) as response:

                if (
                    response.status_code in (401, 403)
                    and attempt == 0
                ):

                    self.refresh_token()
                    continue

                temporary = destination.with_name(
                    destination.name + ".partial"
                )

                size = 0

                try:

                    with temporary.open("wb") as file:

                        for chunk in response.iter_content(
                            chunk_size=256 * 1024
                        ):

                            if not chunk:
                                continue

                            size += len(chunk)

                            if size > self.max_bytes:

                                raise RuntimeError(
                                    "XMLA response exceeds "
                                    "configured size limit"
                                )

                            file.write(chunk)

                    temporary.replace(
                        destination
                    )

                finally:

                    temporary.unlink(
                        missing_ok=True
                    )

                return {
                    "http_status": response.status_code,
                    "content_type": response.headers.get(
                        "Content-Type",
                        "",
                    ),
                    "response_bytes": size,
                    "retry_after": response.headers.get("Retry-After", ""),
                }

        raise RuntimeError(
            "Authentication rejected after JWT refresh"
        )

    def close(self) -> None:
        self.session.close()


# ============================================================
# EXCEL UTILITIES
# ============================================================

def numeric_value(
    raw: str | None,
) -> Any:

    if raw is None:
        return None

    raw = raw.strip()

    if not raw:
        return None

    try:
        number = float(raw)

        if not math.isfinite(number):
            return raw

        if (
            number.is_integer()
            and abs(number) < 1e15
        ):
            return int(number)

        return number

    except ValueError:
        return raw


def excel_text(value: str) -> str:

    if value.startswith(
        ("=", "+", "-", "@")
    ):
        return "'" + value

    return value


def measure_name(
    unique_name: str,
) -> str:

    prefix = "[Measures].["

    if (
        unique_name.startswith(prefix)
        and unique_name.endswith("]")
    ):

        return unique_name[
            len(prefix):-1
        ]

    return unique_name


def iter_excel_rows(
    response_file: Path,
    dimension: str,
):

    result = parse_xmla(
        response_file,
        dimension,
        include_data=True,
    )

    if not result["valid"]:

        raise RuntimeError(
            f"Invalid XML: {response_file}"
        )

    years = result["years"]

    descriptors = result[
        "row_descriptors"
    ]

    cells = result["cells"]

    year_count = len(years)

    if year_count == 0:
        return

    grouped = OrderedDict()

    for row_index, descriptor in enumerate(
        descriptors
    ):

        member = descriptor["member"]

        measure = descriptor["measure"]

        measure_key = measure_name(
            measure["unique_name"]
        )

        for year_index, year in enumerate(
            years
        ):

            ordinal = (
                year_index
                + year_count * row_index
            )

            key = (
                year["unique_name"],
                member["unique_name"],
            )

            if key not in grouped:

                grouped[key] = {
                    "year": year["caption"],
                    "year_unique": year["unique_name"],
                    "member": member["caption"],
                    "member_unique": member["unique_name"],
                    "measures": {},
                }

            if ordinal in cells:

                grouped[key]["measures"][
                    measure_key
                ] = numeric_value(
                    cells[ordinal]
                )

    for item in grouped.values():

        if not item["measures"]:
            continue

        yield [
            excel_text(item["year"]),
            excel_text(item["year_unique"]),
            excel_text(dimension),
            excel_text(item["member"]),
            excel_text(item["member_unique"]),
            *[
                item["measures"].get(name)
                for name in CUBE["measures"]
            ],
        ]


# ============================================================
# EXCEL EXPORT
# ============================================================

def export_hierarchy(
    output: Path,
    job: dict,
    chunks: list[dict],
    status: str,
) -> tuple[Path | None, int]:

    if Workbook is None:

        raise RuntimeError(
            "openpyxl is missing. "
            "Install: pip install openpyxl"
        )

    successful = [
        record
        for record in chunks
        if record.get("status") == "OK"
    ]

    if not successful:
        return None, 0

    successful.sort(
        key=lambda record: (
            record["year_offset"],
            record["member_offset"],
        )
    )

    excel_dir = output / "excel"

    excel_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    destination = (
        excel_dir
        / f"hierarchy_{job['id']:03d}.xlsx"
    )

    temporary = destination.with_name(
        destination.stem + ".tmp.xlsx"
    )

    workbook = Workbook(
        write_only=True
    )

    headers = [
        "YEAR",
        "YEAR_UNIQUE_NAME",
        "HIERARCHY",
        "MEMBER",
        "MEMBER_UNIQUE_NAME",
        *CUBE["measures"],
    ]

    sheet_index = 1

    sheet = workbook.create_sheet(
        title="Data_1"
    )

    sheet.append(headers)

    sheet_rows = 1
    total_rows = 0

    directory = dimension_dir(
        output,
        job["worker_id"],
        job["id"],
    )

    for record in successful:

        response_file = (
            directory
            / record["response_file"]
        )

        if not response_file.exists():
            continue

        for row in iter_excel_rows(
            response_file,
            job["hierarchy"],
        ):

            if sheet_rows >= EXCEL_MAX_ROWS:

                sheet_index += 1

                sheet = workbook.create_sheet(
                    title=f"Data_{sheet_index}"
                )

                sheet.append(headers)

                sheet_rows = 1

            sheet.append(row)

            sheet_rows += 1
            total_rows += 1

    info = workbook.create_sheet(
        title="INFO"
    )

    info_rows = [
        ("JOB_ID", job["id"]),
        ("WORKER_ID", job["worker_id"]),
        ("CUBE", CUBE["cube"]),
        ("CATALOG", CUBE["catalog"]),
        ("HIERARCHY", job["hierarchy"]),
        ("YEAR_AXIS", CUBE["year"]),
        ("STATUS", status),
        ("EXPORTED_ROWS", total_rows),
        ("SUCCESSFUL_CHUNKS", len(successful)),
        (
            "DATA_TYPE",
            "OLAP AGGREGATE - NOT RAW TRANSACTIONS",
        ),
        (
            "NOTE",
            "Independent hierarchy extraction",
        ),
        (
            "NOTE",
            "All members must not be summed with detail",
        ),
    ]

    if status != "COMPLETED":

        info_rows.append((
            "WARNING",
            "PARTIAL OR INCOMPLETE EXTRACTION",
        ))

    for key, value in info_rows:
        info.append([key, value])

    try:

        workbook.save(
            temporary
        )

        temporary.replace(
            destination
        )

    finally:

        temporary.unlink(
            missing_ok=True
        )

    return destination, total_rows


# ============================================================
# WORKER MANIFEST
# ============================================================

def load_worker_manifest(
    output: Path,
    worker_id: int,
) -> dict:

    path = (
        worker_dir(output, worker_id)
        / "manifest.json"
    )

    data = read_json(path)

    if isinstance(data, dict):

        data.setdefault("jobs", {})
        data.setdefault("chunks", {})

        return data

    return {
        "created_at": utc_now(),
        "updated_at": utc_now(),
        "worker_id": worker_id,
        "cube": CUBE["cube"],
        "jobs": {},
        "chunks": {},
    }


def save_worker_manifest(
    output: Path,
    worker_id: int,
    manifest: dict,
) -> None:

    manifest["updated_at"] = utc_now()

    path = (
        worker_dir(output, worker_id)
        / "manifest.json"
    )

    write_json(path, manifest)


def job_chunks(
    manifest: dict,
    job_id: int,
) -> list[dict]:

    return [
        record
        for record in manifest["chunks"].values()
        if record.get("job_id") == job_id
    ]


# ============================================================
# WORKER STATE REPORTER
# ============================================================

class WorkerReporter:

    def __init__(
        self,
        output: Path,
        worker_id: int,
        run_id: str,
        assigned_jobs: int,
        manifest: dict,
    ):

        self.output = output
        self.worker_id = worker_id

        self.path = (
            worker_dir(output, worker_id)
            / "state.json"
        )

        self.state = {
            "run_id": run_id,
            "worker_id": worker_id,
            "status": "RUNNING",
            "started_at": utc_now(),
            "updated_at": utc_now(),
            "assigned_jobs": assigned_jobs,
            "current_job": None,
            "current_dimension": None,
            "current_chunk": None,
            "request_started_at": None,
            "last_message": "Worker initialized",
            "recent": [],
            "completed_jobs": 0,
            "failed_jobs": 0,
            "partial_jobs": 0,
            "running_jobs": 0,
            "chunks_ok": 0,
            "chunks_failed": 0,
            "excel_rows": 0,
        }

        self.sync(manifest)

    def sync(self, manifest: dict) -> None:

        jobs = list(
            manifest["jobs"].values()
        )

        chunks = list(
            manifest["chunks"].values()
        )

        self.state["completed_jobs"] = sum(
            item.get("status") == "COMPLETED"
            for item in jobs
        )

        self.state["failed_jobs"] = sum(
            item.get("status") == "FAILED"
            for item in jobs
        )

        self.state["partial_jobs"] = sum(
            item.get("status") == "PARTIAL"
            for item in jobs
        )

        self.state["running_jobs"] = sum(
            item.get("status") == "RUNNING"
            for item in jobs
        )

        self.state["chunks_ok"] = sum(
            item.get("status") == "OK"
            for item in chunks
        )

        self.state["chunks_failed"] = sum(
            item.get("status") == "FAILED"
            for item in chunks
        )

        self.state["excel_rows"] = sum(
            int(item.get("excel_rows") or 0)
            for item in jobs
        )

        self.save()

    def save(self) -> None:

        self.state["updated_at"] = utc_now()

        write_json(
            self.path,
            self.state,
        )

    def event(
        self,
        message: str,
        level: str = "INFO",
    ) -> None:

        self.state["last_message"] = message

        recent = self.state["recent"]

        recent.append({
            "time": local_time(),
            "level": level,
            "message": message,
        })

        self.state["recent"] = recent[-8:]

        self.save()

        log(
            message,
            self.worker_id,
            level,
        )

    def set_job(
        self,
        job_id: int,
        dimension: str,
    ) -> None:

        self.state["current_job"] = job_id
        self.state["current_dimension"] = dimension
        self.state["current_chunk"] = None
        self.state["request_started_at"] = None

        self.save()

    def set_request(
        self,
        year_offset: int,
        member_offset: int,
    ) -> None:

        self.state["current_chunk"] = (
            f"YEAR {year_offset} | "
            f"MEMBER {member_offset}"
        )

        self.state["request_started_at"] = utc_now()

        self.save()

    def clear_request(self) -> None:

        self.state["request_started_at"] = None

        self.save()

    def finish(
        self,
        status: str,
        message: str,
    ) -> None:

        self.state["status"] = status
        self.state["current_job"] = None
        self.state["current_dimension"] = None
        self.state["current_chunk"] = None
        self.state["request_started_at"] = None
        self.state["finished_at"] = utc_now()

        self.event(
            message,
            "OK" if status == "COMPLETED" else "WARN",
        )


# ============================================================
# EXTRACTION JOB
# ============================================================

def run_job(
    args: argparse.Namespace,
    client: OICClient,
    output: Path,
    job: dict,
    manifest: dict,
    reporter: WorkerReporter,
) -> str:

    worker_id = job["worker_id"]
    job_id = job["id"]
    dimension = job["hierarchy"]

    directory = dimension_dir(
        output,
        worker_id,
        job_id,
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    reporter.set_job(
        job_id,
        dimension,
    )

    reporter.event(
        f"[{job_id}] EXPAND ALL {dimension}"
    )

    manifest["jobs"][str(job_id)] = {
        "hierarchy": dimension,
        "status": "RUNNING",
        "updated_at": utc_now(),
    }

    save_worker_manifest(
        output,
        worker_id,
        manifest,
    )

    reporter.sync(manifest)

    year_offset = 0

    request_count = 0

    failed = False
    partial = False
    no_more_years = False

    last_error = None

    while True:
        if (
            args.max_years
            and year_offset >= args.max_years
        ):

            partial = True
            break

        member_offset = 0

        year_count = None

        while True:

            if (
                request_count
                >= args.max_requests_per_job
            ):

                partial = True
                break

            key = chunk_key(
                job_id,
                year_offset,
                member_offset,
            )

            basename = chunk_name(
                year_offset,
                member_offset,
            )

            request_file = (
                directory
                / f"{basename}.request.xml"
            )

            response_file = (
                directory
                / f"{basename}.response.xml"
            )

            reporter.set_request(
                year_offset,
                member_offset,
            )

            existing = manifest[
                "chunks"
            ].get(key)

            result = None

            # =================================================
            # RESUME SUCCESSFUL RAW XML
            # =================================================

            if (
                existing
                and existing.get("status") == "OK"
                and response_file.exists()
                and existing.get("year_page_size")
                    == args.years_per_request
                and existing.get("member_page_size")
                    == args.members_per_request
            ):

                result = parse_xmla(
                    response_file,
                    dimension,
                )

                if result["valid"]:

                    reporter.event(
                        f"[{job_id}] RESUME "
                        f"YEAR={year_offset} "
                        f"MEMBER={member_offset}",
                        "OK",
                    )

                else:
                    result = None

            # =================================================
            # NEW XMLA REQUEST
            # =================================================

            if result is None:

                mdx = build_mdx(
                    dimension=dimension,
                    year_offset=year_offset,
                    member_offset=member_offset,
                    years_per_request=(
                        args.years_per_request
                    ),
                    members_per_request=(
                        args.members_per_request
                    ),
                )

                soap = build_soap(
                    mdx
                )

                request_file.write_bytes(
                    soap
                )

                record = {
                    "job_id": job_id,
                    "dimension": dimension,
                    "year_offset": year_offset,
                    "member_offset": member_offset,
                    "year_page_size": (
                        args.years_per_request
                    ),
                    "member_page_size": (
                        args.members_per_request
                    ),
                    "request_file": request_file.name,
                    "response_file": response_file.name,
                    "started_at": utc_now(),
                    "status": "RUNNING",
                }

                manifest["chunks"][key] = record

                save_worker_manifest(
                    output,
                    worker_id,
                    manifest,
                )

                reporter.event(
                    f"[{job_id}] REQUEST "
                    f"YEAR={year_offset} "
                    f"MEMBER={member_offset}"
                )

                try:

                    http = client.execute(
                        soap,
                        response_file,
                    )

                    record.update(http)

                    result = parse_xmla(
                        response_file,
                        dimension,
                    )

                    record["result"] = {
                        key: value
                        for key, value in result.items()
                        if key not in (
                            "years",
                            "row_descriptors",
                            "cells",
                        )
                    }

                    success = (
                        http["http_status"] == 200
                        and result["valid"]
                    )

                    if not success:

                        failed = True

                        record["status"] = "FAILED"

                        last_error = (
                            f"HTTP {http['http_status']} | "
                            f"{result.get('error') or result.get('preview') or 'Invalid XMLA response'}"
                        )

                        record["error"] = last_error

                        reporter.event(
                            f"[{job_id}] FAILED: {last_error}",
                            "ERROR",
                        )

                    else:

                        record["status"] = "OK"

                        reporter.event(
                            f"[{job_id}] OK | "
                            f"{result['year_count']} years | "
                            f"{result['member_count']} members | "
                            f"{result['nonempty_cells']} cells | "
                            f"{http['response_bytes']:,} bytes",
                            "OK",
                        )

                except (
                    requests.RequestException,
                    RuntimeError,
                    OSError,
                    ET.ParseError,
                ) as exc:

                    failed = True
                    last_error = str(exc)

                    record["status"] = "FAILED"
                    record["error"] = last_error

                    reporter.event(
                        f"[{job_id}] FAILED: {last_error}",
                        "ERROR",
                    )

                finally:

                    record["completed_at"] = utc_now()

                    manifest["chunks"][key] = record

                    save_worker_manifest(
                        output,
                        worker_id,
                        manifest,
                    )

                    reporter.clear_request()
                    reporter.sync(manifest)

                request_count += 1

            # =================================================
            # PAGINATION
            # =================================================

            if failed:
                break

            current_years = result[
                "year_count"
            ]

            current_members = result[
                "member_count"
            ]

            if year_count is None:
                year_count = current_years

            if current_years == 0:

                no_more_years = True
                break

            if current_members == 0:
                break

            if (
                current_members
                < args.members_per_request
            ):
                break

            member_offset += (
                args.members_per_request
            )

            if args.delay > 0:
                time.sleep(args.delay)

        if failed or partial or no_more_years:
            break

        if (
            year_count is None
            or year_count < args.years_per_request
        ):
            break

        year_offset += (
            args.years_per_request
        )

        if args.delay > 0:
            time.sleep(args.delay)

    # ========================================================
    # JOB STATUS
    # ========================================================

    records = job_chunks(
        manifest,
        job_id,
    )

    successful_records = [
        item
        for item in records
        if item.get("status") == "OK"
    ]

    if failed:

        status = (
            "PARTIAL"
            if successful_records
            else "FAILED"
        )

    elif partial:
        status = "PARTIAL"

    else:
        status = "COMPLETED"

    job_record = {
        "hierarchy": dimension,
        "status": status,
        "updated_at": utc_now(),
        "request_count": request_count,
        "error": last_error,
    }

    manifest["jobs"][str(job_id)] = job_record

    save_worker_manifest(
        output,
        worker_id,
        manifest,
    )

    reporter.event(
        f"[{job_id}] EXPORTING EXCEL"
    )

    # ========================================================
    # EXCEL EXPORT
    # ========================================================

    try:

        excel_path, rows = export_hierarchy(
            output,
            job,
            records,
            status,
        )

        if excel_path:

            job_record["excel_file"] = str(
                excel_path
            )

            job_record["excel_rows"] = rows

            reporter.event(
                f"[{job_id}] EXCEL "
                f"{rows:,} rows | "
                f"{excel_path.name}",
                "OK",
            )

    except Exception as exc:

        job_record["excel_error"] = str(exc)

        reporter.event(
            f"[{job_id}] EXCEL ERROR: {exc}",
            "ERROR",
        )

        if status == "COMPLETED":
            status = "PARTIAL"
            job_record["status"] = status

    save_worker_manifest(
        output,
        worker_id,
        manifest,
    )

    reporter.sync(manifest)

    reporter.event(
        f"[{job_id}] {status}",
        "OK" if status == "COMPLETED" else "WARN",
    )

    return status


# ============================================================
# WORKER EXECUTION
# ============================================================

def run_worker(
    args: argparse.Namespace,
    output: Path,
    plan: dict,
) -> int:

    worker_id = args.worker_id

    if not (
        1 <= worker_id <= plan["workers"]
    ):

        raise RuntimeError(
            f"Invalid worker ID: {worker_id}"
        )

    assigned = [
        job
        for job in plan["jobs"]
        if job["worker_id"] == worker_id
    ]

    directory = worker_dir(
        output,
        worker_id,
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    manifest = load_worker_manifest(
        output,
        worker_id,
    )

    run_id = args.run_id or (
        "manual-" + uuid.uuid4().hex[:8]
    )

    reporter = WorkerReporter(
        output=output,
        worker_id=worker_id,
        run_id=run_id,
        assigned_jobs=len(assigned),
        manifest=manifest,
    )

    console.print()

    console.print(
        Panel(
            Group(
                Align.center(
                    Text(
                        APP_NAME,
                        style="bold bright_cyan",
                    )
                ),
                Align.center(
                    Text(
                        f"WORKER {worker_id:02d} / "
                        f"{plan['workers']:02d}",
                        style="bold white",
                    )
                ),
                Align.center(
                    Text(
                        f"{len(assigned)} assigned hierarchies",
                        style="dim",
                    )
                ),
            ),
            border_style="cyan",
            padding=(1, 2),
        )
    )

    env = read_env()

    cookie = env.get(
        "OIC_COOKIE",
        "",
    ).strip()

    token = env.get(
        "OIC_BEARER_TOKEN",
        "",
    ).strip()

    client = None
    failures = 0
    processed = 0

    try:

        if not cookie:

            raise RuntimeError(
                "OIC_COOKIE missing from .env"
            )

        if Workbook is None:

            raise RuntimeError(
                "openpyxl missing. "
                "Install with: pip install openpyxl"
            )

        client = OICClient(
            cookie=cookie,
            token=token,
            timeout=args.timeout,
            max_response_mb=(
                args.max_response_mb
            ),
            worker_id=worker_id,
        )

        reporter.event(
            "YEAR = COLUMNS | "
            "DIMENSIONS + MEASURES = ROWS"
        )

        for job in assigned:

            job_id = str(
                job["id"]
            )

            existing = manifest[
                "jobs"
            ].get(job_id, {})

            if (
                existing.get("status") == "COMPLETED"
                and not args.force
            ):

                reporter.event(
                    f"[{job_id}] SKIP COMPLETED",
                    "OK",
                )

                continue

            if (
                args.max_jobs
                and processed >= args.max_jobs
            ):
                break

            processed += 1

            status = run_job(
                args=args,
                client=client,
                output=output,
                job=job,
                manifest=manifest,
                reporter=reporter,
            )

            if status != "COMPLETED":
                failures += 1

            if args.delay > 0:
                time.sleep(args.delay)

        reporter.sync(manifest)

        reporter.finish(
            (
                "COMPLETED_WITH_ERRORS"
                if failures
                else "COMPLETED"
            ),
            (
                "Worker finished | "
                f"{reporter.state['completed_jobs']} completed | "
                f"{reporter.state['failed_jobs']} failed | "
                f"{reporter.state['partial_jobs']} partial"
            ),
        )

        return 1 if failures else 0

    except KeyboardInterrupt:

        reporter.finish(
            "STOPPED",
            "Worker interrupted",
        )

        return 130

    except Exception as exc:

        reporter.event(
            f"WORKER ERROR: {exc}",
            "ERROR",
        )

        reporter.finish(
            "ERROR",
            f"Worker stopped: {exc}",
        )

        return 1

    finally:

        if client is not None:
            client.close()

        save_worker_manifest(
            output,
            worker_id,
            manifest,
        )


# ============================================================
# LIVE MONITOR HELPERS
# ============================================================

def read_worker_state(
    output: Path,
    worker_id: int,
    run_id: str | None = None,
) -> dict:

    path = (
        worker_dir(output, worker_id)
        / "state.json"
    )

    state = read_json(path, {})

    if not isinstance(state, dict):
        return {}

    if (
        run_id is not None
        and state.get("run_id") != run_id
    ):
        return {}

    return state


def progress_bar(
    completed: int,
    total: int,
    width: int = 20,
) -> Text:

    total = max(total, 1)

    percentage = min(
        completed / total,
        1.0,
    )

    filled = round(
        width * percentage
    )

    bar = Text()

    bar.append(
        "━" * filled,
        style="bright_cyan",
    )

    bar.append(
        "─" * (width - filled),
        style="grey35",
    )

    bar.append(
        f" {percentage * 100:5.1f}%",
        style="bold white",
    )

    return bar


def status_text(status: str) -> Text:

    colors = {
        "RUNNING": "bright_cyan",
        "COMPLETED": "green",
        "COMPLETED_WITH_ERRORS": "yellow",
        "FAILED": "red",
        "PARTIAL": "yellow",
        "ERROR": "red",
        "STOPPED": "red",
        "STARTING": "yellow",
        "WAITING": "dim",
    }

    labels = {
        "RUNNING": "RUNNING",
        "COMPLETED": "DONE",
        "COMPLETED_WITH_ERRORS": "DONE / ERRORS",
        "FAILED": "FAILED",
        "PARTIAL": "PARTIAL",
        "ERROR": "ERROR",
        "STOPPED": "STOPPED",
        "STARTING": "STARTING",
        "WAITING": "WAITING",
    }

    return Text(
        labels.get(status, status),
        style=(
            "bold "
            + colors.get(status, "white")
        ),
    )


def collect_summary(
    output: Path,
    plan: dict,
    run_id: str | None,
) -> tuple[dict, list[dict]]:

    workers = []

    summary = {
        "total_jobs": len(plan["jobs"]),
        "completed_jobs": 0,
        "failed_jobs": 0,
        "partial_jobs": 0,
        "running_jobs": 0,
        "chunks_ok": 0,
        "chunks_failed": 0,
        "excel_rows": 0,
    }

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        assigned = sum(
            job["worker_id"] == worker_id
            for job in plan["jobs"]
        )

        state = read_worker_state(
            output,
            worker_id,
            run_id,
        )

        if not state:

            state = {
                "worker_id": worker_id,
                "status": "STARTING",
                "assigned_jobs": assigned,
                "completed_jobs": 0,
                "failed_jobs": 0,
                "partial_jobs": 0,
                "running_jobs": 0,
                "chunks_ok": 0,
                "chunks_failed": 0,
                "excel_rows": 0,
                "recent": [],
            }

        workers.append(state)

        for key in (
            "completed_jobs",
            "failed_jobs",
            "partial_jobs",
            "running_jobs",
            "chunks_ok",
            "chunks_failed",
            "excel_rows",
        ):

            summary[key] += int(
                state.get(key) or 0
            )

    summary["finished_jobs"] = (
        summary["completed_jobs"]
        + summary["failed_jobs"]
        + summary["partial_jobs"]
    )

    summary["remaining_jobs"] = max(
        0,
        summary["total_jobs"]
        - summary["finished_jobs"],
    )

    return summary, workers


# ============================================================
# RICH MONITOR RENDER
# ============================================================

def render_monitor(
    output: Path,
    plan: dict,
    run_id: str | None,
    started_at: str,
):

    summary, workers = collect_summary(
        output,
        plan,
        run_id,
    )

    elapsed = format_duration(
        elapsed_seconds(started_at)
    )

    # ========================================================
    # HEADER
    # ========================================================

    header = Panel(
        Group(
            Align.center(
                Text(
                    "OIC PARALLEL COLLECTOR",
                    style="bold bright_cyan",
                )
            ),
            Align.center(
                Text(
                    "LIVE EXTRACTION MONITOR",
                    style="bold white",
                )
            ),
            Align.center(
                Text(
                    "YEAR / DIMENSIONS / XMLA / EXCEL",
                    style="dim",
                )
            ),
        ),
        border_style="bright_cyan",
        padding=(1, 2),
    )

    # ========================================================
    # OVERVIEW
    # ========================================================

    overview = Table(
        box=rich_box.SIMPLE,
        show_header=False,
        expand=True,
        padding=(0, 2),
    )

    overview.add_column(
        "Metric",
        ratio=1,
    )

    overview.add_column(
        "Value",
        justify="right",
        ratio=1,
    )

    overview.add_column(
        "Metric2",
        ratio=1,
    )

    overview.add_column(
        "Value2",
        justify="right",
        ratio=1,
    )

    overview.add_row(
        "TOTAL JOBS",
        f"[bold]{summary['total_jobs']}[/bold]",
        "COMPLETED",
        f"[green]{summary['completed_jobs']}[/green]",
    )

    overview.add_row(
        "REMAINING",
        f"[cyan]{summary['remaining_jobs']}[/cyan]",
        "PARTIAL",
        f"[yellow]{summary['partial_jobs']}[/yellow]",
    )

    overview.add_row(
        "CHUNKS OK",
        f"[green]{summary['chunks_ok']:,}[/green]",
        "CHUNKS FAILED",
        f"[red]{summary['chunks_failed']:,}[/red]",
    )

    overview.add_row(
        "EXCEL ROWS",
        f"[bold]{summary['excel_rows']:,}[/bold]",
        "ELAPSED",
        f"[bold]{elapsed}[/bold]",
    )

    metrics = Panel(
        overview,
        title="[bold]OVERVIEW[/bold]",
        border_style="grey50",
    )

    # ========================================================
    # OVERALL PROGRESS
    # ========================================================

    overall = Panel(
        Align.center(
            progress_bar(
                summary["finished_jobs"],
                summary["total_jobs"],
                width=38,
            )
        ),
        title="[bold]TOTAL PROGRESS[/bold]",
        border_style="grey50",
    )

    # ========================================================
    # WORKER TABLE
    # ========================================================

    table = Table(
        box=rich_box.ROUNDED,
        expand=True,
        header_style="bold cyan",
        border_style="grey35",
        show_lines=False,
    )

    table.add_column(
        "WORKER",
        justify="center",
        no_wrap=True,
    )

    table.add_column(
        "STATUS",
        no_wrap=True,
    )

    table.add_column(
        "JOBS",
        justify="center",
        no_wrap=True,
    )

    table.add_column(
        "PROGRESS",
    )

    table.add_column(
        "CHUNKS",
        justify="right",
        no_wrap=True,
    )

    table.add_column(
        "CURRENT JOB",
        overflow="ellipsis",
    )

    table.add_column(
        "WAIT",
        justify="right",
        no_wrap=True,
    )

    for state in workers:

        assigned = int(
            state.get("assigned_jobs") or 0
        )

        completed = int(
            state.get("completed_jobs") or 0        )

        failed = int(
            state.get("failed_jobs") or 0
        )

        partial = int(
            state.get("partial_jobs") or 0
        )

        finished = (
            completed + failed + partial
        )

        current_job = state.get(
            "current_job"
        )

        current_dimension = state.get(
            "current_dimension"
        )

        if current_job is not None:

            current = (
                f"[{current_job}] "
                f"{current_dimension or ''}"
            )

        else:
            current = "—"

        waiting = elapsed_seconds(
            state.get("request_started_at")
        )

        wait_text = (
            format_duration(waiting)
            if state.get("request_started_at")
            else "—"
        )

        chunks = (
            f"[green]{state.get('chunks_ok', 0)}[/green]"
            " / "
            f"[red]{state.get('chunks_failed', 0)}[/red]"
        )

        table.add_row(
            f"[bold]W{state['worker_id']:02d}[/bold]",
            status_text(
                state.get("status", "WAITING")
            ),
            f"{finished}/{assigned}",
            progress_bar(
                finished,
                assigned,
                width=12,
            ),
            chunks,
            current,
            wait_text,
        )

    workers_panel = Panel(
        table,
        title="[bold]PARALLEL WORKERS[/bold]",
        border_style="bright_cyan",
    )

    # ========================================================
    # CURRENT ACTIVITY
    # ========================================================

    activity = Table(
        box=rich_box.SIMPLE,
        expand=True,
        show_header=True,
        header_style="bold cyan",
        padding=(0, 1),
    )

    activity.add_column(
        "WORKER",
        width=8,
        no_wrap=True,
    )

    activity.add_column(
        "LAST ACTIVITY",
        overflow="ellipsis",
    )

    activity.add_column(
        "TIME",
        justify="right",
        width=10,
    )

    for state in workers:

        message = state.get(
            "last_message"
        ) or "Initializing..."

        recent = state.get(
            "recent"
        ) or []

        event_time = (
            recent[-1].get("time", "")
            if recent
            else ""
        )

        activity.add_row(
            f"W{state['worker_id']:02d}",
            str(message),
            event_time,
        )

    activity_panel = Panel(
        activity,
        title="[bold]LATEST ACTIVITY[/bold]",
        border_style="grey50",
    )

    # ========================================================
    # FOOTER
    # ========================================================

    footer = Panel(
        Text(
            "4 WORKERS  |  "
            "YEAR ON COLUMNS  |  "
            "RAW XML + EXCEL  |  "
            "CTRL+C: CLOSE MONITOR ONLY",
            style="dim",
            justify="center",
        ),
        border_style="grey35",
    )

    return Group(
        header,
        metrics,
        overall,
        workers_panel,
        activity_panel,
        footer,
    )


# ============================================================
# MONITOR LOOP
# ============================================================

def monitor_workers(
    output: Path,
    plan: dict,
    run_id: str,
) -> int:

    started_at = utc_now()

    terminal_states = {
        "COMPLETED",
        "COMPLETED_WITH_ERRORS",
        "ERROR",
        "STOPPED",
    }

    try:

        with Live(
            render_monitor(
                output,
                plan,
                run_id,
                started_at,
            ),
            console=console,
            refresh_per_second=1,
            screen=True,
            transient=False,
        ) as live:

            while True:

                dashboard = render_monitor(
                    output,
                    plan,
                    run_id,
                    started_at,
                )

                live.update(
                    dashboard,
                    refresh=True,
                )

                _, workers = collect_summary(
                    output,
                    plan,
                    run_id,
                )

                if all(
                    state.get("status") in terminal_states
                    for state in workers
                ):
                    break

                time.sleep(
                    MONITOR_REFRESH_SECONDS
                )

    except KeyboardInterrupt:

        console.print(
            "[yellow]Monitor stopped. "
            "Worker tabs continue running.[/yellow]"
        )

        return 130

    merge_manifests(
        output,
        plan,
        quiet=True,
    )

    return 0


# ============================================================
# POWERSHELL TERMINAL LAUNCHER
# ============================================================

def powershell_quote(value: str) -> str:
    value = str(value)

    if any(char in value for char in ("\x00", "\r", "\n")):
        raise ValueError("Invalid PowerShell argument")

    return "'" + value.replace("'", "''") + "'"

def find_powershell() -> str:
    if os.name != "nt":
        raise RuntimeError("PowerShell launcher requires Windows")

    candidates = [
        shutil.which("pwsh.exe"),
        shutil.which("powershell.exe"),
        shutil.which("pwsh"),
        shutil.which("powershell"),
    ]

    program_files = os.environ.get("ProgramFiles")
    system_root = os.environ.get("SystemRoot", r"C:\Windows")

    if program_files:
        candidates.append(
            str(Path(program_files) / "PowerShell" / "7" / "pwsh.exe")
        )

    candidates.append(
        str(
            Path(system_root)
            / "System32"
            / "WindowsPowerShell"
            / "v1.0"
            / "powershell.exe"
        )
    )

    for candidate in candidates:
        if candidate and Path(candidate).is_file():
            return str(Path(candidate).resolve())

    raise RuntimeError(
        "PowerShell 5.1 or PowerShell 7 was not found"
    )

# oic_parallel_collector.py



def launch_terminal_window(
    commands: list[list[str]],
    *,
    title: str,
    cwd: Path,
    tab_titles: list[str] | None = None,
    keep_open: bool = True,
    maximized: bool = False,
) -> str:
    if os.name != "nt":
        raise RuntimeError("PowerShell launcher requires Windows")
    if not commands or any(not command for command in commands):
        raise ValueError("No executable command supplied")
    shell = find_powershell()
    working_dir = Path(cwd).resolve()
    if not working_dir.is_dir():
        raise FileNotFoundError(str(working_dir))
    if len(commands) == 1:
        command = "& " + " ".join(powershell_quote(str(arg)) for arg in commands[0])
        program = [shell, "-NoLogo", "-NoProfile"]
        if keep_open:
            program.append("-NoExit")
        program.extend(["-Command", "$Host.UI.RawUI.WindowTitle = " +
                        powershell_quote(title) + "; " + command])
        startup = subprocess.STARTUPINFO()
        if maximized:
            startup.dwFlags |= subprocess.STARTF_USESHOWWINDOW
            startup.wShowWindow = 3  # SW_MAXIMIZE, not hidden fullscreen
        process = subprocess.Popen(program, cwd=str(working_dir), startupinfo=startup,
                                   creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0))
        return process
    else:
        # Legacy diagnostic launcher only. Workers remain background processes.
        for command in commands:
            subprocess.Popen(command, cwd=str(working_dir),
                             creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    return "POWERSHELL"


def worker_powershell_command(
    args: argparse.Namespace,
    worker_id: int,
    output: Path,
    run_id: str,
) -> str:

    python_path = str(
        Path(sys.executable).resolve()
    )

    script_path = str(
        Path(__file__).resolve()
    )

    arguments = [
        f"--worker-id {worker_id}",
        f"--workers {args.workers}",
        (
            "--output-dir "
            + powershell_quote(str(output))
        ),
        (
            "--source "
            + powershell_quote(
                str(args.source.resolve())
            )
        ),
        (
            f"--years-per-request "
            f"{args.years_per_request}"
        ),
        (
            f"--members-per-request "
            f"{args.members_per_request}"
        ),
        f"--timeout {args.timeout}",
        f"--delay {args.delay}",
        (
            f"--max-response-mb "
            f"{args.max_response_mb}"
        ),
        (
            f"--max-requests-per-job "
            f"{args.max_requests_per_job}"
        ),
        f"--max-years {args.max_years}",
        f"--max-jobs {args.max_jobs}",
        (
            "--run-id "
            + powershell_quote(run_id)
        ),
    ]

    if args.force:
        arguments.append("--force")

    command = (
        f"& {powershell_quote(python_path)} "
        f"{powershell_quote(script_path)} "
        + " ".join(arguments)
    )

    return command

# oic_parallel_collector.py
# Replace the existing legacy launch_worker_tabs() function.

def launch_worker_tabs(args: argparse.Namespace, output: Path, plan: dict) -> str:
    if os.name != "nt":
        raise RuntimeError("Worker launch requires Windows")
    workers = int(plan["workers"])
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
    commands = [worker_powershell_command(args, wid, output, run_id)
                for wid in range(1, workers + 1)]
    shell = find_powershell()
    for cmd in commands:
        subprocess.Popen([shell, "-NoLogo", "-NoProfile", "-Command", cmd],
                         cwd=str(HERE),
                         creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    log(f"Started {workers} background workers through PowerShell", level="OK")
    return monitor_workers(output, plan, run_id)


def merge_manifests(
    output: Path,
    plan: dict,
    quiet: bool = False,
) -> dict:

    summary = {
        "created_at": utc_now(),
        "cube": CUBE["cube"],
        "catalog": CUBE["catalog"],
        "workers": plan["workers"],
        "total_jobs": len(plan["jobs"]),
        "completed": 0,
        "failed": 0,
        "partial": 0,
        "running": 0,
        "pending": 0,
        "successful_chunks": 0,
        "failed_chunks": 0,
        "excel_rows": 0,
        "jobs": [],
    }

    manifests = {}

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        manifests[worker_id] = (
            load_worker_manifest(
                output,
                worker_id,
            )
        )

    for job in plan["jobs"]:

        worker_id = job["worker_id"]

        manifest = manifests[
            worker_id
        ]

        record = manifest[
            "jobs"
        ].get(
            str(job["id"]),
            {},
        )

        status = record.get(
            "status",
            "PENDING",
        )

        chunks = job_chunks(
            manifest,
            job["id"],
        )

        successful = sum(
            chunk.get("status") == "OK"
            for chunk in chunks
        )

        failed = sum(
            chunk.get("status") == "FAILED"
            for chunk in chunks
        )

        summary[
            "successful_chunks"
        ] += successful

        summary[
            "failed_chunks"
        ] += failed

        summary[
            "excel_rows"
        ] += int(
            record.get("excel_rows") or 0
        )

        status_key = status.lower()

        if status_key in summary:
            summary[status_key] += 1

        else:
            summary["pending"] += 1

        summary["jobs"].append({
            "id": job["id"],
            "worker_id": worker_id,
            "hierarchy": job["hierarchy"],
            "status": status,
            "successful_chunks": successful,
            "failed_chunks": failed,
            "excel_file": record.get(
                "excel_file"
            ),
            "excel_rows": record.get(
                "excel_rows"
            ),
            "error": record.get(
                "error"
            ),
        })

    destination = (
        output
        / "master_manifest.json"
    )

    write_json(
        destination,
        summary,
    )

    if not quiet:

        log(
            f"TOTAL={summary['total_jobs']} | "
            f"COMPLETED={summary['completed']} | "
            f"FAILED={summary['failed']} | "
            f"PARTIAL={summary['partial']} | "
            f"PENDING={summary['pending']}"
        )

        log(
            f"MASTER: {destination}"
        )

    return summary


# ============================================================
# EXPORT EXISTING XML
# ============================================================

def export_existing(
    output: Path,
    plan: dict,
) -> int:

    exported = 0
    failed = 0

    for job in plan["jobs"]:

        worker_id = job["worker_id"]

        manifest = load_worker_manifest(
            output,
            worker_id,
        )

        record = manifest[
            "jobs"
        ].get(
            str(job["id"]),
            {},
        )

        status = record.get(
            "status",
            "PARTIAL",
        )

        chunks = job_chunks(
            manifest,
            job["id"],
        )

        try:

            path, rows = export_hierarchy(
                output,
                job,
                chunks,
                status,
            )

            if path:

                exported += 1

                log(
                    f"[{job['id']}] "
                    f"{path.name} | "
                    f"{rows:,} rows",
                    level="OK",
                )

        except Exception as exc:

            failed += 1

            log(
                f"[{job['id']}] EXPORT ERROR: {exc}",
                level="ERROR",
            )

    log(
        f"EXPORTED={exported} | FAILED={failed}"
    )

    return 1 if failed else 0


# ============================================================
# STATIC STATUS VIEW
# ============================================================

def show_status(
    output: Path,
    plan: dict,
) -> int:

    # No run_id filter: show most recent worker states.
    dashboard = render_monitor(
        output=output,
        plan=plan,
        run_id=None,
        started_at=utc_now(),
    )

    console.print(dashboard)

    merge_manifests(
        output,
        plan,
        quiet=True,
    )

    return 0


# ============================================================
# CLI
# ============================================================

def build_parser() -> argparse.ArgumentParser:

    parser = argparse.ArgumentParser(
        description=(
            "OIC XMLA Parallel Collector "
            "with Rich Live Monitor and "
            "PowerShell background workers"
        ),
        formatter_class=(
            argparse.RawDescriptionHelpFormatter
        ),
    )

    # ========================================================
    # OPERATIONS
    # ========================================================

    parser.add_argument(
        "--prepare",
        action="store_true",
        help="Prepare hierarchy job plan",
    )

    parser.add_argument(
        "--launch",
        action="store_true",
        help=(
            "Launch workers as tabs "
            "and display Rich monitor"
        ),
    )

    parser.add_argument(
        "--worker-id",
        type=int,
        default=None,
        help="Run one specific worker",
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=DEFAULT_WORKERS,
        help="Total worker count; default 4",
    )

    parser.add_argument(
        "--status",
        action="store_true",
        help="Show current collection status",
    )

    parser.add_argument(
        "--merge",
        action="store_true",
        help="Merge worker manifests",
    )

    parser.add_argument(
        "--export-only",
        action="store_true",
        help="Export Excel from existing raw XML",
    )

    # ========================================================
    # PATHS
    # ========================================================

    parser.add_argument(
        "--source",
        type=Path,
        default=DEFAULT_SOURCE,
        help="Expand Probe manifest",
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT,
        help="Output directory",
    )

    # ========================================================
    # EXTRACTION
    # ========================================================

    parser.add_argument(
        "--years-per-request",
        type=int,
        default=DEFAULT_YEARS_PER_REQUEST,
    )

    parser.add_argument(
        "--members-per-request",
        type=int,
        default=DEFAULT_MEMBERS_PER_REQUEST,
    )

    parser.add_argument(
        "--timeout",
        type=int,
        default=DEFAULT_TIMEOUT,
    )

    parser.add_argument(
        "--delay",
        type=float,
        default=DEFAULT_DELAY,
    )

    parser.add_argument(
        "--max-response-mb",
        type=int,
        default=DEFAULT_MAX_RESPONSE_MB,
    )

    parser.add_argument(
        "--max-requests-per-job",
        type=int,
        default=DEFAULT_MAX_REQUESTS_PER_JOB,
    )

    parser.add_argument(
        "--max-years",
        type=int,
        default=0,
        help="Limit years for testing; 0 = all",
    )

    parser.add_argument(
        "--max-jobs",
        type=int,
        default=0,
        help="Limit jobs per worker; 0 = all",
    )

    parser.add_argument(
        "--force",
        action="store_true",
        help="Revisit completed jobs",
    )

    # Internal run identifier passed to worker tabs.
    parser.add_argument(
        "--run-id",
        default=None,
        help=argparse.SUPPRESS,
    )

    return parser



# ============================================================
# SELECTED-QUERY WORKFLOW — reuses OICClient, XMLA HTTP, logging,
# Rich, PowerShell and file helpers from the legacy engine.
# The old full-hierarchy CLI remains diagnostic, never the default.
# ============================================================

# ------------------------------------------------------------
# Sequential batch controller; each query retains its existing worker engine.
# ------------------------------------------------------------

def batch_manifest_path(batch_dir: Path) -> Path:
    return batch_dir / "batch_manifest.json"


def process_pid_alive(pid: int) -> bool:
    if not pid or int(pid) < 1:
        return False
    if os.name != "nt":
        try:
            os.kill(int(pid), 0)
            return True
        except ProcessLookupError:
            return False
        except PermissionError:
            return True
    import ctypes
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.OpenProcess.argtypes = (ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong)
    kernel.OpenProcess.restype = ctypes.c_void_p
    kernel.GetExitCodeProcess.argtypes = (ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong))
    kernel.GetExitCodeProcess.restype = ctypes.c_int
    kernel.CloseHandle.argtypes = (ctypes.c_void_p,)
    kernel.CloseHandle.restype = ctypes.c_int
    handle = kernel.OpenProcess(0x1000, False, int(pid))
    if not handle:
        return ctypes.get_last_error() == 5
    try:
        code = ctypes.c_ulong()
        if not kernel.GetExitCodeProcess(handle, ctypes.byref(code)):
            return False
        return code.value == 259
    finally:
        kernel.CloseHandle(handle)


def force_stop_batch(batch_dir: Path) -> int:
    """Terminate only tracked processes from this exact batch after user confirmation.

    Process creation time and full Python executable path are checked to avoid PID reuse.
    The OIC Main UI remains alive and no other Python processes are enumerated.
    """
    import ctypes
    batch = batch_load(Path(batch_dir).resolve())
    run_id = batch["run_id"]
    (Path(batch_dir) / "stop.flag").touch(exist_ok=True)
    current = batch.get("current_query_id")
    if current:
        (Path(batch_dir) / current / "stop.flag").touch(exist_ok=True)
    candidates = []
    for item in batch["queries"]:
        candidates.extend(item.get("worker_processes") or [])
    if batch.get("coordinator_pid"):
        candidates.append({"pid": batch["coordinator_pid"],
                          "created_at": batch.get("coordinator_started"),
                          "executable": str(Path(sys.executable).resolve())})
    kernel = ctypes.WinDLL("kernel32", use_last_error=True) if os.name == "nt" else None
    if kernel:
        kernel.OpenProcess.argtypes = (ctypes.c_ulong, ctypes.c_int, ctypes.c_ulong)
        kernel.OpenProcess.restype = ctypes.c_void_p
        kernel.QueryFullProcessImageNameW.argtypes = (
            ctypes.c_void_p, ctypes.c_ulong, ctypes.c_wchar_p, ctypes.POINTER(ctypes.c_ulong))
        kernel.QueryFullProcessImageNameW.restype = ctypes.c_int
        kernel.GetProcessTimes.argtypes = (
            ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p)
        kernel.GetProcessTimes.restype = ctypes.c_int
        kernel.TerminateProcess.argtypes = (ctypes.c_void_p, ctypes.c_ulong)
        kernel.TerminateProcess.restype = ctypes.c_int
        kernel.CloseHandle.argtypes = (ctypes.c_void_p,)
    count = 0
    for item in candidates:
        pid = int(item.get("pid") or 0)
        if pid <= 0 or pid == os.getpid() or not kernel:
            continue
        # PROCESS_TERMINATE | QUERY_LIMITED_INFORMATION
        handle = kernel.OpenProcess(0x0001 | 0x1000, False, pid)
        if not handle:
            continue
        try:
            pathbuf = ctypes.create_unicode_buffer(32768)
            length = ctypes.c_ulong(len(pathbuf))
            if not kernel.QueryFullProcessImageNameW(handle, 0, pathbuf, ctypes.byref(length)):
                continue
            expected = str(item.get("executable") or "").casefold()
            if not expected or Path(pathbuf.value).as_posix().casefold() != Path(expected).as_posix().casefold():
                continue
            created = item.get("created_at")
            if not created:
                continue
            creation, exit_time, kernel_time, user_time = [ctypes.c_ulonglong() for _ in range(4)]
            if not kernel.GetProcessTimes(handle, ctypes.byref(creation), ctypes.byref(exit_time),
                                          ctypes.byref(kernel_time), ctypes.byref(user_time)):
                continue
            process_start = creation.value / 10_000_000 - 11644473600
            if abs(process_start - float(created)) > 5:
                continue
            if kernel.TerminateProcess(handle, 130):
                count += 1
        finally:
            kernel.CloseHandle(handle)
    # The coordinator may have been killed before finalizing this manifest.
    # Keep any query not completed as interrupted and leave verified XML untouched.
    latest = batch_load(Path(batch_dir))
    if latest.get("run_id") != run_id:
        raise RuntimeError("Run changed during force stop")
    if latest.get("status") not in ("COMPLETED", "STOPPED"):
        for q in latest["queries"]:
            if q["status"] in ("RUNNING", "STARTING"):
                q["status"] = "PARTIAL"
                q["error"] = "Force-stopped: resume to recover uncompleted chunks"
                q["finished_at"] = utc_now()
        latest["current_query_id"] = None
        latest["status"] = "STOPPED"
        latest["finished_at"] = utc_now()
        batch_save(Path(batch_dir), latest)
    return count


def selected_plan(config: dict, run_dir: Path, *, required: bool = False) -> dict | None:
    """Read an already-applied immutable plan; never partition at START."""
    path = run_dir / "task_plan.json"
    if not path.exists():
        if required:
            raise ValueError(f"Precomputed task plan missing: {path}")
        return None  # explicit legacy selected-query CLI compatibility
    plan = read_json(path)
    if not isinstance(plan, dict) or plan.get("version") != 1:
        raise ValueError("Invalid task plan version")
    query_id = run_dir.name
    if plan.get("query_id") != query_id or plan.get("query_hash") != query_hash(config):
        raise ValueError("Immutable query identity/hash disagrees with task plan")
    n = int(config["workers"])
    xs, ys = int(config["x_chunk"]), int(config["y_chunk"])
    expected = [{"worker_id": i, "first_x_offset": (i - 1) * xs,
                 "x_stride": n * xs, "first_y_offset": 0, "y_stride": ys}
                for i in range(1, n + 1)]
    if (plan.get("partition") != "x_page_stride" or plan.get("status") != "PLANNED"
            or plan.get("workers") != n or plan.get("x_chunk") != xs
            or plan.get("y_chunk") != ys or plan.get("total_chunks", "bad") is not None
            or plan.get("assignments") != expected):
        raise ValueError("Applied plan assignment/configuration mismatch")
    return plan


def batch_load(batch_dir: Path) -> dict:
    data = read_json(batch_manifest_path(batch_dir))
    if (not isinstance(data, dict) or data.get("storage_version") != 1
            or not isinstance(data.get("queries"), list) or not data.get("run_id")):
        raise ValueError("Invalid batch manifest")
    seen = set()
    for item in data["queries"]:
        query_id = item.get("query_id")
        if not isinstance(query_id, str) or not re.fullmatch(r"Q\d{3,}", query_id):
            raise ValueError("Invalid batch query id")
        if query_id in seen:
            raise ValueError("Duplicate batch query id")
        seen.add(query_id)
        target = (batch_dir / item["config_path"]).resolve()
        if target != (batch_dir / query_id / "query_config.json").resolve():
            raise ValueError("Unsafe or unexpected batch query path")
        config, directory = selected_paths(target)
        if config.get("run_id") != data["run_id"] or directory.name != query_id:
            raise ValueError("Batch snapshot identity mismatch")
        selected_plan(config, directory, required=True)
    return data


def batch_save(batch_dir: Path, data: dict) -> None:
    write_json(batch_manifest_path(batch_dir), data)


@contextmanager
def file_write_lock(path: Path):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+b") as stream:
        stream.seek(0)
        if not stream.read(1):
            stream.write(b"\0")
            stream.flush()
        deadline = time.monotonic() + 8
        while True:
            try:
                stream.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                break
            except OSError:
                if time.monotonic() > deadline:
                    raise TimeoutError("Cannot acquire output file lock")
                time.sleep(0.05)
        try:
            yield
        finally:
            stream.seek(0)
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def history_upsert(batch_dir: Path, query: dict, run_id: str, config: dict) -> None:
    path = batch_dir.parents[1] / "history.csv"
    with file_write_lock(path.with_name(path.name + ".write.lock")):
        names = ["Run ID", "Query ID", "Category", "Cube", "Started At", "Finished At",
                 "Status", "Workers", "Records", "Query Hash", "Excel Path", "Error Summary"]
        rows = []
        if path.is_file():
            with path.open("r", encoding="utf-8-sig", newline="") as handle:
                rows = list(csv.DictReader(handle))
        updated = {"Run ID": run_id, "Query ID": query["query_id"],
                   "Category": config["code"], "Cube": config["cube"],
                   "Started At": query.get("started_at") or "",
                   "Finished At": query.get("finished_at") or "",
                   "Status": query.get("status") or "",
                   "Workers": config["workers"], "Records": query.get("records") or 0,
                   "Query Hash": query_hash(config), "Excel Path": query.get("excel_path") or "",
                   "Error Summary": query.get("error") or ""}
        rows = [row for row in rows
                if (row.get("Run ID"), row.get("Query ID")) != (run_id, query["query_id"])]
        rows.append(updated)
        path.parent.mkdir(parents=True, exist_ok=True)
        temporary = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
        try:
            with temporary.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=names, extrasaction="ignore")
                writer.writeheader()
                writer.writerows(rows)
                handle.flush()
                os.fsync(handle.fileno())
            os.replace(temporary, path)
        finally:
            temporary.unlink(missing_ok=True)

def batch_run(batch_dir: Path) -> int:
    if os.name != "nt":
        raise RuntimeError("Batch launcher requires Windows PowerShell")
    batch_dir = batch_dir.resolve()
    # Single-controller guarantee, including across separate Main UI instances.
    lock_path = batch_dir / "controller.lock"
    with lock_path.open("a+b") as lock_stream:
        lock_stream.seek(0)
        if not lock_stream.read(1):
            lock_stream.write(b"\0")
            lock_stream.flush()
        import msvcrt
        try:
            lock_stream.seek(0)
            msvcrt.locking(lock_stream.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError as exc:
            raise RuntimeError("A batch controller is already running") from exc
        try:
            try:
                return _batch_run_locked(batch_dir)
            except Exception as exc:
                data = read_json(batch_manifest_path(batch_dir), {}) or {}
                if data.get("status") not in ("COMPLETED", "STOPPED"):
                    data.update(status="FAILED", current_query_id=None,
                                finished_at=utc_now(), error=f"{type(exc).__name__}: {exc}"[:350])
                    batch_save(batch_dir, data)
                raise
        finally:
            lock_stream.seek(0)
            msvcrt.locking(lock_stream.fileno(), msvcrt.LK_UNLCK, 1)


def batch_monitor(batch_dir: Path) -> int:
    """Read-only observer: closing its console does not stop any worker."""
    batch_dir = batch_dir.resolve()
    started = time.monotonic()
    finished = {"COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED"}
    with Live(console=console, refresh_per_second=1, screen=True, transient=False) as live:
        while True:
            try:
                data = batch_load(batch_dir)
                pending = sum(x.get("status") == "PENDING" for x in data["queries"])
                completed = sum(x.get("status") == "COMPLETED" for x in data["queries"])
                failed = sum(x.get("status") in ("PARTIAL", "FAILED", "SESSION_EXPIRED")
                             for x in data["queries"])
                table = Table(title="APPLIED QUERIES", box=rich_box.ROUNDED, expand=True)
                for col in ("QUERY", "CATEGORY", "STATUS", "RECORDS", "EXCEL"):
                    table.add_column(col)
                for q in data["queries"]:
                    table.add_row(q["query_id"], q["code"], q["status"],
                                  str(q.get("records") or 0),
                                  Path(q["excel_path"]).name if q.get("excel_path") else "—")
                current = data.get("current_query_id")
                workers_table = Table(title="ACTIVE WORKERS", box=rich_box.ROUNDED, expand=True)
                for col in ("W", "PID", "STATUS", "CHUNK", "HTTP", "LAST EVENT"):
                    workers_table.add_column(col)
                total_ok = 0
                total_fail = 0
                if current:
                    current_entry = next((q for q in data["queries"] if q["query_id"] == current), None)
                    if current_entry:
                        config, folder = selected_paths(batch_dir / current_entry["config_path"])
                        workers, totals = selected_summary(config, folder)
                        total_ok, total_fail = totals["chunks_ok"], totals["chunks_failed"]
                        for w in workers:
                            workers_table.add_row(str(w.get("worker_id", "?")),
                                str(w.get("pid") or "—"), w.get("status", "STARTING"),
                                str(w.get("current_request") or "—"),
                                str(w.get("last_http") or "—"),
                                str(w.get("last_message") or "—")[:70])
                overview = (f"BATCH {data['run_id']} | {data['status']} | "
                            f"QUERIES {completed}/{len(data['queries'])} COMPLETED, "
                            f"{pending} PENDING, {failed} ERRORS\n"
                            f"ACTIVE {current or '—'} | CHUNKS OK {total_ok}, FAILED {total_fail} "
                            f"| ELAPSED {int(time.monotonic()-started)}s\n"
                            "Unknown total chunks: dynamic discovery, no estimated percentage. "
                            "Closing Monitor does not stop collection.")
                live.update(Group(Panel(overview, title="OIC BATCH LIVE MONITOR"),
                                  table, workers_table), refresh=True)
                if data["status"] in finished:
                    break
                if time.monotonic() - started > 20 and not batch_controller_active(batch_dir):
                    # Coordinator may have crashed; monitor does not mutate run state.
                    console.print("[red]Batch coordinator is not running; inspect process logs.[/red]")
                    return 1
                time.sleep(1)
            except KeyboardInterrupt:
                return 130
    return 0 if data["status"] == "COMPLETED" else 1


def batch_controller_active(batch_dir: Path) -> bool:
    if os.name != "nt":
        return False
    path = batch_dir / "controller.lock"
    if not path.exists():
        return False
    import msvcrt
    with path.open("a+b") as stream:
        stream.seek(0)
        try:
            msvcrt.locking(stream.fileno(), msvcrt.LK_NBLCK, 1)
        except OSError:
            return True
        stream.seek(0)
        msvcrt.locking(stream.fileno(), msvcrt.LK_UNLCK, 1)
    return False


def _batch_run_locked(batch_dir: Path) -> int:
    batch = batch_load(batch_dir)
    stop_path = batch_dir / "stop.flag"
    batch["status"] = "RUNNING"
    batch["coordinator_pid"] = os.getpid()
    batch["coordinator_started"] = time.time()
    batch["started_at"] = batch.get("started_at") or utc_now()
    batch["finished_at"] = None
    batch_save(batch_dir, batch)
    terminal = {"COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED"}
    halt_reason = None
    completed = 0
    if True:  # batch coordinator is headless; monitor is a separate process
        for order, entry in enumerate(batch["queries"], 1):
            if entry["status"] == "COMPLETED":
                completed += 1
                continue
            if stop_path.exists():
                halt_reason = "STOPPED"
                break
            config_path = batch_dir / entry["config_path"]
            config, run_dir = selected_paths(config_path)
            # Verified snapshot; never use current applied list for resumption.
            entry["status"] = "RUNNING"
            entry["started_at"] = entry.get("started_at") or utc_now()
            entry["finished_at"] = None
            entry["error"] = None
            batch["current_query_id"] = entry["query_id"]
            batch_save(batch_dir, batch)
            (run_dir / "stop.flag").unlink(missing_ok=True)
            # Workers may have finished before an interrupted export.
            # Re-export their verified cached XML, without repeating any requests.
            old_workers, old_totals = selected_summary(config, run_dir)
            if (old_totals["workers_completed"] == int(config["workers"])
                    and old_totals["chunks_failed"] == 0
                    and old_totals["chunks_ok"] > 0):
                try:
                    destination, count = selected_export(
                        config, run_dir, status="COMPLETED",
                        started_at=entry["started_at"], finished_at=utc_now())
                    entry.update(status="COMPLETED", excel_path=destination, records=count,
                                 error=None, finished_at=utc_now())
                    completed += 1
                    batch_save(batch_dir, batch)
                    history_upsert(batch_dir, entry, batch["run_id"], config)
                    continue
                except Exception as exc:
                    log(f"Cached export incomplete for {entry['query_id']}: {exc}; "
                        "resuming workers to recover missing chunks", level="WARN")
            processes = []
            log_streams = []
            try:
                for wid in range(1, int(config["workers"]) + 1):
                    directory = run_dir / f"worker_{wid:02d}"
                    directory.mkdir(parents=True, exist_ok=True)
                    write_json(directory / "state.json", {
                        "run_id": config["run_id"], "worker_id": wid, "status": "STARTING",
                        "current_request": None, "request_started_at": None,
                        "last_http": None, "last_message": "Starting this attempt",
                        "updated_at": utc_now(),
                    })
                    log_path = directory / "worker.log"
                    stream = log_path.open("ab")
                    log_streams.append(stream)
                    args = [str(Path(sys.executable).resolve()), str(Path(__file__).resolve()),
                            "--query-config", str(config_path.resolve()), "--worker-id", str(wid)]
                    proc = subprocess.Popen(args, cwd=str(HERE), stdout=stream, stderr=subprocess.STDOUT,
                                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                    processes.append(proc)
                    entry.setdefault("worker_processes", []).append(
                        {"worker_id": wid, "pid": proc.pid,
                         "created_at": time.time(), "executable": str(Path(sys.executable).resolve())})
                    batch_save(batch_dir, batch)
                while True:
                    if stop_path.exists():
                        (run_dir / "stop.flag").touch(exist_ok=True)
                    workers, totals = selected_summary(config, run_dir)
                    # No terminal output here: observer reads manifests and heartbeat files.
                    if all(proc.poll() is not None for proc in processes):
                        break
                    time.sleep(1)
                workers, totals = selected_summary(config, run_dir)
                states = [worker.get("status") for worker in workers]
                if "SESSION_EXPIRED" in states:
                    entry["status"] = "SESSION_EXPIRED"
                    halt_reason = "SESSION_EXPIRED"
                elif stop_path.exists() or "STOPPED" in states:
                    entry["status"] = "STOPPED"
                    halt_reason = "STOPPED"
                elif (all(st == "COMPLETED" for st in states)
                      and totals["chunks_failed"] == 0):
                    entry["status"] = "COMPLETED"
                elif totals["chunks_ok"]:
                    entry["status"] = "PARTIAL"
                else:
                    entry["status"] = "FAILED"
                if totals["chunks_ok"]:
                    # Export from existing OK XML responses; no network re-execution.
                    try:
                        destination, count = selected_export(
                            config, run_dir,
                            status=("PARTIAL" if entry["status"] in
                                    ("STOPPED", "SESSION_EXPIRED") else entry["status"]),
                            started_at=entry["started_at"], finished_at=utc_now())
                        entry["excel_path"] = destination
                        entry["records"] = count
                        if count == 0 and entry["status"] == "COMPLETED":
                            entry["status"] = "FAILED"
                            entry["error"] = "No nonempty rows to export"
                    except Exception as exc:
                        entry["status"] = "FAILED"
                        entry["error"] = f"Export failed: {type(exc).__name__}: {exc}"[:350]
                if entry["status"] == "COMPLETED" and not entry.get("excel_path"):
                    entry["status"] = "FAILED"
                    entry["error"] = "No Excel workbook produced"
                if entry["status"] == "COMPLETED":
                    completed += 1
                if entry["status"] in ("FAILED", "PARTIAL"):
                    entry["error"] = entry.get("error") or (
                        "Worker failed or collection incomplete; consult worker logs")
                if entry["status"] == "SESSION_EXPIRED":
                    entry["error"] = "Authentication expired or rejected"
                entry["finished_at"] = utc_now()
            except Exception as exc:
                entry["status"] = "FAILED"
                entry["error"] = f"{type(exc).__name__}: {exc}"[:350]
                entry["finished_at"] = utc_now()
                log(f"{entry['query_id']}: {entry['error']}", level="ERROR")
                # Failed query is isolated; continue after worker exit only.
                (run_dir / "stop.flag").touch(exist_ok=True)
                for proc in processes:
                    if proc.poll() is None:
                        try:
                            proc.wait(timeout=5)
                        except subprocess.TimeoutExpired:
                            proc.terminate()
                            try:
                                proc.wait(timeout=3)
                            except subprocess.TimeoutExpired:
                                proc.kill()
                                proc.wait(timeout=3)
            finally:
                for stream in log_streams:
                    stream.close()
                batch_save(batch_dir, batch)
                history_upsert(batch_dir, entry, batch["run_id"], config)
            if halt_reason:
                break
    batch["current_query_id"] = None
    batch["finished_at"] = utc_now()
    if halt_reason:
        batch["status"] = halt_reason
    elif all(e["status"] == "COMPLETED" for e in batch["queries"]):
        batch["status"] = "COMPLETED"
    elif any(e["status"] in ("COMPLETED", "PARTIAL") for e in batch["queries"]):
        batch["status"] = "PARTIAL"
    else:
        batch["status"] = "FAILED"
    batch_save(batch_dir, batch)
    console.print(Panel(f"BATCH {batch['run_id']} / {batch['status']}", title="FINAL SUMMARY"))
    return 0 if batch["status"] == "COMPLETED" else 1


def selected_paths(config_path: Path):
    config_path = config_path.resolve()
    run_dir = config_path.parent
    config = read_json(config_path)
    if not isinstance(config, dict):
        raise RuntimeError("Query config not found or invalid")
    validate_query(config)
    return config, run_dir


def selected_worker(args) -> int:
    config, run_dir = selected_paths(args.query_config)
    code, worker_id = config["code"], args.worker_id
    worker_count = int(config["workers"])
    if not 1 <= worker_id <= worker_count:
        raise RuntimeError("Invalid worker ID")
    plan = selected_plan(config, run_dir,
                         required=(run_dir.name.startswith("Q") and run_dir.parent.parent.name == "Runs"))
    assignment = plan["assignments"][worker_id - 1] if plan else {
        "first_x_offset": (worker_id - 1) * int(config["x_chunk"]),
        "x_stride": worker_count * int(config["x_chunk"]),
        "first_y_offset": 0, "y_stride": int(config["y_chunk"])}
    directory = run_dir / f"worker_{worker_id:02d}"
    chunks_dir = directory / "chunks"
    chunks_dir.mkdir(parents=True, exist_ok=True)
    lockpath = directory / "worker.lock"
    lock_created = False
    try:
        fd = os.open(lockpath, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        try:
            os.write(fd, json.dumps({"pid": os.getpid(), "run_id": config["run_id"]}).encode())
        finally:
            os.close(fd)
        lock_created = True
    except FileExistsError:
        old_lock = read_json(lockpath, {}) or {}
        previous_pid = old_lock.get("pid")
        if previous_pid and process_pid_alive(int(previous_pid)):
            raise RuntimeError(f"Worker {worker_id}: process already running (PID {previous_pid})")
        lockpath.unlink(missing_ok=True)
        fd = os.open(lockpath, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
        try:
            os.write(fd, json.dumps({"pid": os.getpid(), "run_id": config["run_id"]}).encode())
        finally:
            os.close(fd)
        lock_created = True
    manifest_path = directory / "manifest.json"
    state_path = directory / "state.json"
    heartbeat_path = directory / "heartbeat.json"
    existing = read_json(manifest_path, {})
    manifest = existing if existing.get("run_id") == config["run_id"] and existing.get("query_hash") == query_hash(config) else {}
    manifest.setdefault("chunks", {})
    manifest["run_id"] = config["run_id"]
    manifest["query_hash"] = query_hash(config)
    manifest["worker_id"] = worker_id
    manifest["updated_at"] = utc_now()
    stop_file = run_dir / "stop.flag"
    heartbeat_stop = threading.Event()
    state = dict(run_id=config["run_id"], worker_id=worker_id, pid=os.getpid(),
                 status="STARTING", current_request=None, request_started_at=None,
                 last_http=None, last_message="Initializing", updated_at=utc_now())
    def save_state():
        state["updated_at"] = utc_now()
        write_json(state_path, state)
    def heartbeat():
        while not heartbeat_stop.is_set():
            write_json(heartbeat_path, dict(run_id=config["run_id"], worker_id=worker_id,
                                            pid=os.getpid(), updated_at=utc_now()))
            heartbeat_stop.wait(2)
    save_state()
    thread = threading.Thread(target=heartbeat, daemon=True)
    thread.start()
    client = None
    x_size, y_size = int(config["x_chunk"]), int(config["y_chunk"])
    errors = 0
    completed = False
    try:
        env = read_env()
        cookie = env.get("OIC_COOKIE", "").strip()
        if not cookie:
            raise RuntimeError("SESSION_EXPIRED: OIC_COOKIE not configured")
        client = OICClient(cookie=cookie, token=env.get("OIC_BEARER_TOKEN", ""),
                           timeout=int(config.get("timeout", 60)),
                           max_response_mb=100, worker_id=worker_id, code=code)
        state["status"] = "RUNNING"
        save_state()
        # Deterministic X paging: a given X page belongs to exactly one worker.
        x_offset = assignment["first_x_offset"]
        while not stop_file.exists():
            y_offset = assignment["first_y_offset"]
            if (run_dir / "pause.flag").exists():
                state["status"] = "PAUSED"; save_state()
                while (run_dir / "pause.flag").exists() and not stop_file.exists():
                    time.sleep(1)
                state["status"] = "RUNNING"; save_state()
            if stop_file.exists():
                break
            x_returned = None
            while not stop_file.exists():
                key = f"{query_hash(config)}_x{x_offset:06d}_y{y_offset:06d}"
                req_path = chunks_dir / f"{key}.request.xml"
                resp_path = chunks_dir / f"{key}.response.xml"
                record = manifest["chunks"].get(key, {})
                data = None
                if record.get("status") == "OK" and resp_path.is_file():
                    try:
                        data = query_parse(resp_path.read_bytes())
                    except (ValueError, ET.ParseError, OSError):
                        data = None
                if data is None:
                    req_path.write_bytes(query_soap(config, x_offset, y_offset))
                    state.update(status="RUNNING", current_request=key,
                                 request_started_at=utc_now(), last_message=f"Request {key}")
                    save_state()
                    log(f"REQUEST X={x_offset} Y={y_offset}", worker_id)
                    http = None
                    retry_history = []
                    transient = {408, 425, 429, 500, 502, 503, 504}
                    max_attempts = max(1, int(config.get("retries", 3)) + 1)
                    for attempt in range(1, max_attempts + 1):
                        if stop_file.exists():
                            break
                        try:
                            http = client.execute(req_path.read_bytes(), resp_path)
                            state["last_http"] = http["http_status"]
                            code_http = int(http["http_status"])
                            if code_http == 200:
                                data = query_parse(resp_path.read_bytes())
                                break
                            if code_http in (401, 403):
                                raise RuntimeError(f"SESSION_EXPIRED: HTTP {code_http} after JWT refresh")
                            if code_http not in transient:
                                raise ValueError(f"XMLA HTTP {code_http}")
                            raise requests.HTTPError(f"XMLA HTTP {code_http}")
                        except (requests.RequestException, ValueError, ET.ParseError, RuntimeError) as exc:
                            code_http = http.get("http_status") if http else None
                            retryable = (code_http in transient or isinstance(
                                exc, (requests.Timeout, requests.ConnectionError, requests.HTTPError,
                                      requests.exceptions.ChunkedEncodingError)))
                            if "SESSION_EXPIRED" in str(exc).upper():
                                retryable = False
                            delay = min(30.0, float(2 ** min(attempt, 5))) + random.uniform(0, 0.4)
                            if code_http in (429, 503) and http:
                                raw_retry = str(http.get("retry_after") or "")
                                if raw_retry.isdecimal():
                                    delay = min(60.0, float(raw_retry))
                            retry_history.append({"attempt": attempt, "timestamp": utc_now(),
                                                  "http_status": code_http,
                                                  "error_type": type(exc).__name__,
                                                  "message": str(exc)[:160], "retryable": retryable,
                                                  "next_retry_delay": delay if retryable and attempt < max_attempts else None})
                            record["retry_history"] = retry_history
                            write_json(manifest_path, manifest)
                            if not retryable or attempt == max_attempts or stop_file.exists():
                                if stop_file.exists():
                                    break
                                raise
                            state.update(status="RETRYING", last_message=f"Retry {attempt}/{max_attempts}: {exc}")
                            save_state()
                            log(f"RETRYING {attempt}/{max_attempts}: {exc}", worker_id, "WARN")
                            deadline = time.monotonic() + delay
                            while not stop_file.exists() and time.monotonic() < deadline:
                                time.sleep(min(.25, max(0, deadline - time.monotonic())))
                            state["status"] = "RUNNING"
                            save_state()
                    if stop_file.exists():
                        break
                    if data is None:
                        raise RuntimeError("Empty XMLA dataset after attempts")
                    record = dict(status="OK", x_offset=x_offset, y_offset=y_offset,
                                  request_file=req_path.name, response_file=resp_path.name,
                                  x_count=len(data["axes"]["Axis0"]),
                                  y_tuples=len(data["axes"]["Axis1"]),
                                  cells=len(data["cells"]), http_status=200,
                                  response_bytes=resp_path.stat().st_size,
                                  retry_history=retry_history, completed_at=utc_now())
                    manifest["chunks"][key] = record
                    manifest["updated_at"] = utc_now()
                    write_json(manifest_path, manifest)
                    log(f"OK X={record['x_count']} Y={record['y_tuples']} cells={record['cells']}", worker_id, "OK")
                # Schema v2 puts Measures on Axis1 with X/ROWS; v1 preserved.
                row_keys = {tuple(m["unique_name"] for m in row
                                  if m["hierarchy"] != "[Measures]")
                            for row in data["axes"]["Axis1"]}
                if int(config.get("schema_version", 1)) == 2:
                    x_count = len(row_keys)
                    y_count = len(data["axes"]["Axis0"])
                else:
                    x_count = len(data["axes"]["Axis0"])
                    y_count = len(row_keys)
                x_returned = x_count
                state.update(current_request=None, request_started_at=None,
                             last_message=f"OK X={x_offset} Y={y_offset} cells={len(data['cells'])}")
                save_state()
                if x_count == 0 or y_count < y_size:
                    break
                y_offset += assignment["y_stride"]
                if stop_file.exists():
                    break
                if stop_file.is_file():
                    break
                time.sleep(float(config.get("delay", 1)))
            if x_returned is None or x_returned < x_size:
                completed = True
                break
            x_offset += assignment["x_stride"]
            if stop_file.exists():
                break
            time.sleep(float(config.get("delay", 1)))
        state["status"] = "STOPPED" if stop_file.exists() else "COMPLETED" if completed else "PARTIAL"
    except KeyboardInterrupt:
        state["status"] = "STOPPED"
    except Exception as exc:
        errors += 1
        text_message = str(exc)
        if "key" in locals():
            previous = manifest["chunks"].get(key, {})
            manifest["chunks"][key] = dict(previous, status="FAILED",
                                          x_offset=x_offset, y_offset=y_offset,
                                          response_file=resp_path.name,
                                          request_file=req_path.name,
                                          http_status=state.get("last_http"),
                                          error=text_message[:300],
                                          completed_at=utc_now())
        has_success = any(c.get("status") == "OK" for c in manifest["chunks"].values())
        state["status"] = ("SESSION_EXPIRED" if any(marker in text_message.upper() for marker in ("SESSION_EXPIRED", "COOKIE SESSION EXPIRED", "AUTH HTTP 401", "AUTH HTTP 403", "AUTHENTICATION REJECTED", "GENTOKEN"))
                           else "PARTIAL" if has_success else "FAILED")
        state["last_message"] = text_message[:200]
        if state["status"] == "SESSION_EXPIRED":
            stop_file.touch(exist_ok=True)
        log(f"FAILED: {text_message}", worker_id, "ERROR")
    finally:
        if client:
            client.close()
        heartbeat_stop.set()
        thread.join(timeout=3)
        state["current_request"] = None
        state["request_started_at"] = None
        save_state()
        manifest["status"] = state["status"]
        manifest["updated_at"] = utc_now()
        write_json(manifest_path, manifest)
        if lock_created:
            try:
                lockpath.unlink(missing_ok=True)
            except OSError:
                pass
    return 1 if errors else 0


def selected_summary(config: dict, run_dir: Path):
    workers = []
    totals = dict(chunks_ok=0, chunks_failed=0, cells=0, bytes=0)
    terminal = {"COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED", "CRASHED"}
    for wid in range(1, int(config["workers"]) + 1):
        directory = run_dir / f"worker_{wid:02d}"
        state = read_json(directory / "state.json", {}) or {}
        heartbeat = read_json(directory / "heartbeat.json", {}) or {}
        if state.get("run_id") != config["run_id"]:
            state = dict(status="STARTING", worker_id=wid, run_id=config["run_id"])
        elif state.get("status") not in terminal:
            pid = state.get("pid")
            last = parse_time(heartbeat.get("updated_at"))
            age = ((datetime.now(timezone.utc) - last).total_seconds() if last else None)
            # Startup grace prevents old/stale states masquerading as healthy workers.
            if pid and not process_pid_alive(int(pid)):
                state["status"] = "CRASHED"
            elif age is not None and age > 20:
                state["status"] = "STALE"
            elif not pid and age is not None and age > 20:
                state["status"] = "CRASHED"
        manifest = read_json(directory / "manifest.json", {}) or {}
        if manifest.get("run_id") == config["run_id"]:
            for chunk in manifest.get("chunks", {}).values():
                if chunk.get("status") == "OK":
                    totals["chunks_ok"] += 1
                    totals["cells"] += int(chunk.get("cells") or 0)
                    totals["bytes"] += int(chunk.get("response_bytes") or 0)
                elif chunk.get("status") == "FAILED":
                    totals["chunks_failed"] += 1
        workers.append(state)
    totals["workers_completed"] = sum(w.get("status") == "COMPLETED" for w in workers)
    totals["workers_failed"] = sum(w.get("status") in {"PARTIAL", "FAILED", "SESSION_EXPIRED"} for w in workers)
    totals["workers_terminal"] = sum(w.get("status") in terminal for w in workers)
    return workers, totals


def selected_export(config: dict, run_dir: Path, *, status: str | None = None,
                    started_at: str | None = None, finished_at: str | None = None) -> tuple[str, int]:
    if Workbook is None:
        raise RuntimeError("openpyxl required for Excel export")
    batch = run_dir.name.startswith("Q") and run_dir.parent.parent.name == "Runs"
    root = run_dir.parents[2] if batch else run_dir.parents[1]
    excel_dir = root / "Excel"
    excel_dir.mkdir(parents=True, exist_ok=True)
    run_id = config["run_id"]
    code = config["code"]
    query_id = run_dir.name if batch else config.get("query_id", "Q001")
    path = excel_dir / f"{run_id}_{code}_{query_id}.xlsx"
    tmp = run_dir / f".{run_id}_{code}_{query_id}_{uuid.uuid4().hex}.tmp.xlsx"
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("DATA")
    x_count, y_count = len(config["x"]), len(config["y"])
    headers = ["INSURANCE_CODE", "CUBE", "CATALOG"]
    for axis, count in (("X", x_count), ("Y", y_count)):
        for i in range(count):
            headers.extend([f"{axis}{i+1}_HIERARCHY", f"{axis}{i+1}_FIELD_CAPTION",
                            f"{axis}{i+1}_MEMBER_CAPTION", f"{axis}{i+1}_UNIQUE_NAME",
                            f"{axis}{i+1}_LEVEL"])
    headers += ["MEASURE", "MEASURE_CAPTION", "VALUE", "LAST_DATA_UPDATE", "RETRIEVED_AT"]
    display_map = {f.get("mdx"): f.get("caption", f.get("mdx"))
                   for f in config.get("selected_fields", []) if isinstance(f, dict)}
    sheet.append(headers)
    row_count, sheet_rows, sheet_number = 0, 1, 1
    for wid in range(1, int(config["workers"]) + 1):
        directory = run_dir / f"worker_{wid:02d}"
        manifest = read_json(directory / "manifest.json", {}) or {}
        if (manifest.get("run_id") != run_id or
                manifest.get("query_hash") != query_hash(config)):
            continue
        for name, chunk in sorted(manifest.get("chunks", {}).items()):
            if chunk.get("status") != "OK":
                continue
            response = directory / "chunks" / chunk["response_file"]
            if not response.is_file():
                raise FileNotFoundError(f"Missing cached XMLA response: {response}")
            data = query_parse(response.read_bytes())
            for record in query_excel_records(config, data):
                if sheet_rows >= EXCEL_MAX_ROWS:
                    sheet_number += 1
                    sheet = workbook.create_sheet(f"DATA_{sheet_number}")
                    sheet.append(headers)
                    sheet_rows = 1
                entry = [code, config["cube"], config["catalog"]]
                for values, names in ((record["X"], config["x"]), (record["Y"], config["y"])):
                    for field in names:
                        matching = next((m for m in values if m["hierarchy"] == field), {})
                        entry.extend([field, display_map.get(field, field), matching.get("caption"),
                                      matching.get("unique_name"), matching.get("level")])
                raw = record["VALUE"]
                # Excel numeric precision cannot exceed 15 significant digits.
                try:
                    from decimal import Decimal, InvalidOperation
                    number = Decimal(str(raw))
                    precision = len(number.normalize().as_tuple().digits)
                    value = float(number) if number.is_finite() and precision <= 15 else str(raw)
                except (InvalidOperation, TypeError, ValueError, OverflowError):
                    value = raw
                entry.extend([record["MEASURE"],
                              display_map.get(record["MEASURE"], record["MEASURE"]),
                              value, record["LAST_DATA_UPDATE"], chunk.get("completed_at")])
                sheet.append([excel_text(v) if isinstance(v, str) else v for v in entry])
                row_count += 1
                sheet_rows += 1
    if not row_count:
        raise RuntimeError("No nonempty XMLA result rows available for export")
    workers, totals = selected_summary(config, run_dir)
    final_status = status or ("COMPLETED" if
                totals["workers_completed"] == int(config["workers"])
                and not totals["chunks_failed"] else "PARTIAL")
    info = workbook.create_sheet("INFO")
    metadata = [
        ("RUN_ID", run_id), ("QUERY_ID", query_id),
        ("INSURANCE_CODE", code), ("INSURANCE_CLASS", QUERY_PROFILES[code]["name"]),
        ("CUBE", config["cube"]), ("CATALOG", config["catalog"]),
        ("WORKERS", config["workers"]), ("X_CHUNK", config["x_chunk"]),
        ("Y_CHUNK", config["y_chunk"]), ("STARTED_AT", started_at or ""),
        ("FINISHED_AT", finished_at or ""), ("EXPORTED_AT", utc_now()),
        ("STATUS", final_status), ("TOTAL_RECORDS", row_count),
        ("QUERY_HASH", query_hash(config)),
        ("NOTE", "OLAP aggregates; do not add All and detail totals"),
    ]
    t1, t2 = parse_time(started_at), parse_time(finished_at)
    if t1 and t2:
        metadata.append(("DURATION_SECONDS", max(0, int((t2 - t1).total_seconds()))))
    for axis, fields in (("X / ROWS", config["x"]),
                         ("Y / COLUMNS", config["y"]),
                         ("VALUES / MEASURES", config["measures"])):
        for index, field in enumerate(fields, 1):
            metadata.append((f"{axis} {index}", f"{display_map.get(field, field)} | {field}"))
    metadata.extend([
        ("SUCCESSFUL_CHUNKS", totals["chunks_ok"]),
        ("FAILED_CHUNKS", totals["chunks_failed"]),
        ("DATA_TYPE", "OLAP aggregate, not individual transactions"),
    ])
    for key, value in metadata:
        info.append([excel_text(key), excel_text(str(value))])
    try:
        workbook.save(tmp)
        os.replace(tmp, path)
    finally:
        tmp.unlink(missing_ok=True)
    return str(path), row_count


def selected_monitor(config: dict, run_dir: Path) -> int:
    start = utc_now()
    terminal = {"COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED"}
    with Live(console=console, refresh_per_second=1, screen=True, transient=False) as live:
        while True:
            workers, totals = selected_summary(config, run_dir)
            table = Table(title=f"OIC SELECTED QUERY | {config['code']} | {config['run_id']}",
                          box=rich_box.ROUNDED, expand=True)
            for col in ("Worker", "Status", "Current Chunk", "Last HTTP", "Last Event"):
                table.add_column(col)
            for worker in workers:
                table.add_row(str(worker.get("worker_id", "?")), worker.get("status", "STARTING"),
                              str(worker.get("current_request") or "—"),
                              str(worker.get("last_http") or "—"),
                              str(worker.get("last_message") or "—")[:70])
            status = (f"X / ROWS: {', '.join(config['x'])}\nY / COLUMNS: {', '.join(config['y'])}\n"
                      f"Measures: {len(config['measures'])} | Workers: {config['workers']}\n"
                      f"Chunks OK: {totals['chunks_ok']} | Cells: {totals['cells']:,} | "
                      f"Bytes: {totals['bytes']:,} | Elapsed: {format_duration(elapsed_seconds(start))}\n"
                      "Chunk total unknown until exploration finishes; percentages intentionally omitted.")
            live.update(Group(Panel(status, title="RUN OVERVIEW"), table), refresh=True)
            write_json(run_dir / "master_manifest.json",
                       dict(run_id=config["run_id"], updated_at=utc_now(),
                            summary=totals, workers=workers))
            if all(w.get("status") in terminal for w in workers) and elapsed_seconds(start) >= 8:
                break
            time.sleep(1)
    try:
        destination, rows = selected_export(config, run_dir)
        log(f"EXCEL: {destination} ({rows:,} rows)", level="OK")
    except Exception as exc:
        log(f"EXCEL ERROR: {exc}", level="ERROR")
    return 0


def selected_launch(config: dict, run_dir: Path, config_path: Path) -> int:
    if os.name != "nt":
        raise RuntimeError("Selected-query launch requires Windows")
    # Single-query compatibility: one visible PowerShell monitor window.
    # Workers are launched directly in the monitor process.
    launch_terminal_window(
        [[str(Path(sys.executable).resolve()), str(Path(__file__).resolve()),
          "--query-config", str(config_path.resolve()), "--monitor-only", "--run-workers"]],
        title="OIC | Live Monitor", cwd=HERE, keep_open=True)
    return 0


def selected_monitor_with_workers(config: dict, run_dir: Path, config_path: Path) -> int:
    processes = []
    streams = []
    try:
        for wid in range(1, int(config["workers"]) + 1):
            folder = run_dir / f"worker_{wid:02d}"
            folder.mkdir(parents=True, exist_ok=True)
            stream = (folder / "worker.log").open("ab")
            streams.append(stream)
            processes.append(subprocess.Popen(
                [str(Path(sys.executable).resolve()), str(Path(__file__).resolve()),
                 "--query-config", str(config_path), "--worker-id", str(wid)],
                stdout=stream, stderr=subprocess.STDOUT, cwd=str(HERE),
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)))
        # selected_monitor expects terminal worker states and exports when finished.
        return selected_monitor(config, run_dir)
    finally:
        for proc in processes:
            proc.wait()
        for stream in streams:
            stream.close()


def selected_main(argv) -> int:
    parser = argparse.ArgumentParser(description="Run an explicit OIC XMLA query or batch")
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--query-config", type=Path)
    group.add_argument("--batch-dir", type=Path)
    parser.add_argument("--launch", action="store_true")
    parser.add_argument("--batch-monitor", action="store_true")
    parser.add_argument("--worker-id", type=int, default=0)
    parser.add_argument("--monitor-only", action="store_true")
    parser.add_argument("--run-workers", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--export-only", action="store_true")
    args = parser.parse_args(argv)
    if args.batch_dir:
        if args.worker_id or args.monitor_only or args.export_only or (args.launch == args.batch_monitor):
            parser.error("Use --batch-dir DIR (--launch | --batch-monitor)")
        return batch_monitor(args.batch_dir) if args.batch_monitor else batch_run(args.batch_dir)
    config, run_dir = selected_paths(args.query_config)
    if args.worker_id:
        return selected_worker(args)
    if args.monitor_only:
        return (selected_monitor_with_workers(config, run_dir, args.query_config.resolve())
                if args.run_workers else selected_monitor(config, run_dir))
    if args.export_only:
        path, rows = selected_export(config, run_dir)
        log(f"EXPORTED {rows:,} rows to {path}", level="OK")
        return 0
    if args.launch:
        return selected_launch(config, run_dir, args.query_config)
    parser.error("Specify --launch, --worker-id, --monitor-only or --export-only")
    return 2


def main() -> int:

    if "--query-config" in sys.argv or "--batch-dir" in sys.argv:
        return selected_main(sys.argv[1:])

    if len(sys.argv) == 1:
        log("Use START_OIC.bat to select a query; full hierarchy collection is not the default.", level="WARN")
        return 2

    parser = build_parser()

    args = parser.parse_args()

    # ========================================================
    # VALIDATION
    # ========================================================

    if args.workers < 1:
        parser.error("--workers must be >= 1")

    if args.years_per_request < 1:
        parser.error(
            "--years-per-request must be >= 1"
        )

    if args.members_per_request < 1:
        parser.error(
            "--members-per-request must be >= 1"
        )

    if args.timeout < 1:
        parser.error("--timeout must be >= 1")

    if args.delay < 0:
        parser.error("--delay must be >= 0")

    if args.max_response_mb < 1:
        parser.error(
            "--max-response-mb must be >= 1"
        )

    if args.max_requests_per_job < 1:
        parser.error(
            "--max-requests-per-job must be >= 1"
        )

    if args.max_years < 0:
        parser.error("--max-years must be >= 0")

    if args.max_jobs < 0:
        parser.error("--max-jobs must be >= 0")

    args.source = (
        args.source
        .expanduser()
        .resolve()
    )

    output = (
        args.output_dir
        .expanduser()
        .resolve()
    )

    # ========================================================
    # DEFAULT OPERATION
    # ========================================================

    selected = any([
        args.prepare,
        args.launch,
        args.worker_id is not None,
        args.status,
        args.merge,
        args.export_only,
    ])

    if not selected:
        args.launch = True

    try:

        # ====================================================
        # PREPARE / LOAD JOB PLAN
        # ====================================================

        plan = prepare_plan(
            output=output,
            source=args.source,
            workers=args.workers,
        )

        # ====================================================
        # PREPARE
        # ====================================================

        if args.prepare:

            print_plan(plan)

            console.print(
                f"\n[dim]Plan: "
                f"{output / 'job_plan.json'}[/dim]"
            )

            return 0

        # ====================================================
        # STATUS
        # ====================================================

        if args.status:

            return show_status(
                output,
                plan,
            )

        # ====================================================
        # MERGE
        # ====================================================

        if args.merge:

            merge_manifests(
                output,
                plan,
            )

            return 0

        # ====================================================
        # EXPORT
        # ====================================================

        if args.export_only:

            return export_existing(
                output,
                plan,
            )

        # ====================================================
        # INDIVIDUAL WORKER
        # ====================================================

        if args.worker_id is not None:

            return run_worker(
                args,
                output,
                plan,
            )

        # ====================================================
        # LAUNCH 4 TABS + ORIGINAL TERMINAL MONITOR
        # ====================================================

        if args.launch:

            console.print()

            console.print(
                Panel(
                    Group(
                        Align.center(
                            Text(
                                APP_NAME,
                                style="bold bright_cyan",                            )
                        ),
                        Align.center(
                            Text(
                                "PARALLEL EXTRACTION",
                                style="bold white",
                            )
                        ),
                        Align.center(
                            Text(
                                f"{plan['total_hierarchies']} "
                                "HIERARCHIES"
                            )
                        ),
                        Align.center(
                            Text(
                                f"{plan['workers']} WORKER TABS",
                                style="bold green",
                            )
                        ),
                    ),
                    border_style="cyan",
                )
            )

            log(
                "Starting background workers via PowerShell"
            )

            log(
                f"Creating {plan['workers']} worker processes"
            )

            run_id = launch_worker_tabs(
                args,
                output,
                plan,
            )

            log(
                "Worker processes launched",
                level="OK",
            )

            log(
                "Starting Rich Live Monitor"
            )

            return monitor_workers(
                output,
                plan,
                run_id,
            )

        return 0

    except KeyboardInterrupt:

        log(
            "Interrupted",
            level="WARN",
        )

        return 130

    except Exception as exc:

        log(
            f"ERROR: {exc}",
            level="ERROR",
        )

        return 1


if __name__ == "__main__":
    sys.exit(main())
