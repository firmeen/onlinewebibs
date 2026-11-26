#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC PARALLEL COLLECTOR
======================

XMLA / MDX / OLAP Aggregate Collector
Rich Terminal Monitor + Windows Terminal worker tabs
4 Parallel Workers + Raw XML + Excel + Resume

INSTALL
-------
python -m pip install curl_cffi openpyxl rich

COMMANDS
--------
Prepare jobs:
    python oic_parallel_collector.py --prepare

Start a PowerShell worker supervisor:
    python oic_parallel_collector.py

Or:
    python oic_parallel_collector.py --launch

Workers run in visible Windows Terminal tabs. Each has its own Python process.

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
Windows Terminal (wt.exe) and PowerShell
Python 3.10+
curl_cffi
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
from datetime import datetime, timezone
from pathlib import Path
from typing import Any
from urllib.parse import urlparse
from xml.etree import ElementTree as ET

from curl_cffi import requests as curl_requests
from curl_cffi.requests.exceptions import RequestException as CurlRequestError

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
_selected_log_path: Path | None = None


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

    stamp = local_time()
    line = Text()
    line.append(stamp + " ", style="dim")
    line.append(prefix + " ", style="bold")
    line.append(str(message), style=color)
    console.print(line)
    if _selected_log_path is not None:
        try:
            with _selected_log_path.open("a", encoding="utf-8") as handle:
                handle.write(f"{stamp} {prefix} {level} {message}\n")
        except OSError:
            console.print("Worker log file could not be written.", style="yellow")


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

    temporary = path.with_name(
        path.name + "." + uuid.uuid4().hex + ".tmp"
    )
    try:
        temporary.write_text(
            json.dumps(data, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        for attempt in range(5):
            try:
                temporary.replace(path)
                break
            except PermissionError:
                if attempt == 4:
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

        self.session = curl_requests.Session(verify=True)

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
        started = time.perf_counter()
        metrics = dict(execute_ms=None, token_refresh_ms=0.0,
                       header_wait_ms=None, first_body_ms=None, receive_ms=None,
                       file_write_ms=None, finalize_ms=None, http_attempts=0,
                       http_status=None)
        # Retain partial timing on exceptions without changing exception types.
        self.last_execute_metrics = metrics
        try:
            if jwt_exp(self.token) <= int(time.time()) + 60:
                refresh_started = time.perf_counter()
                try:
                    self.refresh_token()
                finally:
                    metrics["token_refresh_ms"] += (time.perf_counter() - refresh_started) * 1000

            for attempt in range(2):
                if attempt:
                    refresh_started = time.perf_counter()
                    try:
                        self.refresh_token()
                    finally:
                        metrics["token_refresh_ms"] += (time.perf_counter() - refresh_started) * 1000

                headers = {
                    "Accept": "text/plain, */*; q=0.01",
                    "Content-Type": "text/xml",
                    "X-Requested-With": "XMLHttpRequest",
                    "Origin": BASE_URL,
                    "Referer": self.referer(),
                    "Authorization": "Bearer " + self.token,
                }
                metrics.update(header_wait_ms=None, first_body_ms=None, receive_ms=None,
                               file_write_ms=None, finalize_ms=None)
                metrics["http_attempts"] += 1
                transfer_started = time.perf_counter()
                with self.session.stream(
                    "POST", XMLA_URL, data=request_body, headers=headers,
                    timeout=(15, self.timeout), allow_redirects=False,
                ) as response:
                    metrics["header_wait_ms"] = (time.perf_counter() - transfer_started) * 1000
                    metrics["http_status"] = response.status_code
                    if response.status_code in (401, 403) and attempt == 0:
                        continue
                    temporary = destination.with_name(destination.name + ".partial")
                    size = 0
                    metrics["file_write_ms"] = 0.0
                    receive_started = time.perf_counter()
                    try:
                        try:
                            with temporary.open("wb") as file:
                                for chunk in response.iter_content():
                                    if not chunk:
                                        continue
                                    if metrics["first_body_ms"] is None:
                                        metrics["first_body_ms"] = (time.perf_counter() - transfer_started) * 1000
                                    size += len(chunk)
                                    if size > self.max_bytes:
                                        raise RuntimeError("XMLA response exceeds configured size limit")
                                    write_started = time.perf_counter()
                                    try:
                                        file.write(chunk)
                                    finally:
                                        metrics["file_write_ms"] += (time.perf_counter() - write_started) * 1000
                        finally:
                            metrics["receive_ms"] = (time.perf_counter() - receive_started) * 1000
                        finalize_started = time.perf_counter()
                        try:
                            temporary.replace(destination)
                        finally:
                            metrics["finalize_ms"] = (time.perf_counter() - finalize_started) * 1000
                    finally:
                        temporary.unlink(missing_ok=True)
                    metrics.update(content_type=response.headers.get("Content-Type", ""),
                                   response_bytes=size)
                    break
            else:
                raise RuntimeError("Authentication rejected after JWT refresh")
        finally:
            metrics["execute_ms"] = (time.perf_counter() - started) * 1000
        return dict(metrics)

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
                    CurlRequestError,
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
# WINDOWS TERMINAL LAUNCHER
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
    commands: list[list[str]], *,
    title: str,
    cwd: Path,
    tab_titles: list[str] | None = None,
    keep_open: bool | list[bool] = True,
    maximized: bool = False,
    window_id: str | None = None,
) -> str:
    """Launch independent visible tabs in one named Windows Terminal window."""
    if os.name != "nt":
        raise RuntimeError("Windows Terminal launcher requires Windows")
    wt = shutil.which("wt.exe")
    if not wt:
        raise RuntimeError("Windows Terminal (wt.exe) is required; no hidden-worker fallback")
    if not commands or any(not cmd for cmd in commands):
        raise ValueError("No terminal command")
    workdir = Path(cwd).resolve()
    if not workdir.is_dir():
        raise FileNotFoundError(str(workdir))
    titles = tab_titles if tab_titles is not None else [title] * len(commands)
    stay = [keep_open] * len(commands) if isinstance(keep_open, bool) else keep_open
    if len(titles) != len(commands) or len(stay) != len(commands):
        raise ValueError("One title and keep-open setting per terminal command required")
    shell = find_powershell()
    args = [wt, "-w", window_id or ("OIC-" + uuid.uuid4().hex)]
    if maximized:
        args.append("--maximized")
    for index, (command, tab_title, keep) in enumerate(zip(commands, titles, stay)):
        if index:
            args.append(";")
        # Encode the shell command so paths and semicolons are not parsed by WT.
        statement = "& " + " ".join(powershell_quote(str(v)) for v in command)
        if not keep:
            # The worker's real outcome is in state.json and its Python exit code.
            # Exit the shell cleanly so the default WT profile closes the tab.
            statement += "; exit 0"
        encoded = base64.b64encode(statement.encode("utf-16-le")).decode("ascii")
        args += ["new-tab", "--startingDirectory", str(workdir), "--title", tab_title,
                 "--suppressApplicationTitle", shell, "-NoLogo", "-NoProfile"]
        if keep:
            args.append("-NoExit")
        args += ["-EncodedCommand", encoded]
    subprocess.Popen(args, cwd=str(workdir), close_fds=True)
    return "WINDOWS TERMINAL"

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

def launch_worker_tabs(
    args: argparse.Namespace,
    output: Path,
    plan: dict,
) -> str:
    if os.name != "nt":
        raise RuntimeError("Worker terminal requires Windows")

    workers = int(plan["workers"])

    if workers < 1:
        raise ValueError("Worker count must be at least 1")

    run_id = (
        datetime.now().strftime("%Y%m%d_%H%M%S")
        + "_"
        + uuid.uuid4().hex[:6]
    )

    shell = find_powershell()

    commands = []
    titles = []

    for worker_id in range(1, workers + 1):
        worker_command = worker_powershell_command(
            args=args,
            worker_id=worker_id,
            output=output,
            run_id=run_id,
        )

        commands.append(
            [
                shell,
                "-NoLogo",
                "-NoProfile",
                "-Command",
                worker_command,
            ]
        )

        titles.append(
            f"OIC | Worker {worker_id:02d}"
        )

    mode = launch_terminal_window(
        commands,
        title="OIC | Parallel Workers",
        cwd=HERE,
        tab_titles=titles,
        keep_open=False,
    )

    log(
        f"Worker terminal: {mode} | "
        f"{workers} worker processes",
        level="OK",
    )

    return run_id


# ============================================================
# MASTER MANIFEST
# ============================================================

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
            "with Rich Live Monitor and PowerShell workers"
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
            "Launch background workers "
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

def build_selected_task_plan(config: dict) -> dict:
    """Immutable, deterministic worker allocations prepared at APPLY time.

    Exact chunk count depends on server member counts: discovery remains
    dynamic, but every X page already has exactly one worker owner.
    """
    validate_query(config)
    count = int(config["workers"])
    x_size = int(config["x_chunk"])
    return {
        "query_hash": query_hash(config),
        "query_id": config.get("query_id"),
        "code": config["code"],
        "cube": config["cube"],
        "cube_type": config.get("cube_type", "FY"),
        "created_at": utc_now(),
        "task_count": None,
        "task_count_status": "DYNAMIC",
        "assignments": [
            {"worker_id": wid,
             "x_offset_start": (wid - 1) * x_size,
             "x_offset_stride": count * x_size,
             "y_offset_start": 0,
             "y_offset_step": int(config["y_chunk"])}
            for wid in range(1, count + 1)
        ],
    }


def selected_paths(config_path: Path):
    config_path = config_path.resolve()
    run_dir = config_path.parent
    config = read_json(config_path)
    if not isinstance(config, dict):
        raise RuntimeError("Query config not found or invalid")
    validate_query(config)
    return config, run_dir


def _process_creation(kernel, handle) -> int:
    import ctypes
    from ctypes import wintypes

    times = [wintypes.FILETIME() for _ in range(4)]
    kernel.GetProcessTimes.argtypes = [wintypes.HANDLE] + [ctypes.POINTER(wintypes.FILETIME)] * 4
    kernel.GetProcessTimes.restype = wintypes.BOOL
    if not kernel.GetProcessTimes(handle, *(ctypes.byref(t) for t in times)):
        raise ctypes.WinError(ctypes.get_last_error())
    return (times[0].dwHighDateTime << 32) | times[0].dwLowDateTime


def _current_process_creation() -> int | None:
    if os.name != "nt":
        return None
    import ctypes
    from ctypes import wintypes
    kernel = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel.GetCurrentProcess.restype = wintypes.HANDLE
    return _process_creation(kernel, kernel.GetCurrentProcess())


class _WorkerProcess:
    """Hold a real Windows process handle, avoiding PID reuse on force-stop."""

    def __init__(self, pid: int, created: int | None, *, can_terminate: bool = False):
        import ctypes
        from ctypes import wintypes
        self.kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        k = self.kernel
        k.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
        k.OpenProcess.restype = wintypes.HANDLE
        k.CloseHandle.argtypes = [wintypes.HANDLE]
        k.WaitForSingleObject.argtypes = [wintypes.HANDLE, wintypes.DWORD]
        k.WaitForSingleObject.restype = wintypes.DWORD
        k.TerminateProcess.argtypes = [wintypes.HANDLE, wintypes.UINT]
        k.TerminateProcess.restype = wintypes.BOOL
        k.GetExitCodeProcess.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.DWORD)]
        k.GetExitCodeProcess.restype = wintypes.BOOL
        self.handle = k.OpenProcess(0x100000 | 0x1000 | (1 if can_terminate else 0), False, pid)
        if not self.handle:
            if ctypes.get_last_error() == 87:  # process no longer exists
                return
            raise ctypes.WinError(ctypes.get_last_error())
        try:
            if created is not None and _process_creation(k, self.handle) != created:
                self.close()
        except BaseException:
            self.close()
            raise

    def running(self) -> bool:
        return bool(self.handle and self.kernel.WaitForSingleObject(self.handle, 0) == 258)

    def exit_code(self) -> int | None:
        import ctypes
        from ctypes import wintypes
        code = wintypes.DWORD()
        if self.handle and self.kernel.GetExitCodeProcess(self.handle, ctypes.byref(code)):
            return int(code.value)
        return None

    def terminate(self) -> None:
        if self.running() and not self.kernel.TerminateProcess(self.handle, 1) and self.running():
            import ctypes
            raise ctypes.WinError(ctypes.get_last_error())

    def close(self) -> None:
        if self.handle:
            self.kernel.CloseHandle(self.handle)
            self.handle = None


def selected_worker(args) -> int:
    config, run_dir = selected_paths(args.query_config)
    code, worker_id = config["code"], args.worker_id
    worker_count = int(config["workers"])
    if not 1 <= worker_id <= worker_count:
        raise RuntimeError("Invalid worker ID")
    directory = run_dir / f"worker_{worker_id:02d}"
    chunks_dir = directory / "chunks"
    chunks_dir.mkdir(parents=True, exist_ok=True)
    manifest_path = directory / "manifest.json"
    state_path = directory / "state.json"
    heartbeat_path = directory / "heartbeat.json"
    attempt_id = getattr(args, "attempt_id", None)
    cancel_file = directory / f"cancel-{attempt_id}.flag" if attempt_id else None
    if attempt_id:
        initial = read_json(state_path, {}) or {}
        if (initial.get("attempt_id") != attempt_id or initial.get("status") != "STARTING"
                or initial.get("pid") or cancel_file.exists()):
            return 1
    global _selected_log_path
    _selected_log_path = directory / "worker.log"
    existing = read_json(manifest_path, {})
    manifest = existing if existing.get("run_id") == config["run_id"] and existing.get("query_hash") == query_hash(config) else {}
    manifest.setdefault("chunks", {})
    manifest["run_id"] = config["run_id"]
    manifest["query_hash"] = query_hash(config)
    manifest["worker_id"] = worker_id
    manifest["updated_at"] = utc_now()
    stop_file = run_dir / "stop.flag"
    batch_stop = run_dir.parent / "stop.flag"
    def stop_requested():
        return (stop_file.exists() or batch_stop.exists()
                or bool(cancel_file and cancel_file.exists()))
    shared_cooldown = run_dir.parent / "cooldown.json"
    def wait_for_server():
        while not stop_requested():
            info = read_json(shared_cooldown, {}) or {}
            remaining = float(info.get("until_ts", 0)) - time.time()
            if remaining <= 0:
                return
            state["status"] = "PAUSED"
            set_phase("COOLDOWN", "Shared server cooldown")
            time.sleep(min(1, remaining))
        return
    def backoff_server(http_status: int, attempt: int):
        old = read_json(shared_cooldown, {}) or {}
        until = max(float(old.get("until_ts", 0)),
                    time.time() + min(120, 5 * (2 ** attempt) + worker_id))
        write_json(shared_cooldown, {
            "until_ts": until, "http_status": http_status,
            "worker_id": worker_id, "updated_at": utc_now(),
        })
    planned = read_json(run_dir / "task_plan.json", {}) or {}
    if planned.get("query_hash") != query_hash(config):
        raise RuntimeError("Missing or stale apply-time task plan")
    allocations = {int(item["worker_id"]): item
                   for item in planned.get("assignments", [])}
    allocation = allocations.get(worker_id)
    if not allocation:
        raise RuntimeError("Worker was not assigned in task plan")
    heartbeat_stop = threading.Event()
    state = dict(run_id=config["run_id"], worker_id=worker_id, pid=os.getpid(),
                 attempt_id=attempt_id, process_created=_current_process_creation(),
                 status="STARTING", current_request=None, request_started_at=None,
                 last_http=None, last_message="Initializing", updated_at=utc_now(),
                 started_at=utc_now(), http_engine="curl_cffi",
                 requests_total=0, requests_finished=0,
                 request_time_total_ms=0.0, last_request_ms=None,
                 current_phase="INITIALIZING", current_x_offset=None, current_y_offset=None,
                 last_event_at=utc_now(), slow_requests=0, recent_requests=[],
                 last_execute_http=None, last_parse_ms=None, last_parse_source=None)
    timing_fields = ("header_wait_ms", "first_body_ms", "receive_ms", "file_write_ms",
                     "finalize_ms", "token_refresh_ms", "http_attempts")
    state.update({"last_" + name: None for name in timing_fields})
    def save_state():
        state["updated_at"] = utc_now()
        write_json(state_path, state)
    def set_phase(phase, message=None):
        state["current_phase"] = phase
        if phase != "REQUESTING":
            state["request_started_at"] = None
        if message is not None:
            state.update(last_message=message, last_event_at=utc_now())
        save_state()
    def heartbeat():
        while not heartbeat_stop.is_set():
            write_json(heartbeat_path, dict(run_id=config["run_id"], worker_id=worker_id,
                                            pid=os.getpid(), attempt_id=attempt_id,
                                            updated_at=utc_now()))
            heartbeat_stop.wait(2)
    if attempt_id:
        initial = read_json(state_path, {}) or {}
        if (cancel_file.exists() or initial.get("attempt_id") != attempt_id
                or initial.get("status") != "STARTING" or initial.get("pid")):
            return 1
    save_state()
    thread = threading.Thread(target=heartbeat, daemon=True)
    thread.start()
    client = None
    x_size, y_size = int(config["x_chunk"]), int(config["y_chunk"])
    errors = 0
    completed = False
    try:
        console.print(Panel(
            Text(f"OIC PARALLEL COLLECTOR\nWORKER {worker_id:02d} / {worker_count:02d}\n"
                 f"{config.get('query_id', 'Q001')} | {code}/{config.get('cube_type', 'FY')} | CURL_CFFI\n"
                 f"X start={allocation['x_offset_start']} stride={allocation['x_offset_stride']} | CHUNK={x_size}/{y_size}",
                 justify="center"), border_style="cyan"))
        log("Task plan loaded", worker_id, "OK")
        if stop_requested():
            state["status"] = "STOPPED"
            return 0
        env = read_env()
        cookie = env.get("OIC_COOKIE", "").strip()
        if not cookie:
            raise RuntimeError("SESSION_EXPIRED: OIC_COOKIE not configured")
        client = OICClient(cookie=cookie, token=env.get("OIC_BEARER_TOKEN", ""),
                           timeout=int(config.get("timeout", 60)),
                           max_response_mb=100, worker_id=worker_id, code=code)
        state["status"] = "RUNNING"
        set_phase("IDLE", "HTTP session ready")
        log("HTTP session initialized | curl_cffi", worker_id, "OK")
        # Deterministic X paging: a given X page belongs to exactly one worker.
        x_offset = int(allocation["x_offset_start"])
        while not stop_requested():
            y_offset = 0
            if (run_dir / "pause.flag").exists():
                state["status"] = "PAUSED"; set_phase("PAUSED", "Paused")
                while (run_dir / "pause.flag").exists() and not stop_requested():
                    time.sleep(1)
                state["status"] = "RUNNING"; set_phase("IDLE", "Resumed")
            if stop_requested():
                break
            x_returned = None
            while not stop_requested():
                key = f"{query_hash(config)}_x{x_offset:06d}_y{y_offset:06d}"
                req_path = chunks_dir / f"{key}.request.xml"
                resp_path = chunks_dir / f"{key}.response.xml"
                record = manifest["chunks"].get(key, {})
                data = None
                cached = False
                state.update(current_x_offset=x_offset, current_y_offset=y_offset,
                             current_request=key)
                chunk_http = None
                if record.get("status") == "OK" and resp_path.is_file():
                    set_phase("PARSING", f"Cached XML X={x_offset} Y={y_offset}")
                    parse_started = time.perf_counter()
                    try:
                        data = query_parse(resp_path.read_bytes())
                        cached = True
                    except (ValueError, ET.ParseError, OSError):
                        data = None
                    finally:
                        state.update(last_parse_ms=(time.perf_counter() - parse_started) * 1000,
                                     last_parse_source="CACHE")
                if data is None:
                    wait_for_server()
                    if stop_requested():
                        break
                    req_path.write_bytes(query_soap(config, x_offset, y_offset))
                    state.update(status="RUNNING", current_request=key)
                    http = None
                    for attempt in range(2):
                        if stop_requested():
                            break
                        try:
                            http = None
                            chunk_http = None
                            body = req_path.read_bytes()
                            client.last_execute_metrics = {}
                            state.update(status="RUNNING", request_started_at=utc_now(),
                                         last_execute_http=None, last_parse_ms=None,
                                         last_parse_source=None, requests_total=state["requests_total"] + 1)
                            set_phase("REQUESTING", f"REQUEST X={x_offset} Y={y_offset}")
                            log(f"REQUEST X={x_offset} Y={y_offset}", worker_id)
                            request_started = time.perf_counter()
                            sample = None
                            try:
                                http = client.execute(body, resp_path)
                            finally:
                                timing = http if http is not None else client.last_execute_metrics
                                duration = timing.get("execute_ms")
                                if duration is None:
                                    duration = (time.perf_counter() - request_started) * 1000
                                state["last_request_ms"] = round(duration, 2)
                                state["request_time_total_ms"] += duration
                                state["requests_finished"] += 1
                                chunk_http = timing.get("http_status")
                                state["last_execute_http"] = chunk_http
                                if chunk_http is not None:
                                    state["last_http"] = chunk_http
                                for name in timing_fields:
                                    state["last_" + name] = timing.get(name)
                                state["slow_requests"] += int(duration >= 5000)
                                sample = dict(completed_at=utc_now(), duration_ms=duration,
                                              http_status=chunk_http, parse_ms=None,
                                              outcome="OK" if http and chunk_http == 200 else "ERROR")
                                sample.update({name: timing.get(name) for name in timing_fields})
                                state["recent_requests"].append(sample)
                                del state["recent_requests"][:-100]
                                set_phase("IDLE", f"HTTP {chunk_http if chunk_http is not None else 'ERROR'} | {duration:.0f} ms")
                            log(f"HTTP {http['http_status']} | {duration:.0f} ms | "
                                f"{http['response_bytes']:,} bytes", worker_id,
                                "WARN" if duration >= 5000 or http["http_status"] != 200 else "OK")
                            if http["http_status"] == 200:
                                set_phase("PARSING", f"Parse XML X={x_offset} Y={y_offset}")
                                parse_started = time.perf_counter()
                                try:
                                    data = query_parse(resp_path.read_bytes())
                                except BaseException:
                                    sample["outcome"] = "ERROR"
                                    raise
                                finally:
                                    state["last_parse_ms"] = (time.perf_counter() - parse_started) * 1000
                                    state["last_parse_source"] = "HTTP"
                                    sample["parse_ms"] = state["last_parse_ms"]
                                    save_state()
                                break
                            if http["http_status"] in (401, 403):
                                raise RuntimeError(f"AUTH HTTP {http['http_status']}")
                            if http["http_status"] not in (408, 425, 429, 500, 502, 503, 504) or attempt:
                                raise RuntimeError(f"XMLA HTTP {http['http_status']}")
                            if http["http_status"] in (429, 503):
                                backoff_server(http["http_status"], attempt)
                            wait_for_server()
                            if stop_requested():
                                break
                            state["status"] = "RUNNING"; set_phase("IDLE")
                        except (CurlRequestError, ValueError, ET.ParseError, RuntimeError) as exc:
                            if attempt == 1 or (http and http.get("http_status") not in
                                                       (408, 425, 429, 500, 502, 503, 504)):
                                raise
                            state["status"] = "RETRYING"
                            set_phase("RETRYING", f"Retry: {exc}"[:200])
                            log(f"Transient error: {exc}", worker_id, "WARN")
                            if stop_requested():
                                break
                            time.sleep(min(30, 2 ** (attempt + 1) + worker_id))
                    if data is None:
                        if stop_requested():
                            break
                        raise RuntimeError("Empty XMLA dataset after attempts")
                    record = dict(status="OK", x_offset=x_offset, y_offset=y_offset,
                                  request_file=req_path.name, response_file=resp_path.name,
                                  x_count=len(data["axes"]["Axis0"]),
                                  y_tuples=len(data["axes"]["Axis1"]),
                                  cells=len(data["cells"]), http_status=200,
                                  response_bytes=resp_path.stat().st_size,
                                  completed_at=utc_now())
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
                valid_cells = sum(value is not None for value in data["cells"].values())
                kind = ("EMPTY_PAGE" if x_count == 0 or y_count == 0 else
                        "DATA" if valid_cells else "NO_CELLS")
                metadata = dict(x_members=x_count, y_members=y_count,
                                axis0_tuples=len(data["axes"]["Axis0"]),
                                axis1_tuples=len(data["axes"]["Axis1"]),
                                measure_count=len(config.get("measures", [])),
                                chunk_kind=kind, nonempty_cells=valid_cells)
                if not cached or any(record.get(k) != v for k, v in metadata.items()):
                    set_phase("SAVING", f"Save chunk X={x_offset} Y={y_offset}")
                    record.update(metadata)
                    manifest["chunks"][key] = record
                    manifest["updated_at"] = utc_now()
                    write_json(manifest_path, manifest)
                message = (f"{'CACHE ' if cached else ''}{kind} | X={x_count} Y={y_count} "
                           f"M={metadata['measure_count']} CELLS={valid_cells}")
                if not cached:
                    log(message, worker_id, "OK")
                state["current_request"] = None
                set_phase("IDLE", f"{kind} X={x_offset} Y={y_offset} | cells={valid_cells}")
                if x_count == 0 or y_count < y_size:
                    break
                y_offset += y_size
                set_phase("DELAY")
                time.sleep(float(config.get("delay", 1)))
            if x_returned is None or x_returned < x_size:
                completed = True
                break
            x_offset += int(allocation["x_offset_stride"])
            set_phase("DELAY")
            time.sleep(float(config.get("delay", 1)))
        state["status"] = "STOPPED" if stop_requested() else "COMPLETED" if completed else "PARTIAL"
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
                                          http_status=locals().get("chunk_http"),
                                          error=text_message[:300],
                                          completed_at=utc_now())
        has_success = any(c.get("status") == "OK" for c in manifest["chunks"].values())
        state["status"] = ("SESSION_EXPIRED" if any(marker in text_message.upper() for marker in ("SESSION_EXPIRED", "COOKIE SESSION EXPIRED", "AUTH HTTP 401", "AUTH HTTP 403"))
                           else "PARTIAL" if has_success else "FAILED")
        state.update(last_message=text_message[:200], last_event_at=utc_now())
        log(f"FAILED: {text_message}", worker_id, "ERROR")
    finally:
        if client:
            client.close()
        heartbeat_stop.set()
        thread.join(timeout=3)
        state["current_request"] = None
        state["request_started_at"] = None
        manifest["status"] = state["status"]
        manifest["updated_at"] = utc_now()
        write_json(manifest_path, manifest)
        state["exit_code"] = 1 if errors else 0
        state["current_phase"] = "FINISHED"
        state["last_message"] = f"{state['status']}: {state.get('last_message', '')}"[:200]
        state["last_event_at"] = utc_now()
        save_state()
        log(state["last_message"], worker_id,
            "OK" if state["status"] == "COMPLETED" else "WARN")
    return 1 if errors else 0


WORKER_TERMINAL = {"COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED"}


def selected_summary(config: dict, run_dir: Path):
    workers = []
    chunk_fields = ("chunks_ok", "chunks_failed", "data_chunks", "empty_pages",
                    "no_cell_chunks", "unclassified_chunks", "cells", "bytes")
    totals = dict.fromkeys(chunk_fields, 0)
    totals.update(requests_total=0, requests_finished=0, request_time_total_ms=0.0,
                  slow_requests=0)
    now = datetime.now(timezone.utc)
    recent = []
    starts = []
    history_limited = False
    expected_hash = query_hash(config)
    kinds = {"DATA": "data_chunks", "EMPTY_PAGE": "empty_pages", "NO_CELLS": "no_cell_chunks"}
    for wid in range(1, int(config["workers"]) + 1):
        directory = run_dir / f"worker_{wid:02d}"
        state = read_json(directory / "state.json", {}) or {}
        heartbeat = read_json(directory / "heartbeat.json", {}) or {}
        if state.get("run_id") != config["run_id"]:
            state = dict(status="STARTING", worker_id=wid, run_id=config["run_id"])
        if (heartbeat.get("run_id") == config["run_id"]
                and heartbeat.get("attempt_id") == state.get("attempt_id")
                and heartbeat.get("pid") == state.get("pid")):
            last = parse_time(heartbeat.get("updated_at"))
        else:
            last = None
        state["heartbeat_age"] = elapsed_seconds(last.isoformat()) if last else None
        if state.get("status") not in WORKER_TERMINAL and state.get("pid"):
            if (last and state["heartbeat_age"] > 10) or (
                    not last and elapsed_seconds(state.get("started_at")) > 10):
                state["status"] = "STALE"
        state.update(dict.fromkeys(chunk_fields, 0))
        manifest = read_json(directory / "manifest.json", {}) or {}
        if (manifest.get("run_id") == config["run_id"]
                and manifest.get("query_hash") == expected_hash):
            for chunk in manifest.get("chunks", {}).values():
                if chunk.get("status") == "OK":
                    state["chunks_ok"] += 1
                    state[kinds.get(chunk.get("chunk_kind"), "unclassified_chunks")] += 1
                    state["cells"] += int(chunk.get("cells") or 0)
                    state["bytes"] += int(chunk.get("response_bytes") or 0)
                elif chunk.get("status") == "FAILED":
                    state["chunks_failed"] += 1
        for name in (*chunk_fields, "requests_total", "requests_finished",
                     "request_time_total_ms", "slow_requests"):
            totals[name] += state.get(name) or 0
        state["current_request_ms"] = None
        if (state.get("status") not in WORKER_TERMINAL
                and state.get("current_phase") == "REQUESTING"):
            start = parse_time(state.get("request_started_at"))
            if start:
                state["current_request_ms"] = max(0.0, (now - start).total_seconds() * 1000)
        started = parse_time(state.get("started_at"))
        if started:
            starts.append(start)
        samples = state.get("recent_requests") or []
        worker_recent = []
        for sample in samples:
            ended = parse_time(sample.get("completed_at"))
            if ended and 0 <= (now - ended).total_seconds() <= 60:
                worker_recent.append(sample)
        if len(samples) >= 100 and len(worker_recent) == len(samples):
            history_limited = True
        recent.extend(worker_recent)
        workers.append(state)
    maximum = max((w["chunks_ok"] for w in workers), default=0)
    totals["max_worker_chunks_ok"] = maximum
    for w in workers:
        w["relative_progress"] = w["chunks_ok"] / maximum if maximum else 0.0
    for status in ("RUNNING", "RETRYING", "PAUSED", "STARTING", "STALE", "COMPLETED", "STOPPED"):
        totals["workers_" + status.lower()] = sum(w.get("status") == status for w in workers)
    totals["workers_failed"] = sum(w.get("status") in {"PARTIAL", "FAILED", "SESSION_EXPIRED"} for w in workers)
    totals["workers_terminal"] = sum(w.get("status") in WORKER_TERMINAL for w in workers)
    totals["average_request_ms"] = (totals["request_time_total_ms"] / totals["requests_finished"]
                                    if totals["requests_finished"] else None)
    durations = sorted(float(v["duration_ms"]) for v in recent if v.get("duration_ms") is not None)
    totals["recent_sample_count"] = len(durations)
    totals["recent_average_ms"] = sum(durations) / len(durations) if durations else None
    totals["recent_p95_ms"] = durations[math.ceil(.95 * len(durations)) - 1] if durations else None
    totals["recent_slow_requests"] = sum(v >= 5000 for v in durations)
    totals["slow_inflight_workers"] = sum((w.get("current_request_ms") or 0) >= 5000 for w in workers)
    totals["recent_history_limited"] = history_limited
    # Freeze the attempt clock once every worker has reached a terminal state.
    ended = now
    if workers and totals["workers_terminal"] == len(workers):
        ends = [parse_time(w.get("updated_at")) for w in workers]
        ended = max((v for v in ends if v), default=now)
    elapsed = max(0.0, (ended - min(starts)).total_seconds()) if starts else 0.0
    observed = min(60.0, max(0.0, (now - min(starts)).total_seconds())) if starts else 0.0
    totals["overall_request_rate"] = totals["requests_total"] / elapsed if elapsed else None
    totals["recent_request_rate"] = len(durations) / observed if observed and durations else None
    for name in ("header_wait_ms", "first_body_ms", "receive_ms", "file_write_ms",
                 "finalize_ms", "token_refresh_ms", "parse_ms"):
        values = [float(v[name]) for v in recent if v.get(name) is not None]
        totals["recent_" + name] = sum(values) / len(values) if values else None
    return workers, totals

def selected_export(config: dict, run_dir: Path) -> tuple[str, int]:
    if Workbook is None:
        raise RuntimeError("openpyxl required for Excel export")
    output_root = Path(config.get("output_root", run_dir)).resolve()
    excel_dir = output_root / "Excel"
    excel_dir.mkdir(parents=True, exist_ok=True)
    query_id = str(config.get("query_id", "Q001"))
    cube_type = str(config.get("cube_type", "FY"))
    filename = f"{config['run_id']}_{config['code']}_{cube_type}_{query_id}.xlsx"
    path = excel_dir / filename
    tmp = excel_dir / (filename + ".tmp")
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("DATA")
    x_count, y_count = len(config["x"]), len(config["y"])
    headers = ["INSURANCE_CODE", "CUBE", "CATALOG"]
    for axis, count in (("X", x_count), ("Y", y_count)):
        for i in range(count):
            headers += [f"{axis}{i+1}_HIERARCHY", f"{axis}{i+1}_FIELD_CAPTION",
                        f"{axis}{i+1}_MEMBER_CAPTION", f"{axis}{i+1}_UNIQUE_NAME",
                        f"{axis}{i+1}_LEVEL"]
    headers += ["MEASURE", "MEASURE_CAPTION", "VALUE", "LAST_DATA_UPDATE", "RETRIEVED_AT"]
    display_map = {f.get("mdx"): f.get("caption", f.get("mdx"))
                   for f in config.get("selected_fields", []) if isinstance(f, dict)}
    sheet.append(headers)
    row_count, sheet_rows, sheet_number = 0, 1, 1
    for wid in range(1, int(config["workers"]) + 1):
        directory = run_dir / f"worker_{wid:02d}"
        manifest = read_json(directory / "manifest.json", {}) or {}
        if manifest.get("run_id") != config["run_id"] or manifest.get("query_hash") != query_hash(config):
            continue
        for name, chunk in sorted(manifest.get("chunks", {}).items()):
            if chunk.get("status") != "OK":
                continue
            response = directory / "chunks" / chunk["response_file"]
            if not response.is_file():
                continue
            data = query_parse(response.read_bytes())
            for record in query_excel_records(config, data):
                if sheet_rows >= EXCEL_MAX_ROWS:
                    sheet_number += 1
                    sheet = workbook.create_sheet(f"Data_{sheet_number}")
                    sheet.append(headers)
                    sheet_rows = 1
                entry = [config["code"], config["cube"], config["catalog"]]
                for values, names in ((record["X"], config["x"]), (record["Y"], config["y"])):
                    for index, field in enumerate(names):
                        matching = next((m for m in values if m["hierarchy"] == field), {})
                        entry += [field, display_map.get(field, field), matching.get("caption"),
                                  matching.get("unique_name"), matching.get("level")]
                raw = record["VALUE"]
                try:
                    number = float(raw)
                    value = number if math.isfinite(number) else raw
                except (TypeError, ValueError):
                    value = raw
                entry += [record["MEASURE"], display_map.get(record["MEASURE"], record["MEASURE"]),
                          value, record["LAST_DATA_UPDATE"], chunk.get("completed_at")]
                sheet.append([excel_text(v) if isinstance(v, str) else v for v in entry])
                row_count += 1
                sheet_rows += 1
    workers, totals = selected_summary(config, run_dir)
    info = workbook.create_sheet("INFO")
    for key, val in (("RUN_ID", config["run_id"]),
                     ("QUERY_ID", config.get("query_id", "Q001")),
                     ("INSURANCE_CODE", config["code"]),
                     ("CUBE_TYPE", config.get("cube_type", "FY")),
                     ("CUBE", config["cube"]),
                     ("CATALOG", config["catalog"]),
                     ("STARTED_AT", config.get("created_at", "")),
                     ("EXPORTED_AT", utc_now()),
                     ("QUERY_HASH", query_hash(config)),
                     ("X_ROWS", " | ".join(config["x"])),
                     ("Y_COLUMNS", " | ".join(config["y"])),
                     ("VALUES", " | ".join(config["measures"])),
                     ("WORKERS", config["workers"]),
                     ("X_CHUNK", config["x_chunk"]),
                     ("Y_CHUNK", config["y_chunk"]),
                     ("STATUS", "COMPLETED" if totals["workers_completed"] == config["workers"] else "PARTIAL"),
                     ("OUTPUT_ROWS", row_count),
                     ("NOTE", "OLAP aggregates; do not add All and detail totals")):
        info.append([key, val])
    workbook.save(tmp)
    tmp.replace(path)
    return str(path), row_count


def prepare_worker_tabs(config: dict, run_dir: Path, config_path: Path,
                        attempt_id: str) -> tuple[list[list[str]], list[str]]:
    """Prepare one independent Python command and fresh state per assigned worker."""
    commands, titles = [], []
    python = str(Path(sys.executable).resolve())
    script = str(Path(__file__).resolve())
    worker_count = int(config["workers"])
    # Never start a second copy of a surviving worker after a supervisor crash.
    for wid in range(1, worker_count + 1):
        directory = run_dir / f"worker_{wid:02d}"
        old = read_json(directory / "state.json", {}) or {}
        if (old.get("run_id") == config["run_id"] and old.get("pid")
                and old.get("status") not in WORKER_TERMINAL):
            proc = _WorkerProcess(int(old["pid"]), old.get("process_created"))
            try:
                if proc.running():
                    raise RuntimeError(f"Worker {wid:02d} from a prior attempt is still running")
            finally:
                proc.close()
    for wid in range(1, worker_count + 1):
        directory = run_dir / f"worker_{wid:02d}"
        directory.mkdir(parents=True, exist_ok=True)
        (directory / "heartbeat.json").unlink(missing_ok=True)
        write_json(directory / "state.json", dict(
            run_id=config["run_id"], worker_id=wid, attempt_id=attempt_id,
            status="STARTING", updated_at=utc_now(), last_message="Launching",
            http_engine="curl_cffi"))
        commands.append([python, script, "--query-config", str(config_path),
                         "--worker-id", str(wid), "--attempt-id", attempt_id])
        titles.append(f"{config.get('query_id', 'Q001')} | WORKER {wid:02d}")
    return commands, titles


def wait_for_query_workers(config: dict, run_dir: Path, attempt_id: str, *,
                           control_dir: Path, on_tick=None, on_stopping=None) -> None:
    """Supervise actual Python process handles, never the wt.exe launcher."""
    pending = set(range(1, int(config["workers"]) + 1))
    handles = {}
    launch_deadline = time.monotonic() + 60
    stopping_since = None
    stop = control_dir / "stop.flag"
    force = control_dir / "force_stop.flag"

    def finish(wid, state, message, *, stopped=False):
        directory = run_dir / f"worker_{wid:02d}"
        (directory / f"cancel-{attempt_id}.flag").touch()
        state.update(run_id=config["run_id"], worker_id=wid, attempt_id=attempt_id,
                     status="STOPPED" if stopped else "FAILED",
                     last_message=message, current_request=None, request_started_at=None,
                     updated_at=utc_now())
        write_json(directory / "state.json", state)
        pending.discard(wid)

    try:
        while pending:
            stopping = stop.exists() or force.exists()
            if stopping and stopping_since is None:
                stopping_since = time.monotonic()
                (run_dir / "stop.flag").touch()
                if on_stopping:
                    on_stopping()
            killing = force.exists() or (stopping_since is not None and
                time.monotonic() - stopping_since > max(30, int(config.get("timeout", 60)) + 30))
            for wid in list(pending):
                directory = run_dir / f"worker_{wid:02d}"
                state = read_json(directory / "state.json", {}) or {}
                if state.get("attempt_id") != attempt_id:
                    state = {}
                if wid not in handles and state.get("pid") and state.get("process_created"):
                    handles[wid] = _WorkerProcess(int(state["pid"]),
                                                  int(state["process_created"]), can_terminate=True)
                proc = handles.get(wid)
                if proc is not None:
                    if proc.running():
                        if killing:
                            proc.terminate()
                        continue
                    code = proc.exit_code()
                    # The final state can be written between the first read and process exit.
                    latest = read_json(directory / "state.json", {}) or {}
                    if latest.get("attempt_id") == attempt_id:
                        state = latest
                    if state.get("status") not in WORKER_TERMINAL:
                        finish(wid, state, f"Worker exited without final state (exit={code})",
                               stopped=stopping)
                    elif code not in (None, 0) and state.get("status") == "COMPLETED":
                        finish(wid, state, f"Worker exited with code {code}", stopped=stopping)
                    else:
                        pending.discard(wid)
                    proc.close()
                    del handles[wid]
                elif state.get("status") in WORKER_TERMINAL:
                    pending.discard(wid)
                elif stopping or time.monotonic() >= launch_deadline:
                    (directory / f"cancel-{attempt_id}.flag").touch()
                    latest = read_json(directory / "state.json", {}) or {}
                    if (latest.get("attempt_id") == attempt_id and latest.get("pid")
                            and latest.get("process_created")):
                        continue  # registration raced cancellation; supervise its exit
                    finish(wid, state, "Worker cancelled before startup" if stopping else
                           "Worker tab did not register within 60s", stopped=stopping)
            if on_tick:
                on_tick()
            if pending:
                time.sleep(0.5)
    except BaseException:
        # Prevent late launches or orphan workers if the supervisor fails.
        (run_dir / "stop.flag").touch()
        for wid in pending:
            directory = run_dir / f"worker_{wid:02d}"
            (directory / f"cancel-{attempt_id}.flag").touch()
        for proc in handles.values():
            proc.terminate()
        raise
    finally:
        for proc in handles.values():
            proc.close()


def selected_monitor(config: dict, run_dir: Path, attempt_id: str | None = None) -> int:
    start = utc_now()
    qid = config.get("query_id", "Q001")
    entry = dict(query_id=qid, code=config["code"], cube_type=config.get("cube_type", "FY"),
                 workers=config["workers"], status="RUNNING", started_at=start)
    with Live(console=console, refresh_per_second=1, screen=True, auto_refresh=False) as live:
        def tick():
            workers, totals = selected_summary(config, run_dir)
            live.update(render_batch_monitor(
                dict(run_id=config["run_id"], status=entry["status"], current_query=qid,
                     started_at=start, queries=[entry]), config, workers, totals,
                read_json(run_dir.parent / "cooldown.json", {}) or {}), refresh=True)
            write_json(run_dir / "master_manifest.json", dict(
                run_id=config["run_id"], updated_at=utc_now(), summary=totals, workers=workers))
            return totals
        if attempt_id:
            wait_for_query_workers(config, run_dir, attempt_id, control_dir=run_dir, on_tick=tick)
        else:
            while tick()["workers_terminal"] < int(config["workers"]):
                time.sleep(1)
        _, totals = selected_summary(config, run_dir)
        entry["status"] = "COMPLETED" if totals["workers_completed"] == config["workers"] else (
            "STOPPED" if (run_dir / "stop.flag").exists() else "PARTIAL" if totals["chunks_ok"] else "FAILED")
        entry["finished_at"] = utc_now()
        try:
            destination, rows = selected_export(config, run_dir)
            entry.update(excel_path=destination, rows=rows)
        except Exception as exc:
            entry.update(status="PARTIAL" if totals["chunks_ok"] else "FAILED", error=f"EXCEL: {exc}")
        tick()
    log(f"{entry['status']} | EXCEL: {entry.get('excel_path', '—')} | rows={entry.get('rows', 0)}",
        level="OK" if entry["status"] == "COMPLETED" else "WARN")
    return 0 if entry["status"] == "COMPLETED" else 1


def selected_launch(config: dict, run_dir: Path, config_path: Path) -> int:
    attempt_id = uuid.uuid4().hex
    commands, titles = prepare_worker_tabs(config, run_dir, config_path, attempt_id)
    monitor = [str(Path(sys.executable).resolve()), str(Path(__file__).resolve()),
               "--query-config", str(config_path), "--monitor-only", "--attempt-id", attempt_id]
    launch_terminal_window([monitor, *commands], title="OIC | Execution", cwd=HERE,
                           tab_titles=["LIVE MONITOR", *titles],
                           keep_open=False, maximized=True)
    return 0

# ============================================================
# BATCH SUPERVISOR — one execution window, sequential queries,
# independent visible parallel worker tabs.
# ============================================================

BATCH_TERMINAL = {"COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED", "PENDING"}

def _batch_snapshot(batch_dir: Path, *, status: str, current: str | None,
                    entries: list[dict], error: str = "") -> None:
    previous = read_json(batch_dir / "batch_manifest.json", {}) or {}
    started = utc_now() if status == "STARTING" else previous.get("started_at", utc_now())
    write_json(batch_dir / "batch_manifest.json", {
        "run_id": batch_dir.name, "pid": os.getpid(), "started_at": started,
        "status": status, "current_query": current,
        "queries": entries, "error": error,
        "updated_at": utc_now(),
        "finished_at": utc_now() if status in ("COMPLETED", "PARTIAL", "FAILED", "STOPPED") else None,
    })

def _update_history(root: Path, record: dict) -> None:
    path = root / "history.csv"
    columns = ("run_id", "query_id", "code", "cube_type", "cube",
               "status", "started_at", "finished_at", "rows", "excel_path", "error")
    rows = []
    if path.is_file():
        with path.open("r", encoding="utf-8-sig", newline="") as f:
            rows = list(csv.DictReader(f))
    rows = [row for row in rows
            if (row.get("run_id"), row.get("query_id")) !=
            (record["run_id"], record["query_id"])]
    rows.append(record)
    temp = path.with_suffix(".tmp")
    with temp.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)
    temp.replace(path)

def render_batch_monitor(manifest: dict, config: dict, workers: list, totals: dict,
                         cooldown: dict):
    entries = manifest.get("queries", [])
    current = manifest.get("current_query")
    entry = next((e for e in entries if e.get("query_id") == current), {})
    if not entry and entries:
        entry = next((e for e in reversed(entries) if e.get("started_at")), {})
    status = manifest.get("status", "STARTING")
    colors = dict(COMPLETED="green", RUNNING="bright_cyan", RETRYING="yellow",
                  PAUSED="yellow", STALE="yellow", FAILED="red", PARTIAL="yellow",
                  SESSION_EXPIRED="red", STOPPED="yellow", STOPPING="yellow",
                  STARTING="cyan", EXPORTING="cyan", PENDING="dim")
    def styled(value):
        return Text(str(value), style=colors.get(str(value), "white"))
    def panel(value, title, style="cyan"):
        return Panel(value, title=title, border_style=style, padding=(0, 1))
    def short(value, length=65):
        return str(value or "—")[:length]
    def elapsed(start, end=None):
        a, b = parse_time(start), parse_time(end)
        return max(0, int(((b or datetime.now(timezone.utc)) - a).total_seconds())) if a else 0
    done = sum(e.get("status") == "COMPLETED" for e in entries)
    failed = sum(e.get("status") in ("FAILED", "PARTIAL", "SESSION_EXPIRED") for e in entries)
    stopped = sum(e.get("status") == "STOPPED" for e in entries)
    waiting = sum(e.get("status") == "PENDING" for e in entries)
    n = len(entries)
    coverage = done / n if n else 0.0
    query_elapsed = elapsed(entry.get("started_at"), entry.get("finished_at"))
    calls = int(totals.get("requests_total") or 0)
    avg = totals.get("average_request_ms")
    rate = f"{calls / query_elapsed:.2f}/s" if query_elapsed else "—"
    cool = max(0, int(float(cooldown.get("until_ts") or 0) - time.time()))
    bytes_mb = int(totals.get("bytes") or 0) / 1024 ** 2
    engine = ", ".join(sorted({str(w["http_engine"]) for w in workers if w.get("http_engine")})) or "—"
    active = sum(totals.get("workers_" + k, 0) for k in ("running", "retrying", "paused"))
    running = totals.get("workers_running", 0)
    retrying = totals.get("workers_retrying", 0)
    paused = totals.get("workers_paused", 0)
    count = int(config.get("workers") or 0)
    header = panel(Text("OIC PARALLEL COLLECTOR | LIVE EXTRACTION MONITOR", justify="center",
                        style="bold bright_cyan"), "XMLA / MDX / EXCEL")
    overview = Text.assemble(("Run: ", "dim"), short(manifest.get("run_id"), 45),
        (" | ", "dim"), styled(status),
        f" | Elapsed {format_duration(elapsed(manifest.get('started_at'), manifest.get('finished_at')))}",
        f"\nQuery: {entry.get('query_id', '—')} | {config.get('code', '—')}/{config.get('cube_type', '—')}"
        f" | HTTP: {engine} | Parallel {count}")
    def captions(axis):
        values = config.get(axis, [])
        # Config contains stable MDX; don't guess human-readable captions.
        width = console.size.width // 2 if console.size.width >= 110 else console.size.width
        return short(", ".join(str(v) for v in values), max(22, width - 20))
    configuration = Text(f"X / ROWS: {captions('x')}\nY / COLUMNS: {captions('y')}\n"
        f"Values: {len(config.get('measures', []))} | X/Y chunk: "
        f"{config.get('x_chunk', '—')}/{config.get('y_chunk', '—')}")
    width = 18
    bar = "█" * int(coverage * width) + "░" * (width - int(coverage * width))
    progress = Text.assemble((f"{bar} {done}/{n} COMPLETED ({coverage * 100:.1f}%)", "green"),
        f" | Pending {waiting} | Failed/partial {failed} | Stopped {stopped}")
    metrics = Text(
        f"Chunks OK {totals.get('chunks_ok', 0):,} | Failed {totals.get('chunks_failed', 0):,}"
        f" | Cells {totals.get('cells', 0):,} | {bytes_mb:.2f} MiB\n"
        f"Active {active}/{count} (run {running}, retry {retrying}, pause {paused}) | Execute calls {calls:,} | Avg "
        + (f"{avg:.0f} ms" if avg is not None else "—")
        + f" | Rate {rate} | Cooldown {str(cool) + 's' if cool else 'NONE'}\n"
        f"Excel: {entry.get('export_status', 'PENDING')} | Rows {entry.get('rows', '—')} "
        "| Dynamic chunks: total unknown")
    worker_table = Table(box=rich_box.SIMPLE, expand=True, padding=(0, 1), title="05 / PARALLEL WORKERS")
    for name in ("ID", "Status", "Current X/Y", "OK", "Fail", "HTTP", "ms", "HB", "Last event"):
        worker_table.add_column(name, no_wrap=True, overflow="ellipsis")
    for w in workers:
        current_key = str(w.get("current_request") or "")
        match = re.search(r"_x(\d+)_y(\d+)$", current_key)
        xy = f"{int(match[1])}/{int(match[2])}" if match else "—"
        latency = w.get("last_request_ms")
        hb = w.get("heartbeat_age")
        worker_table.add_row(f"W{int(w.get('worker_id') or 0):02d}", styled(w.get("status", "STARTING")),
            xy, str(w.get("chunks_ok", 0)), str(w.get("chunks_failed", 0)),
            str(w.get("last_http") or "—"), f"{latency:.0f}" if latency is not None else "—",
            f"{hb}s" if hb is not None else "—", short(w.get("last_message"), 45))
    queue = Table(box=rich_box.SIMPLE, expand=True, padding=(0, 1), title="06 / QUERY QUEUE")
    for name in ("Query", "Class/Cube", "Status", "Workers", "Chunks", "Rows"):
        queue.add_column(name, overflow="ellipsis", no_wrap=True)
    # Keep current query visible on shorter terminals.
    wide = console.size.width >= 110
    row_limit = max(1, min(len(entries), console.size.height - (19 if wide else 28) - len(workers)))
    index = next((i for i, e in enumerate(entries) if e.get("query_id") == current), len(entries) - 1)
    first = max(0, min(index - row_limit // 2, len(entries) - row_limit))
    for e in entries[first:first + row_limit]:
        running = e.get("query_id") == current
        queue.add_row(str(e.get("query_id", "—")), f"{e.get('code', '—')}/{e.get('cube_type', '—')}",
                      styled(e.get("status", "PENDING")),
                      f"{totals.get('workers_terminal', 0)}/{count} ended" if running else str(e.get("workers", "—")),
                      str(totals.get("chunks_ok", 0) if running else e.get("chunks_ok", "—")),
                      str(e.get("rows", "—")))
    if len(entries) > row_limit:
        queue.caption = f"Showing {first + 1}–{first + row_limit} of {len(entries)} queries"
    recent = sorted(workers, key=lambda w: str(w.get("updated_at", "")), reverse=True)[:(3 if console.size.height >= 36 else 1)]
    events = Text("\n".join(f"{str(w.get('updated_at', ''))[11:19]} W{int(w.get('worker_id') or 0):02d} "
                             f"{short(w.get('last_message'), 110)}" for w in recent) or "Waiting for workers")
    if manifest.get("error") or entry.get("error"):
        events.append("\n" + short(manifest.get("error") or entry.get("error"), 160), style="red")
    info = (panel(overview, "01 / RUN OVERVIEW"), panel(configuration, "02 / QUERY CONFIGURATION"))
    stats = (panel(progress, "03 / BATCH PROGRESS"), panel(metrics, "04 / EXTRACTION METRICS"))
    if wide:
        top = Table.grid(expand=True)
        top.add_column(ratio=1)
        top.add_column(ratio=1)
        top.add_row(*info)
        top.add_row(*stats)
        sections = [top]
    else:
        sections = [*info, *stats]
    return Group(header, *sections, worker_table, queue, panel(events, "07 / LATEST ACTIVITY"),
                 Text("Refresh 1s | STOP / FORCE STOP: Main | Ctrl+C: close monitor", style="dim"))


def batch_monitor(batch_dir: Path) -> int:
    with Live(console=console, refresh_per_second=1, screen=True,
              auto_refresh=False, transient=False) as live:
        while True:
            manifest = read_json(batch_dir / "batch_manifest.json", {}) or {}
            entries = manifest.get("queries", [])
            current = manifest.get("current_query")
            if not current:
                last = next((e for e in reversed(entries) if e.get("started_at")), {})
                current = last.get("query_id")
            qdir = batch_dir / current if current else None
            config = read_json(qdir / "query_config.json", {}) if qdir else {}
            workers, totals = selected_summary(config, qdir) if config else ([], {})
            live.update(render_batch_monitor(manifest, config or {}, workers, totals,
                read_json(batch_dir / "cooldown.json", {}) or {}), refresh=True)
            if manifest.get("status") in ("COMPLETED", "PARTIAL", "FAILED", "STOPPED"):
                break
            time.sleep(1)
    return 0

def run_batch(batch_dir: Path) -> int:
    batch_dir = Path(batch_dir).resolve()
    batch = read_json(batch_dir / "batch_config.json", {}) or {}
    entries = [dict(entry, status=entry.get("status", "PENDING"))
               for entry in batch.get("queries", [])]
    if not entries or batch.get("run_id") != batch_dir.name:
        raise RuntimeError("Invalid or missing batch configuration")
    output_root = Path(batch["output_root"]).resolve()
    earlier = read_json(batch_dir / "batch_manifest.json", {}) or {}
    earlier_entries = {item["query_id"]: item for item in earlier.get("queries", [])
                       if isinstance(item, dict) and item.get("query_id")}
    stop = batch_dir / "stop.flag"
    force = batch_dir / "force_stop.flag"
    for entry in entries:
        previous = earlier_entries.get(entry["query_id"], {})
        if previous.get("status") == "COMPLETED" and previous.get("excel_path") and Path(previous["excel_path"]).is_file():
            entry.update(previous)
    _batch_snapshot(batch_dir, status="STARTING", current=None, entries=entries)
    python = str(Path(sys.executable).resolve())
    script = str(Path(__file__).resolve())
    window_id = "OIC-Execution-" + uuid.uuid4().hex
    monitor_command = [python, script, "--batch-monitor", str(batch_dir)]
    monitor_open = False
    try:
        for entry in entries:
            qid = entry["query_id"]
            qdir = batch_dir / qid
            cfg_path = qdir / "query_config.json"
            config, _ = selected_paths(cfg_path)
            expected = read_json(qdir / "task_plan.json", {})
            if expected.get("query_hash") != query_hash(config):
                entry.update(status="FAILED", error="Apply-time plan does not match configuration")
                _batch_snapshot(batch_dir, status="RUNNING", current=qid, entries=entries)
                continue
            previous = earlier_entries.get(qid, {})
            if (previous.get("status") == "COMPLETED"
                    and previous.get("excel_path")
                    and Path(previous["excel_path"]).is_file()):
                entry.update(previous)
                continue
            if stop.exists() or force.exists():
                break
            (qdir / "stop.flag").unlink(missing_ok=True)
            entry.update(status="RUNNING", error="", started_at=utc_now())
            _batch_snapshot(batch_dir, status="RUNNING", current=qid, entries=entries)
            attempt_id = uuid.uuid4().hex
            launch_error = None
            commands, titles = prepare_worker_tabs(config, qdir, cfg_path, attempt_id)
            try:
                if not monitor_open:
                    launch_terminal_window([monitor_command, *commands],
                                           title="OIC | Execution", cwd=HERE,
                                           tab_titles=["LIVE MONITOR", *titles], keep_open=False,
                                           maximized=True, window_id=window_id)
                    monitor_open = True
                else:
                    launch_terminal_window(commands, title="OIC | Execution", cwd=HERE,
                                           tab_titles=titles, keep_open=False, window_id=window_id)
            except Exception as exc:
                launch_error = str(exc)
                for wid in range(1, int(config["workers"]) + 1):
                    wd = qdir / f"worker_{wid:02d}"
                    (wd / f"cancel-{attempt_id}.flag").touch()
                    state_path = wd / "state.json"
                    state = read_json(state_path, {}) or {}
                    state.update(status="FAILED", last_message="Terminal launch failed: " + launch_error[:200],
                                 updated_at=utc_now())
                    write_json(state_path, state)
            if launch_error is None:
                def mark_stopping():
                    entry["status"] = "STOPPING"
                    _batch_snapshot(batch_dir, status="STOPPING", current=qid, entries=entries)
                wait_for_query_workers(config, qdir, attempt_id, control_dir=batch_dir,
                                       on_stopping=mark_stopping)
            else:
                entry["error"] = "WORKER LAUNCH: " + launch_error[:250]

            workers, totals = selected_summary(config, qdir)
            entry["chunks_ok"] = totals["chunks_ok"]
            entry["finished_at"] = utc_now()
            if stop.exists() or force.exists():
                entry["status"] = "STOPPED"
            elif launch_error:
                entry["status"] = "PARTIAL" if totals["chunks_ok"] else "FAILED"
            elif totals["workers_completed"] == config["workers"]:
                entry["status"] = "COMPLETED"
            else:
                entry["status"] = "PARTIAL" if totals["chunks_ok"] else "FAILED"
            try:
                extraction_status = entry["status"]
                entry.update(status="EXPORTING", export_status="EXPORTING")
                _batch_snapshot(batch_dir, status="RUNNING", current=qid, entries=entries)
                output, row_count = selected_export(config, qdir)
                entry.update(status=extraction_status, excel_path=output, rows=row_count,
                             export_status="COMPLETED")
            except Exception as exc:
                entry.update(status="PARTIAL" if totals["chunks_ok"] else "FAILED",
                             export_status="FAILED", error=f"EXCEL: {exc}")
            entry["finished_at"] = utc_now()
            _update_history(output_root, dict(
                run_id=batch_dir.name, query_id=qid, code=config["code"],
                cube_type=config.get("cube_type", "FY"), cube=config["cube"],
                status=entry["status"], started_at=entry["started_at"],
                finished_at=entry["finished_at"], rows=entry.get("rows", 0),
                excel_path=entry.get("excel_path", ""), error=entry.get("error", "")))
            _batch_snapshot(batch_dir, status="RUNNING", current=qid, entries=entries)
        final_status = ("STOPPED" if stop.exists() or force.exists()
                        else "COMPLETED" if all(e["status"] == "COMPLETED" for e in entries)
                        else "PARTIAL")
        _batch_snapshot(batch_dir, status=final_status, current=None, entries=entries)
        return 0 if final_status == "COMPLETED" else 1
    except Exception as exc:
        _batch_snapshot(batch_dir, status="FAILED", current=None,
                        entries=entries, error=str(exc)[:400])
        raise


def selected_main(argv) -> int:
    parser = argparse.ArgumentParser(description="Run an explicit user-selected OIC XMLA query")
    parser.add_argument("--query-config", type=Path, required=True)
    parser.add_argument("--launch", action="store_true")
    parser.add_argument("--worker-id", type=int, default=0)
    parser.add_argument("--monitor-only", action="store_true")
    parser.add_argument("--export-only", action="store_true")
    parser.add_argument("--attempt-id")
    args = parser.parse_args(argv)
    config, run_dir = selected_paths(args.query_config)
    if args.worker_id:
        try:
            return selected_worker(args)
        except Exception as exc:
            # Record startup failures as well as failures outside the extraction loop.
            if args.attempt_id and 1 <= args.worker_id <= int(config["workers"]):
                path = run_dir / f"worker_{args.worker_id:02d}" / "state.json"
                state = read_json(path, {}) or {}
                if (state.get("attempt_id") == args.attempt_id
                        and state.get("pid") in (None, os.getpid())
                        and not (path.parent / f"cancel-{args.attempt_id}.flag").exists()):
                    state.update(status="FAILED", pid=os.getpid(),
                                 process_created=_current_process_creation(),
                                 exit_code=1, last_message=str(exc)[:200], updated_at=utc_now())
                    write_json(path, state)
            log(str(exc), args.worker_id, "ERROR")
            return 1
    if args.monitor_only:
        return selected_monitor(config, run_dir, args.attempt_id)
    if args.export_only:
        path, rows = selected_export(config, run_dir)
        log(f"EXPORTED {rows:,} rows to {path}", level="OK")
        return 0
    if args.launch:
        return selected_launch(config, run_dir, args.query_config.resolve())
    parser.error("Specify --launch, --worker-id, --monitor-only or --export-only")
    return 2


# ============================================================
# MAIN
# ============================================================

def main() -> int:

    if "--batch-run" in sys.argv or "--batch-monitor" in sys.argv:
        operation = sys.argv[1] if len(sys.argv) > 1 else ""
        if len(sys.argv) != 3 or operation not in ("--batch-run", "--batch-monitor"):
            raise ValueError("Expected --batch-run/--batch-monitor directory")
        directory = Path(sys.argv[2]).resolve()
        return run_batch(directory) if operation == "--batch-run" else batch_monitor(directory)
    if "--query-config" in sys.argv:
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
                "Opening PowerShell collection window"
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