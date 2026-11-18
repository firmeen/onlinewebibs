#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC PARALLEL COLLECTOR
======================

XMLA / MDX / OLAP Aggregate Collector
Rich Terminal Monitor + Windows Terminal Tabs
4 Parallel Workers + Raw XML + Excel + Resume

INSTALL
-------
python -m pip install requests openpyxl rich

COMMANDS
--------
Prepare jobs:
    python oic_parallel_collector.py --prepare

Start 4 workers in ONE NEW Windows Terminal window:
    python oic_parallel_collector.py

Or:
    python oic_parallel_collector.py --launch

The original PowerShell displays a LIVE RICH MONITOR.
A new Windows Terminal window opens with 4 worker tabs.

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
Windows Terminal (wt.exe)
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

from collections import OrderedDict
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

    temporary = path.with_name(
        path.name + ".tmp"
    )

    temporary.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2,
        ) + "\n",
        encoding="utf-8",
    )

    temporary.replace(path)


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
    ):

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
            f"?Class={CUBE['class']}"
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
                f"?Class={CUBE['class']}"
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
# WINDOWS TERMINAL LAUNCHER
# ============================================================

def powershell_quote(value: str) -> str:

    return (
        "'"
        + value.replace("'", "''")
        + "'"
    )


def find_windows_terminal() -> str | None:

    path = (
        shutil.which("wt.exe")
        or shutil.which("wt")
    )

    if path:
        return path

    local_app_data = os.environ.get(
        "LOCALAPPDATA",
        "",
    )

    if local_app_data:

        candidate = (
            Path(local_app_data)
            / "Microsoft"
            / "WindowsApps"
            / "wt.exe"
        )

        if candidate.exists():
            return str(candidate)

    return None


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


def launch_worker_tabs(
    args: argparse.Namespace,
    output: Path,
    plan: dict,
) -> str:

    if os.name != "nt":

        raise RuntimeError(
            "Windows Terminal tabs require Windows."
        )

    wt = find_windows_terminal()

    if not wt:

        raise RuntimeError(
            "Windows Terminal (wt.exe) not found."
        )

    run_id = (
        datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + "_"
        + uuid.uuid4().hex[:6]
    )

    # ========================================================
    # ONE NEW WINDOWS TERMINAL WINDOW
    #
    # All workers are created by ONE wt.exe invocation.
    #
    # This prevents:
    #   - Four separate Windows Terminal windows
    #   - Multiple independent terminal launchers
    #
    # Command structure:
    #
    # wt -w new new-tab <worker1>
    #           ; new-tab <worker2>
    #           ; new-tab <worker3>
    #           ; new-tab <worker4>
    #
    # The original PowerShell remains the Rich monitor.
    # ========================================================

    command = [
        wt,
        "-w",
        "new",
    ]

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        if worker_id > 1:
            command.append(";")

        ps_command = worker_powershell_command(
            args,
            worker_id,
            output,
            run_id,
        )

        command.extend([
            "new-tab",
            "--title",
            f"OIC | Worker {worker_id:02d}",
            "--startingDirectory",
            str(HERE),
            "powershell.exe",
            "-NoLogo",
            "-NoExit",
            "-Command",
            ps_command,
        ])

    subprocess.Popen(
        command,
        cwd=str(HERE),
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
            "with Rich Live Monitor and "
            "Windows Terminal worker tabs"
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
# MAIN
# ============================================================

def main() -> int:

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
                "Opening ONE new Windows Terminal window"
            )

            log(
                f"Creating {plan['workers']} worker tabs"
            )

            run_id = launch_worker_tabs(
                args,
                output,
                plan,
            )

            log(
                "Worker tabs launched",
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