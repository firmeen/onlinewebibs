#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC PARALLEL COLLECTOR
======================

XMLA / MDX / OLAP Aggregate Collector
Rich Live Monitor + Windows Terminal Tabs
4 Parallel Workers + Raw XML + Excel + Resume

INSTALL
-------
python -m pip install requests openpyxl rich

COMMANDS
--------
Prepare job plan:
    python oic_parallel.py --prepare

Launch 4 worker tabs and Rich monitor:
    python oic_parallel.py

Test one hierarchy:
    python oic_parallel.py --worker-id 1 --max-jobs 1 --max-years 1

Run one worker:
    python oic_parallel.py --worker-id 1

Check status:
    python oic_parallel.py --status

Merge manifests:
    python oic_parallel.py --merge

Re-export Excel:
    python oic_parallel.py --export-only

Resume:
    python oic_parallel.py

Graceful shutdown:
    Ctrl+C in monitor:
        Request graceful stop for ALL workers.

    Ctrl+C again:
        Exit monitor.

    Ctrl+C inside a worker tab:
        Stop only that worker.

Existing successful XML responses are reused.

FEATURES
--------
- 4 workers in 4 tabs of ONE new Windows Terminal window
- Original PowerShell displays Rich Live Monitor
- Isolated Worker manifests and output directories
- Atomic JSON writes with unique temporary filenames
- Windows file access retry
- Worker process locks
- Worker heartbeat and crash detection
- HTTP 429 / 500 / 502 / 503 / 504 recovery
- HTTP 401 / 403 JWT refresh
- Connection and read timeout handling
- Exponential backoff with jitter
- Shared cooldown and adaptive concurrency
- Retry history
- Graceful shutdown
- Raw XML preservation
- Excel export
- Resume

DATA MODEL
----------
COLUMNS = YEAR only
ROWS    = One hierarchy x Measures

Default:
    1 YEAR x 20 Members per request

EXCEL
-----
One Excel workbook per hierarchy.

NOTE
----
This collector extracts OLAP Aggregates.
It does not extract individual transactions.
Independent hierarchy results are not a combined
multidimensional transaction table.

Only the verified MV cube profile is configured.
"""

from __future__ import annotations

import argparse
import base64
import ctypes
import json
import math
import os
import random
import re
import shutil
import subprocess
import sys
import threading
import time
import uuid

from collections import OrderedDict
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import urlparse
from xml.etree import ElementTree as ET

import requests

from rich.align import Align
from rich.console import Console, Group
from rich.live import Live
from rich.panel import Panel
from rich.table import Table
from rich.text import Text
from rich import box as rbox

try:
    from openpyxl import Workbook
except ImportError:
    Workbook = None


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

APP_NAME = "OIC PARALLEL COLLECTOR"
APP_VERSION = "4.0.0"

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
# VERIFIED MV CUBE PROFILE
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

DEFAULT_RETRIES = 3

DEFAULT_MAX_RESPONSE_MB = 100
DEFAULT_MAX_REQUESTS_PER_JOB = 5000

MONITOR_REFRESH_SECONDS = 1.0

EXCEL_MAX_ROWS = 1_048_576

RETRYABLE_HTTP = {
    408,
    425,
    429,
    500,
    502,
    503,
    504,
}

USER_AGENT = (
    "Mozilla/5.0 "
    "(Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 "
    "(KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)

console = Console(highlight=False)


# ============================================================
# EXCEPTIONS
# ============================================================

class StopRequested(Exception):
    pass


class CollectionError(Exception):

    def __init__(
        self,
        message: str,
        *,
        retryable: bool = False,
        http_status: int | None = None,
    ):

        super().__init__(message)

        self.retryable = retryable
        self.http_status = http_status
        self.history = []


# ============================================================
# BASIC UTILITIES
# ============================================================

def utc_now() -> str:

    return datetime.now(
        timezone.utc
    ).isoformat(
        timespec="seconds"
    )


def local_time() -> str:

    return datetime.now().strftime(
        "%H:%M:%S"
    )


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

    styles = {
        "INFO": "cyan",
        "OK": "green",
        "WARN": "yellow",
        "ERROR": "red",
    }

    color = styles.get(level, "white")

    console.print(
        f"[dim]{local_time()}[/dim] "
        f"[bold]{prefix}[/bold] "
        f"[{color}]{message}[/]"
    )


def atomic_json(
    path: Path,
    data: Any,
    *,
    attempts: int = 14,
) -> None:
    """
    Atomic JSON writer.

    Uses a unique temporary filename for each write.

    Addresses Windows sharing violations:
        WinError 5
        WinError 32
        WinError 33

    Does not share a fixed state.json.tmp filename.
    """

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary = path.with_name(
        f"{path.name}."
        f"{os.getpid()}."
        f"{uuid.uuid4().hex}.tmp"
    )

    try:

        with temporary.open(
            "w",
            encoding="utf-8",
        ) as stream:

            json.dump(
                data,
                stream,
                ensure_ascii=False,
                indent=2,
            )

            stream.write("\n")
            stream.flush()

            os.fsync(
                stream.fileno()
            )

        for attempt in range(attempts):

            try:

                os.replace(
                    temporary,
                    path,
                )

                return

            except OSError as exc:

                winerror = getattr(
                    exc,
                    "winerror",
                    None,
                )

                retryable = (
                    winerror in (5, 32, 33)
                    or isinstance(
                        exc,
                        PermissionError,
                    )
                )

                if not retryable:
                    raise

                if attempt == attempts - 1:
                    raise

                delay = min(
                    0.05 * (attempt + 1),
                    0.45,
                )

                delay += random.uniform(
                    0,
                    0.04,
                )

                time.sleep(delay)

    finally:

        try:
            temporary.unlink(
                missing_ok=True
            )

        except OSError:
            pass


def read_json(
    path: Path,
    default: Any = None,
) -> Any:

    for attempt in range(5):

        try:

            with path.open(
                "r",
                encoding="utf-8",
            ) as stream:

                return json.load(stream)

        except FileNotFoundError:
            return default

        except (
            PermissionError,
            OSError,
            ValueError,
            UnicodeError,
        ):

            if attempt == 4:
                return default

            time.sleep(
                0.03 * (attempt + 1)
            )

    return default


def read_env() -> dict[str, str]:

    values = dict(os.environ)

    path = HERE / ".env"

    if not path.exists():
        return values

    for line in path.read_text(
        encoding="utf-8-sig"
    ).splitlines():

        line = line.strip()

        if line.startswith("export "):
            line = line[7:].strip()

        if (
            not line
            or line.startswith("#")
            or "=" not in line
        ):
            continue

        key, value = line.split(
            "=",
            1,
        )

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

        token = token.removeprefix(
            "Bearer "
        ).strip()

        payload = token.split(".")[1]

        payload += "=" * (
            -len(payload) % 4
        )

        decoded = base64.urlsafe_b64decode(
            payload
        )

        data = json.loads(
            decoded.decode("utf-8")
        )

        return int(
            data.get("exp", 0)
        )

    except (
        ValueError,
        IndexError,
        TypeError,
        KeyError,
    ):
        return 0


def pid_alive(pid: int | None) -> bool:

    try:

        pid = int(pid)

        if pid <= 0:
            return False

        if pid == os.getpid():
            return True

        if os.name == "nt":

            kernel = ctypes.windll.kernel32

            kernel.OpenProcess.argtypes = [
                ctypes.c_ulong,
                ctypes.c_int,
                ctypes.c_ulong,
            ]

            kernel.OpenProcess.restype = (
                ctypes.c_void_p
            )

            kernel.GetExitCodeProcess.argtypes = [
                ctypes.c_void_p,
                ctypes.POINTER(
                    ctypes.c_ulong
                ),
            ]

            kernel.CloseHandle.argtypes = [
                ctypes.c_void_p
            ]

            handle = kernel.OpenProcess(
                0x1000,
                0,
                pid,
            )

            if not handle:
                return ctypes.get_last_error() == 5

            try:

                code = ctypes.c_ulong()

                success = (
                    kernel.GetExitCodeProcess(
                        handle,
                        ctypes.byref(code),
                    )
                )

                return (
                    bool(success)
                    and code.value == 259
                )

            finally:

                kernel.CloseHandle(
                    handle
                )

        os.kill(pid, 0)

        return True

    except ProcessLookupError:
        return False

    except PermissionError:
        return True

    except Exception:
        return False


def exclusive_lock(
    path: Path,
    *,
    timeout: float = 0,
    stale_after: float = 30,
) -> None:

    started = time.monotonic()

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    while True:

        try:

            descriptor = os.open(
                str(path),
                os.O_CREAT
                | os.O_EXCL
                | os.O_WRONLY,
                0o600,
            )

            try:

                payload = {
                    "pid": os.getpid(),
                    "created": time.time(),
                }

                os.write(
                    descriptor,
                    json.dumps(
                        payload
                    ).encode("utf-8"),
                )

            finally:

                os.close(descriptor)

            return

        except FileExistsError:

            owner = read_json(
                path,
                {},
            )

            owner_pid = (
                owner.get("pid")
                if isinstance(owner, dict)
                else None
            )

            try:

                age = (
                    time.time()
                    - path.stat().st_mtime
                )

            except OSError:
                continue

            if (
                age > stale_after
                and (
                    not owner_pid
                    or not pid_alive(owner_pid)
                )
            ):

                try:
                    path.unlink()

                except OSError:
                    pass

                continue

            if (
                time.monotonic() - started
                >= timeout
            ):

                raise CollectionError(
                    f"Lock is held: {path} "
                    f"(PID {owner_pid or 'unknown'})"
                )

            time.sleep(
                0.12
                + random.uniform(0, 0.08)
            )


def unlock(path: Path) -> None:

    owner = read_json(
        path,
        {},
    )

    if (
        isinstance(owner, dict)
        and owner.get("pid") == os.getpid()
    ):

        try:

            path.unlink(
                missing_ok=True
            )

        except OSError:
            pass


def validate_hierarchy(
    name: str,
) -> str:

    name = name.strip()

    pattern = (
        r"\[[^\]\r\n]+\]"
        r"\."
        r"\[[^\]\r\n]+\]"
    )

    if not re.fullmatch(
        pattern,
        name,
    ):

        raise ValueError(
            f"Invalid hierarchy: {name}"
        )

    return name


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
    job: dict,
) -> Path:

    return (
        worker_dir(
            output,
            job["worker_id"],
        )
        / f"dim_{job['id']:03d}"
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


# ============================================================
# DISCOVER HIERARCHIES
# ============================================================

def discover_hierarchies(
    source: Path,
) -> list[str]:

    data = read_json(
        source,
        {},
    )

    candidates = []

    if isinstance(data, dict):

        candidates.extend(
            data.get(
                "discovered_hierarchies"
            ) or []
        )

        for request in data.get(
            "requests",
            [],
        ):

            if not isinstance(
                request,
                dict,
            ):
                continue

            result = (
                request.get("result")
                or {}
            )

            candidates.extend(
                result.get(
                    "discovered_hierarchies"
                ) or []
            )

    if not candidates:

        for path in sorted(
            source.parent.glob(
                "*.response.xml"
            )
        ):

            try:

                root = ET.parse(
                    path
                ).getroot()

                for element in root.iter():

                    tag = element.tag.rsplit(
                        "}",
                        1,
                    )[-1]

                    if tag == "HierarchyInfo":

                        candidates.append(
                            element.get(
                                "name",
                                "",
                            )
                        )

            except (
                ET.ParseError,
                OSError,
            ):
                continue

            if candidates:
                break

    found = []
    seen = set()

    for name in candidates:

        if not isinstance(
            name,
            str,
        ):
            continue

        name = name.strip()

        if (
            not name
            or name in seen
            or name == CUBE["year"]
            or name == "[Measures]"
        ):
            continue

        try:

            validate_hierarchy(
                name
            )

        except ValueError:
            continue

        found.append(name)
        seen.add(name)

    return found


# ============================================================
# JOB PLAN
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

    path = (
        output
        / "job_plan.json"
    )

    existing = read_json(
        path
    )

    if isinstance(
        existing,
        dict,
    ):

        if existing.get("workers") != workers:

            raise CollectionError(
                "Existing job_plan.json uses "
                f"{existing.get('workers')} workers."
            )

        return existing

    hierarchies = discover_hierarchies(
        source
    )

    if not hierarchies:

        raise CollectionError(
            "No discovered hierarchies found: "
            f"{source}"
        )

    jobs = []

    for index, hierarchy in enumerate(
        hierarchies,
        start=1,
    ):

        worker_id = (
            (index - 1) % workers
        ) + 1

        jobs.append({
            "id": index,
            "worker_id": worker_id,
            "hierarchy": hierarchy,
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
        "partition_method": "round_robin",
        "source_manifest": str(source),
        "jobs": jobs,
    }

    atomic_json(
        path,
        plan,
    )

    return plan


def verify_chunk_configuration(
    output: Path,
    plan: dict,
    args: argparse.Namespace,
) -> None:

    expected = (
        args.years_per_request,
        args.members_per_request,
    )

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        manifest = load_worker_manifest(
            output,
            worker_id,
        )

        for chunk in manifest[
            "chunks"
        ].values():

            actual = (
                chunk.get(
                    "year_page_size"
                ),
                chunk.get(
                    "member_page_size"
                ),
            )

            if (
                all(
                    value is not None
                    for value in actual
                )
                and actual != expected
            ):

                raise CollectionError(
                    "Existing chunks use paging sizes "
                    f"{actual}; requested {expected}. "
                    "Use the original chunk sizes "
                    "or a different --output-dir."
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

    year_set = (
        f"Subset("
        f"{year}.DefaultMember.Children, "
        f"{year_offset}, "
        f"{years_per_request}"
        f")"
    )

    member_set = (
        f"Subset("
        f"{dimension}.AllMembers, "
        f"{member_offset}, "
        f"{members_per_request}"
        f")"
    )

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
        "CELL PROPERTIES "
        "VALUE, FORMAT_STRING, LANGUAGE"
    )


# ============================================================
# SOAP BUILDER
# ============================================================

def build_soap(
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
        "year_count": 0,
        "member_count": 0,
        "row_tuple_count": 0,
        "nonempty_cells": 0,
        "error": None,
    }

    try:

        root = ET.parse(
            path
        ).getroot()

    except (
        ET.ParseError,
        OSError,
    ) as exc:

        result["error"] = str(exc)

        return result

    for element in root.iter():

        tag = element.tag.rsplit(
            "}",
            1,
        )[-1]

        if tag == "Fault":

            message = " ".join(
                text.strip()
                for text in element.itertext()
                if text.strip()
            )

            result["error"] = (
                "SOAP Fault: "
                + message[:500]
            )

            return result

    dataset = root.find(
        f".//{{{MD_NS}}}root"
    )

    if dataset is None:

        result["error"] = (
            "Missing multidimensional XMLA dataset"
        )

        return result

    axes = {}

    for axis in dataset.findall(
        f"./{{{MD_NS}}}Axes/"
        f"{{{MD_NS}}}Axis"
    ):

        tuples = []

        for item in axis.findall(
            f"{{{MD_NS}}}Tuples/"
            f"{{{MD_NS}}}Tuple"
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
                        ) or ""
                    ),
                    "caption": (
                        member.findtext(
                            f"{{{MD_NS}}}Caption"
                        ) or ""
                    ),
                })

            tuples.append(members)

        axes[
            axis.get("name", "")
        ] = tuples

    if (
        "Axis0" not in axes
        or "Axis1" not in axes
    ):

        result["error"] = (
            "Missing Axis0 YEAR "
            "or Axis1 DIMENSION + MEASURES"
        )

        return result

    years = []

    for item in axes["Axis0"]:

        year = next(
            (
                member
                for member in item
                if member["hierarchy"]
                == CUBE["year"]
            ),
            None,
        )

        if year is None:

            result["error"] = (
                "YEAR hierarchy not found on Axis0"
            )

            return result

        years.append(year)

    rows = []
    unique_members = set()

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

        if (
            member is None
            or measure is None
        ):

            result["error"] = (
                "Dimension or Measure "
                "not found on Axis1"
            )

            return result

        unique_members.add(
            member["unique_name"]
        )

        rows.append(
            (member, measure)
        )

    cells = {}

    for cell in dataset.findall(
        f"./{{{MD_NS}}}CellData/"
        f"{{{MD_NS}}}Cell"
    ):

        value = cell.find(
            f"{{{MD_NS}}}Value"
        )

        if value is None:
            continue

        try:

            ordinal = int(
                cell.get(
                    "CellOrdinal",
                    "",
                )
            )

        except ValueError:
            continue

        cells[ordinal] = value.text

    result.update({
        "valid": True,
        "year_count": len(years),
        "member_count": len(
            unique_members
        ),
        "row_tuple_count": len(rows),
        "nonempty_cells": len(cells),
    })

    if include_data:

        result["years"] = years
        result["rows"] = rows
        result["cells"] = cells

    return result


# ============================================================
# EXCEL UTILITIES
# ============================================================

def numeric_value(
    raw: str | None,
) -> Any:

    if (
        raw is None
        or not str(raw).strip()
    ):
        return None

    try:

        number = float(raw)

        if not math.isfinite(number):
            return str(raw)

        if (
            number.is_integer()
            and abs(number) < 1e15
        ):
            return int(number)

        return number

    except (
        ValueError,
        OverflowError,
    ):
        return str(raw)


def excel_text(
    value: str,
) -> str:

    value = str(
        value or ""
    )

    if value.startswith(
        ("=", "+", "-", "@")
    ):
        return "'" + value

    return value


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

        raise CollectionError(
            f"Invalid saved XML: "
            f"{response_file.name}: "
            f"{result['error']}"
        )

    years = result["years"]
    rows = result["rows"]
    cells = result["cells"]

    if not years:
        return

    grouped = OrderedDict()

    for row_index, (
        member,
        measure,
    ) in enumerate(rows):

        measure_name = measure[
            "unique_name"
        ]

        prefix = "[Measures].["

        if (
            measure_name.startswith(prefix)
            and measure_name.endswith("]")
        ):

            measure_name = measure_name[
                len(prefix):-1
            ]

        for year_index, year in enumerate(
            years
        ):

            ordinal = (
                year_index
                + len(years) * row_index
            )

            key = (
                year["unique_name"],
                member["unique_name"],
            )

            if key not in grouped:

                grouped[key] = (
                    year,
                    member,
                    {},
                )

            if ordinal in cells:

                grouped[key][2][
                    measure_name
                ] = numeric_value(
                    cells[ordinal]
                )

    for (
        year,
        member,
        measures,
    ) in grouped.values():

        if not measures:
            continue

        yield [
            excel_text(
                year["caption"]
            ),
            excel_text(
                year["unique_name"]
            ),
            excel_text(
                dimension
            ),
            excel_text(
                member["caption"]
            ),
            excel_text(
                member["unique_name"]
            ),
            *[
                measures.get(name)
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

        raise CollectionError(
            "openpyxl missing. "
            "Run: python -m pip install openpyxl"
        )

    successful = sorted(
        (
            chunk
            for chunk in chunks
            if chunk.get("status") == "OK"
        ),
        key=lambda chunk: (
            chunk["year_offset"],
            chunk["member_offset"],
        ),
    )

    if not successful:
        return None, 0

    excel_dir = (
        output
        / "excel"
    )

    excel_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    destination = (
        excel_dir
        / f"hierarchy_{job['id']:03d}.xlsx"
    )

    temporary = destination.with_name(
        f"{destination.stem}."
        f"{uuid.uuid4().hex}.tmp.xlsx"
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
    row_count = 0
    sheet_rows = 1

    sheet = workbook.create_sheet(
        "Data_1"
    )

    sheet.append(headers)

    try:

        for chunk in successful:

            response_file = (
                dimension_dir(
                    output,
                    job,
                )
                / chunk.get(
                    "response_file",
                    "",
                )
            )

            if not response_file.is_file():

                raise CollectionError(
                    "Missing saved XML chunk: "
                    f"{response_file.name}"
                )

            for row in iter_excel_rows(
                response_file,
                job["hierarchy"],
            ):

                if sheet_rows >= EXCEL_MAX_ROWS:

                    sheet_index += 1

                    sheet = workbook.create_sheet(
                        f"Data_{sheet_index}"
                    )

                    sheet.append(headers)

                    sheet_rows = 1

                sheet.append(row)

                row_count += 1
                sheet_rows += 1

        info = workbook.create_sheet(
            "INFO"
        )

        info_rows = [
            ("JOB_ID", job["id"]),
            ("WORKER_ID", job["worker_id"]),
            ("STATUS", status),
            ("HIERARCHY", job["hierarchy"]),
            ("YEAR_COLUMNS", CUBE["year"]),
            ("CUBE", CUBE["cube"]),
            ("ROWS", row_count),
            (
                "SUCCESSFUL_CHUNKS",
                len(successful),
            ),
            (
                "DATA_TYPE",
                "OLAP AGGREGATE",
            ),
            (
                "NOTICE",
                "Not raw transactions",
            ),
            (
                "NOTICE",
                "Do not sum All members "
                "with their detail members",
            ),
            (
                "WARNING",
                (
                    "INCOMPLETE"
                    if status != "COMPLETED"
                    else ""
                ),
            ),
        ]

        for key, value in info_rows:
            info.append([key, value])

        workbook.save(
            temporary
        )

        for attempt in range(12):

            try:

                os.replace(
                    temporary,
                    destination,
                )

                break

            except PermissionError:

                if attempt == 11:

                    raise CollectionError(
                        "Cannot replace open Excel file: "
                        f"{destination.name}"
                    )

                time.sleep(
                    0.15 * (attempt + 1)
                )

    finally:

        try:

            temporary.unlink(
                missing_ok=True
            )

        except OSError:
            pass

    return destination, row_count


# ============================================================
# MANIFEST
# ============================================================

def worker_manifest_path(
    output: Path,
    worker_id: int,
) -> Path:

    return (
        worker_dir(
            output,
            worker_id,
        )
        / "manifest.json"
    )


def load_worker_manifest(
    output: Path,
    worker_id: int,
) -> dict:

    data = read_json(
        worker_manifest_path(
            output,
            worker_id,
        )
    )

    if not isinstance(
        data,
        dict,
    ):

        data = {
            "created_at": utc_now(),
            "worker_id": worker_id,
            "jobs": {},
            "chunks": {},
        }

    data.setdefault(
        "jobs",
        {},
    )

    data.setdefault(
        "chunks",
        {},
    )

    return data


def save_worker_manifest(
    output: Path,
    worker_id: int,
    manifest: dict,
) -> None:

    manifest["updated_at"] = utc_now()

    atomic_json(
        worker_manifest_path(
            output,
            worker_id,
        ),
        manifest,
    )


def job_chunks(
    manifest: dict,
    job_id: int,
) -> list[dict]:

    return [
        chunk
        for chunk in manifest[
            "chunks"
        ].values()
        if chunk.get("job_id") == job_id
    ]


# ============================================================
# WORKER REPORTER
# ============================================================

class WorkerReporter:
    """
    Reports worker progress to state.json.

    Uses:
        - Unique temporary filenames
        - Atomic writes
        - Thread synchronization
        - Background heartbeat

    Monitoring failures do not interrupt XMLA extraction.
    """

    def __init__(
        self,
        output: Path,
        worker_id: int,
        run_id: str,
        assigned_jobs: int,
        manifest: dict,
    ):

        self.worker_id = worker_id

        self.path = (
            worker_dir(
                output,
                worker_id,
            )
            / "state.json"
        )

        self.mutex = threading.RLock()

        self.finished = threading.Event()

        self.last_write_error = 0.0

        self.state = {
            "run_id": run_id,
            "pid": os.getpid(),
            "worker_id": worker_id,
            "status": "RUNNING",
            "started_at": utc_now(),
            "updated_at": utc_now(),
            "assigned_jobs": assigned_jobs,
            "current_job": None,
            "current_dimension": None,
            "current_chunk": None,
            "request_started_at": None,
            "last_message": "Worker started",
            "recent": [],
        }

        self.sync(manifest)

        self.thread = threading.Thread(
            target=self._heartbeat,
            daemon=True,
        )

        self.thread.start()

    def _save(self) -> None:

        self.state[
            "updated_at"
        ] = utc_now()

        try:

            atomic_json(
                self.path,
                self.state,
            )

        except OSError as exc:

            elapsed = (
                time.monotonic()
                - self.last_write_error
            )

            if elapsed > 30:

                self.last_write_error = (
                    time.monotonic()
                )

                log(
                    "State file temporarily unavailable: "
                    f"{exc}",
                    self.worker_id,
                    "WARN",
                )

    def _heartbeat(self) -> None:

        while not self.finished.wait(2.0):

            with self.mutex:
                self._save()

    def update(
        self,
        **kwargs,
    ) -> None:

        with self.mutex:

            self.state.update(
                kwargs
            )

            self._save()

    def event(
        self,
        message: str,
        level: str = "INFO",
        status: str | None = None,
    ) -> None:

        with self.mutex:

            self.state[
                "last_message"
            ] = message

            if status is not None:
                self.state["status"] = status

            recent = self.state[
                "recent"
            ]

            recent.append({
                "time": local_time(),
                "level": level,
                "message": message,
            })

            self.state["recent"] = (
                recent[-8:]
            )

            self._save()

        log(
            message,
            self.worker_id,
            level,
        )

    def sync(
        self,
        manifest: dict,
    ) -> None:

        jobs = list(
            manifest["jobs"].values()
        )

        chunks = list(
            manifest["chunks"].values()
        )

        self.update(
            completed_jobs=sum(
                job.get("status") == "COMPLETED"
                for job in jobs
            ),
            failed_jobs=sum(
                job.get("status") == "FAILED"
                for job in jobs
            ),
            partial_jobs=sum(
                job.get("status")
                in ("PARTIAL", "INTERRUPTED")
                for job in jobs
            ),
            running_jobs=sum(
                job.get("status") == "RUNNING"
                for job in jobs
            ),
            chunks_ok=sum(
                chunk.get("status") == "OK"
                for chunk in chunks
            ),
            chunks_failed=sum(
                chunk.get("status") == "FAILED"
                for chunk in chunks
            ),
            excel_rows=sum(
                int(
                    job.get("excel_rows")
                    or 0
                )
                for job in jobs
            ),
        )

    def close(
        self,
        status: str,
        message: str,
    ) -> None:

        self.event(
            message,
            (
                "OK"
                if status == "COMPLETED"
                else "WARN"
            ),
            status=status,
        )

        self.update(
            current_job=None,
            current_dimension=None,
            current_chunk=None,
            request_started_at=None,
            finished_at=utc_now(),
        )

        self.finished.set()

        self.thread.join(
            timeout=3
        )


# ============================================================
# WORKER LOCK
# ============================================================

class WorkerLock:

    def __init__(
        self,
        output: Path,
        worker_id: int,
    ):

        self.path = (
            worker_dir(
                output,
                worker_id,
            )
            / "worker.lock"
        )

    def __enter__(self):

        exclusive_lock(
            self.path,
            timeout=0,
            stale_after=10,
        )

        return self

    def __exit__(
        self,
        *_,
    ):

        unlock(
            self.path
        )


# ============================================================
# SHARED WORKER CONTROL
# ============================================================

class SharedControl:
    """
    Shared control for all worker processes.

    Provides:
        - Graceful stop signal
        - Shared cooldown
        - Adaptive concurrency
        - Request slot coordination
        - Basic circuit-breaker behavior

    Uses JSON files and exclusive file locks.
    No Redis, database or external queue required.
    """

    def __init__(
        self,
        output: Path,
        run_id: str,
        workers: int,
    ):

        self.path = (
            output
            / "shared_control.json"
        )

        self.lockfile = (
            output
            / "shared_control.lock"
        )

        self.run_id = run_id
        self.workers = workers

    def initial(self) -> dict:

        return {
            "run_id": self.run_id,
            "started_at": utc_now(),
            "cooldown_until": 0.0,
            "max_parallel": self.workers,
            "workers": self.workers,
            "slots": {},
            "pressure": 0,
            "healthy_streak": 0,
            "stop_requested": False,
            "last_error": "",
        }

    def _locked(self):

        parent = self

        class Guard:

            def __enter__(self):

                exclusive_lock(
                    parent.lockfile,
                    timeout=15,
                    stale_after=30,
                )

                return self

            def __exit__(
                self,
                *_,
            ):

                unlock(
                    parent.lockfile
                )

        return Guard()

    def reset(self) -> None:

        with self._locked():

            atomic_json(
                self.path,
                self.initial(),
            )

    def ensure(self) -> None:

        with self._locked():

            data = read_json(
                self.path
            )

            if (
                not isinstance(data, dict)
                or data.get("run_id")
                != self.run_id
            ):

                atomic_json(
                    self.path,
                    self.initial(),
                )

    def read(self) -> dict:

        data = read_json(
            self.path,
            {},
        )

        if (
            isinstance(data, dict)
            and data.get("run_id")
            == self.run_id
        ):
            return data

        return {}

    def mutate(
        self,
        callback,
    ) -> Any:

        with self._locked():

            data = read_json(
                self.path,
                {},
            )

            if (
                not isinstance(data, dict)
                or data.get("run_id")
                != self.run_id
            ):

                raise CollectionError(
                    "Shared control run ID mismatch"
                )

            result = callback(data)

            atomic_json(
                self.path,
                data,
            )

            return result

    def is_stopping(self) -> bool:

        return bool(
            self.read().get(
                "stop_requested",
                False,
            )
        )

    def request_stop(self) -> None:

        self.mutate(
            lambda data: data.update(
                stop_requested=True
            )
        )

    def wait(
        self,
        seconds: float,
    ) -> None:

        deadline = (
            time.monotonic()
            + max(0, seconds)
        )

        while time.monotonic() < deadline:

            if self.is_stopping():

                raise StopRequested(
                    "Graceful stop requested"
                )

            time.sleep(
                min(
                    0.5,
                    max(
                        0,
                        deadline - time.monotonic(),
                    ),
                )
            )

    def acquire(
        self,
        worker_id: int,
        reporter: WorkerReporter,
    ) -> None:

        label = str(worker_id)
        previous_status = ""

        while True:

            if self.is_stopping():

                raise StopRequested(
                    "Graceful stop requested"
                )

            def attempt(data):

                slots = data.setdefault(
                    "slots",
                    {},
                )

                for key, item in list(
                    slots.items()
                ):

                    if not pid_alive(
                        item.get("pid", 0)
                    ):

                        slots.pop(
                            key,
                            None,
                        )

                remaining = max(
                    0.0,
                    float(
                        data.get(
                            "cooldown_until",
                            0,
                        )
                    ) - time.time(),
                )

                if remaining > 0:

                    return (
                        "COOLDOWN",
                        remaining,
                    )

                limit = max(
                    1,
                    int(
                        data.get(
                            "max_parallel",
                            self.workers,
                        )
                    ),
                )

                if (
                    len(slots) >= limit
                    and label not in slots
                ):

                    return (
                        "WAITING",
                        1.0,
                    )

                slots[label] = {
                    "pid": os.getpid(),
                    "since": time.time(),
                }

                return (
                    "ACQUIRED",
                    0.0,
                )

            status, seconds = self.mutate(
                attempt
            )

            if status == "ACQUIRED":

                if previous_status:

                    reporter.event(
                        "Request slot acquired",
                        status="RUNNING",
                    )

                return

            if previous_status != status:

                reporter.event(
                    f"{status}: "
                    f"{int(math.ceil(seconds))}s",
                    "WARN",
                    status=status,
                )

                previous_status = status

            self.wait(
                min(
                    max(seconds, 0.3),
                    1.0,
                )
            )

    def release(
        self,
        worker_id: int,
    ) -> None:

        try:

            self.mutate(
                lambda data: (
                    data.setdefault(
                        "slots",
                        {},
                    ).pop(
                        str(worker_id),
                        None,
                    )
                )
            )

        except (
            OSError,
            CollectionError,
        ) as exc:

            log(
                f"Shared slot release: {exc}",
                worker_id,
                "WARN",
            )

    def pressure(
        self,
        status: int | None,
        delay: float,
    ) -> None:

        def change(data):

            data["pressure"] = (
                int(
                    data.get(
                        "pressure",
                        0,
                    )
                ) + 1
            )

            data[
                "healthy_streak"
            ] = 0

            data["last_error"] = (
                f"HTTP {status}"
                if status
                else "Network timeout"
            )

            if status in (429, 503):

                data[
                    "cooldown_until"
                ] = max(
                    float(
                        data.get(
                            "cooldown_until",
                            0,
                        )
                    ),
                    time.time() + delay,
                )

            if data["pressure"] >= 2:

                data[
                    "max_parallel"
                ] = max(
                    1,
                    int(
                        data.get(
                            "max_parallel",
                            self.workers,
                        )
                    ) - 1,
                )

                data["pressure"] = 0

        self.mutate(
            change
        )

    def healthy(self) -> None:

        def change(data):

            data["pressure"] = 0

            data[
                "healthy_streak"
            ] = (
                int(
                    data.get(
                        "healthy_streak",
                        0,
                    )
                ) + 1
            )

            enough_success = (
                data["healthy_streak"] >= 12
            )

            cooldown_finished = (
                time.time()
                >= float(
                    data.get(
                        "cooldown_until",
                        0,
                    )
                )
            )

            if (
                enough_success
                and cooldown_finished
            ):

                data[
                    "max_parallel"
                ] = min(
                    self.workers,
                    int(
                        data.get(
                            "max_parallel",
                            self.workers,
                        )
                    ) + 1,
                )

                data[
                    "healthy_streak"
                ] = 0

        self.mutate(
            change
        )


# ============================================================
# HTTP RETRY-AFTER PARSER
# ============================================================

def parse_retry_after(
    header: str | None,
) -> float | None:

    if not header:
        return None

    try:

        return max(
            0.0,
            min(
                float(header.strip()),
                3600.0,
            ),
        )

    except (
        ValueError,
        TypeError,
    ):

        try:

            deadline = parsedate_to_datetime(
                header
            )

            remaining = (
                deadline
                - datetime.now(
                    timezone.utc
                )
            ).total_seconds()

            return max(
                0.0,
                min(
                    remaining,
                    3600.0,
                ),
            )

        except (
            ValueError,
            TypeError,
            OverflowError,
        ):

            return None


# ============================================================
# AUTHENTICATED OIC HTTP CLIENT
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
            max_response_mb
            * 1024
            * 1024
        )

        self.session = requests.Session()

        self.session.headers.update({
            "User-Agent": USER_AGENT,
            "Accept-Language": (
                "th-TH,th;q=0.9"
            ),
        })

        hostname = urlparse(
            BASE_URL
        ).hostname

        cookie = cookie.removeprefix(
            "Cookie:"
        ).strip()

        for segment in cookie.split(";"):

            name, separator, value = (
                segment.strip().partition("=")
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

            raise CollectionError(
                "GenToken HTTP "
                f"{response.status_code}"
            )

        try:

            token = response.json().get(
                "d",
                "",
            )

        except ValueError as exc:

            raise CollectionError(
                "GenToken returned invalid JSON"
            ) from exc

        if token == "SESSION_EXPIRED":

            raise CollectionError(
                "Cookie session expired. "
                "Update OIC_COOKIE in .env."
            )

        if (
            not isinstance(token, str)
            or jwt_exp(token)
            <= time.time() + 30
        ):

            raise CollectionError(
                "GenToken returned unusable JWT"
            )

        self.token = token.removeprefix(
            "Bearer "
        ).strip()

    def execute_once(
        self,
        request_body: bytes,
        destination: Path,
    ) -> dict:

        if (
            jwt_exp(self.token)
            <= time.time() + 60
        ):

            self.refresh_token()

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

            total_bytes = 0

            temporary = destination.with_name(
                f"{destination.name}."
                f"{uuid.uuid4().hex}.partial"
            )

            try:

                with temporary.open(
                    "wb"
                ) as stream:

                    for chunk in response.iter_content(
                        chunk_size=256 * 1024
                    ):

                        if not chunk:
                            continue

                        total_bytes += len(chunk)

                        if (
                            total_bytes
                            > self.max_bytes
                        ):

                            raise CollectionError(
                                "XMLA response exceeds "
                                "--max-response-mb"
                            )

                        stream.write(chunk)

                os.replace(
                    temporary,
                    destination,
                )

            finally:

                temporary.unlink(
                    missing_ok=True
                )

            return {
                "http_status": response.status_code,
                "response_bytes": total_bytes,
                "retry_after": parse_retry_after(
                    response.headers.get(
                        "Retry-After"
                    )
                ),
                "content_type": response.headers.get(
                    "Content-Type",
                    "",
                ),
            }

    def close(self) -> None:

        self.session.close()


# ============================================================
# HTTP RECOVERY
# ============================================================

def execute_resilient(
    client: OICClient,
    control: SharedControl,
    reporter: WorkerReporter,
    request_body: bytes,
    directory: Path,
    basename: str,
    dimension: str,
    args: argparse.Namespace,
) -> tuple[dict, dict, list[dict]]:

    destination = (
        directory
        / f"{basename}.response.xml"
    )

    history = []

    max_attempts = args.retries + 1

    auth_refreshed = False

    attempt = 0

    while attempt < max_attempts:

        if control.is_stopping():

            raise StopRequested(
                "Graceful stop requested"
            )

        attempt += 1

        attempt_file = (
            directory
            / (
                f"{basename}."
                f"attempt_{attempt}.response.xml"
            )
        )

        reporter.update(
            current_chunk=(
                f"{basename} "
                f"attempt {attempt}"
            )
        )

        try:

            control.acquire(
                client.worker_id,
                reporter,
            )

            reporter.update(
                request_started_at=utc_now(),
                status="RUNNING",
            )

            try:

                response = client.execute_once(
                    request_body,
                    attempt_file,
                )

            finally:

                control.release(
                    client.worker_id
                )

            status_code = response[
                "http_status"
            ]

            if status_code == 200:

                parsed = parse_xmla(
                    attempt_file,
                    dimension,
                )

                if parsed["valid"]:

                    os.replace(
                        attempt_file,
                        destination,
                    )

                    control.healthy()

                    reporter.update(
                        status="RUNNING",
                        request_started_at=None,
                    )

                    return (
                        response,
                        parsed,
                        history,
                    )

                error = CollectionError(
                    parsed["error"]
                    or "Invalid XMLA response",
                    retryable=False,
                    http_status=status_code,
                )

            else:

                preview = ""

                if attempt_file.exists():

                    preview = (
                        attempt_file.read_bytes()[:180]
                        .decode(
                            "utf-8",
                            errors="replace",
                        )
                    )

                error = CollectionError(
                    f"HTTP {status_code}: {preview}",
                    retryable=(
                        status_code
                        in RETRYABLE_HTTP
                    ),
                    http_status=status_code,
                )

            history.append({
                "attempt": attempt,
                "http_status": status_code,
                "error": str(error),
                "time": utc_now(),
                "retryable": error.retryable,
            })

            if (
                status_code in (401, 403)
                and not auth_refreshed
            ):

                auth_refreshed = True

                reporter.event(
                    f"HTTP {status_code}: "
                    "refreshing JWT once",
                    "WARN",
                    status="RETRYING",
                )

                client.refresh_token()

                max_attempts += 1

                continue

            if status_code in (429, 503):

                delay = response[
                    "retry_after"
                ]

                if delay is None:

                    delay = (
                        [3, 7, 15][
                            min(attempt - 1, 2)
                        ]
                        + random.uniform(0, 1)
                    )

                control.pressure(
                    status_code,
                    delay,
                )

            elif status_code in (
                408,
                425,
                500,
                502,
                504,
            ):

                control.pressure(
                    status_code,
                    0,
                )

        except StopRequested:
            raise

        except requests.RequestException as exc:

            can_retry = (
                isinstance(
                    exc,
                    (
                        requests.Timeout,
                        requests.ConnectionError,
                        requests.exceptions.ChunkedEncodingError,
                    ),
                )
                and not isinstance(
                    exc,
                    requests.exceptions.SSLError,
                )
            )

            error = CollectionError(
                f"{type(exc).__name__}: {exc}",
                retryable=can_retry,
            )

            history.append({
                "attempt": attempt,
                "http_status": None,
                "error": str(error),
                "time": utc_now(),
                "retryable": can_retry,
            })

        finally:

            reporter.update(
                request_started_at=None
            )

        if (
            not error.retryable
            or attempt >= max_attempts
        ):

            final_error = CollectionError(
                str(error),
                retryable=False,
                http_status=error.http_status,
            )

            final_error.history = list(
                history
            )

            raise final_error

        delay = (
            [3, 7, 15][
                min(attempt - 1, 2)
            ]
            + random.uniform(0, 1)
        )

        reporter.event(
            f"RETRYING {attempt}/{args.retries} | "
            f"{error} | "
            f"backoff {delay:.1f}s",
            "WARN",
            status="RETRYING",
        )

        control.wait(
            delay
        )

    raise CollectionError(
        "Retry budget exhausted"
    )


# ============================================================
# EXTRACTION JOB
# ============================================================

def run_job(
    job: dict,
    args: argparse.Namespace,
    client: OICClient,
    control: SharedControl,
    reporter: WorkerReporter,
    manifest: dict,
    output: Path,
) -> str:

    worker_id = job["worker_id"]
    job_id = job["id"]
    dimension = job["hierarchy"]

    directory = dimension_dir(
        output,
        job,
    )

    directory.mkdir(
        parents=True,
        exist_ok=True,
    )

    reporter.update(
        current_job=job_id,
        current_dimension=dimension,
        status="RUNNING",
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
    requests_sent = 0

    status = "COMPLETED"
    last_error = None
    exhausted = False

    try:

        while not exhausted:

            if control.is_stopping():

                raise StopRequested(
                    "Stop requested"
                )

            if (
                args.max_years
                and year_offset >= args.max_years
            ):

                status = "PARTIAL"
                break

            member_offset = 0
            current_years = None

            while True:

                if control.is_stopping():

                    raise StopRequested(
                        "Stop requested"
                    )

                if (
                    requests_sent
                    >= args.max_requests_per_job
                ):

                    status = "PARTIAL"
                    exhausted = True
                    break

                basename = chunk_name(
                    year_offset,
                    member_offset,
                )

                key = chunk_key(
                    job_id,
                    year_offset,
                    member_offset,
                )

                response_file = (
                    directory
                    / f"{basename}.response.xml"
                )

                request_file = (
                    directory
                    / f"{basename}.request.xml"
                )

                existing = manifest[
                    "chunks"
                ].get(
                    key,
                    {},
                )

                result = None

                # ============================================
                # RECOVER PREVIOUSLY SUCCESSFUL RAW XML
                # ============================================

                paging_matches = (
                    existing.get(
                        "year_page_size",
                        args.years_per_request,
                    )
                    == args.years_per_request
                    and existing.get(
                        "member_page_size",
                        args.members_per_request,
                    )
                    == args.members_per_request
                )

                if (
                    response_file.is_file()
                    and paging_matches
                ):

                    parsed = parse_xmla(
                        response_file,
                        dimension,
                    )

                    valid_page = (
                        parsed["valid"]
                        and parsed["year_count"]
                        <= args.years_per_request
                        and parsed["member_count"]
                        <= args.members_per_request
                    )

                    if valid_page:

                        result = parsed

                        manifest["chunks"][key] = {
                            **existing,
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
                            "response_file": (
                                response_file.name
                            ),
                            "request_file": (
                                request_file.name
                            ),
                            "status": "OK",
                            "result": parsed,
                            "recovered": (
                                existing.get("status")
                                != "OK"
                            ),
                        }

                        save_worker_manifest(
                            output,
                            worker_id,
                            manifest,
                        )

                        reporter.event(
                            f"[{job_id}] RESUME "
                            f"YEAR={year_offset} "
                            f"MEMBER={member_offset}",
                            "OK",
                        )

                # ============================================
                # NEW REQUEST
                # ============================================

                if result is None:

                    mdx = build_mdx(
                        dimension,
                        year_offset,
                        member_offset,
                        args.years_per_request,
                        args.members_per_request,
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
                        "request_file": (
                            request_file.name
                        ),
                        "response_file": (
                            response_file.name
                        ),
                        "started_at": utc_now(),
                        "status": "RUNNING",
                        "retry_history": [],
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

                        (
                            http,
                            result,
                            history,
                        ) = execute_resilient(
                            client,
                            control,
                            reporter,
                            soap,
                            directory,
                            basename,
                            dimension,
                            args,
                        )

                        record.update({
                            "http_status": (
                                http["http_status"]
                            ),
                            "response_bytes": (
                                http["response_bytes"]
                            ),
                            "result": result,
                            "retry_history": history,
                            "status": "OK",
                        })

                        reporter.event(
                            f"[{job_id}] OK | "
                            f"{result['year_count']} years | "
                            f"{result['member_count']} members | "
                            f"{result['nonempty_cells']} cells | "
                            f"{http['response_bytes']:,} bytes",
                            "OK",
                            status="RUNNING",
                        )

                    except StopRequested:
                        raise

                    except (
                        CollectionError,
                        requests.RequestException,
                        OSError,
                    ) as exc:

                        previous_success = any(
                            chunk.get("status") == "OK"
                            for chunk in job_chunks(
                                manifest,
                                job_id,
                            )
                        )

                        status = (
                            "PARTIAL"
                            if previous_success
                            else "FAILED"
                        )

                        last_error = str(exc)

                        record.update({
                            "status": "FAILED",
                            "error": last_error,
                            "retry_history": getattr(
                                exc,
                                "history",
                                [],
                            ),
                            "http_status": getattr(
                                exc,
                                "http_status",
                                None,
                            ),
                        })

                        reporter.event(
                            f"[{job_id}] FAILED: "
                            f"{last_error}",
                            "ERROR",
                            status="RUNNING",
                        )

                        exhausted = True

                    finally:

                        record[
                            "completed_at"
                        ] = utc_now()

                        save_worker_manifest(
                            output,
                            worker_id,
                            manifest,
                        )

                        reporter.sync(
                            manifest
                        )

                    requests_sent += 1

                    if exhausted:
                        break

                # ============================================
                # PAGINATION
                # ============================================

                if current_years is None:

                    current_years = result[
                        "year_count"
                    ]

                if result["year_count"] == 0:

                    exhausted = True
                    break

                if (
                    result["member_count"]
                    < args.members_per_request
                ):
                    break

                member_offset += (
                    args.members_per_request
                )

                control.wait(
                    args.delay
                )

            if exhausted:
                break

            if (
                current_years is None
                or current_years
                < args.years_per_request
            ):
                break

            year_offset += (
                args.years_per_request
            )

            control.wait(
                args.delay
            )

    except StopRequested:

        status = "INTERRUPTED"
        last_error = "Graceful stop requested"

    except KeyboardInterrupt:

        status = "INTERRUPTED"
        last_error = "Worker interrupted"

    # ========================================================
    # SAVE JOB STATUS
    # ========================================================

    job_record = {
        "hierarchy": dimension,
        "status": status,
        "updated_at": utc_now(),
        "error": last_error,
        "request_count_this_run": requests_sent,
    }

    manifest["jobs"][
        str(job_id)
    ] = job_record

    save_worker_manifest(
        output,
        worker_id,
        manifest,
    )

    # ========================================================
    # EXPORT SUCCESSFUL DATA
    # ========================================================

    try:

        chunks = job_chunks(
            manifest,
            job_id,
        )

        if chunks:

            reporter.event(
                f"[{job_id}] EXPORTING EXCEL"
            )

            excel_path, rows = export_hierarchy(
                output,
                job,
                chunks,
                status,
            )

            if excel_path:

                job_record.update({
                    "excel_file": str(
                        excel_path
                    ),
                    "excel_rows": rows,
                })

                reporter.event(
                    f"[{job_id}] EXCEL "
                    f"{rows:,} rows | "
                    f"{excel_path.name}",
                    "OK",
                )

    except Exception as exc:

        job_record[
            "excel_error"
        ] = str(exc)

        if status == "COMPLETED":

            status = "PARTIAL"

            job_record[
                "status"
            ] = status

        reporter.event(
            f"[{job_id}] EXCEL ERROR: {exc}",
            "ERROR",
        )

    save_worker_manifest(
        output,
        worker_id,
        manifest,
    )

    reporter.sync(
        manifest
    )

    reporter.event(
        f"[{job_id}] {status}",
        (
            "OK"
            if status == "COMPLETED"
            else "WARN"
        ),
        status="RUNNING",
    )

    return status


# ============================================================
# RUN WORKER
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

        raise CollectionError(
            "Worker ID must be "
            f"1-{plan['workers']}"
        )

    assigned = [
        job
        for job in plan["jobs"]
        if job["worker_id"] == worker_id
    ]

    with WorkerLock(
        output,
        worker_id,
    ):

        manifest = load_worker_manifest(
            output,
            worker_id,
        )

        control_data = read_json(
            output / "shared_control.json",
            {},
        )

        run_id = (
            args.run_id
            or (
                control_data.get("run_id")
                if isinstance(
                    control_data,
                    dict,
                )
                else None
            )
            or (
                "manual-"
                + uuid.uuid4().hex[:8]
            )
        )

        control = SharedControl(
            output,
            run_id,
            plan["workers"],
        )

        control.ensure()

        reporter = WorkerReporter(
            output,
            worker_id,
            run_id,
            len(assigned),
            manifest,
        )

        client = None

        failures = 0
        processed = 0
        stopped = False

        console.print(
            Panel(
                Align.center(
                    "[bold cyan]"
                    "OIC PARALLEL COLLECTOR"
                    "[/]\n"
                    f"WORKER {worker_id:02d} / "
                    f"{plan['workers']:02d}\n"
                    f"{len(assigned)} "
                    "assigned hierarchies"
                ),
                border_style="cyan",
            )
        )

        try:

            env = read_env()

            cookie = env.get(
                "OIC_COOKIE",
                "",
            ).strip()

            token = env.get(
                "OIC_BEARER_TOKEN",
                "",
            ).strip()

            if not cookie:

                raise CollectionError(
                    "OIC_COOKIE missing from .env"
                )

            if Workbook is None:

                raise CollectionError(
                    "Install openpyxl: "
                    "python -m pip install openpyxl"
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

                if control.is_stopping():

                    stopped = True
                    break

                previous = manifest[
                    "jobs"
                ].get(
                    str(job["id"]),
                    {},
                )

                if (
                    previous.get("status")
                    == "COMPLETED"
                    and not args.force
                ):

                    reporter.event(
                        f"[{job['id']}] "
                        "SKIP COMPLETED",
                        "OK",
                    )

                    continue

                if (
                    args.max_jobs
                    and processed
                    >= args.max_jobs
                ):
                    break

                processed += 1

                status = run_job(
                    job,
                    args,
                    client,
                    control,
                    reporter,
                    manifest,
                    output,
                )

                if status == "INTERRUPTED":

                    stopped = True
                    break

                if status != "COMPLETED":
                    failures += 1

                control.wait(
                    args.delay
                )

            reporter.sync(
                manifest
            )

            if stopped:

                final_status = "STOPPED"

            elif failures:

                final_status = (
                    "COMPLETED_WITH_ERRORS"
                )

            else:

                final_status = "COMPLETED"

            reporter.close(
                final_status,
                (
                    f"Worker {final_status} | "
                    f"completed="
                    f"{reporter.state.get('completed_jobs', 0)} "
                    f"partial="
                    f"{reporter.state.get('partial_jobs', 0)} "
                    f"failed="
                    f"{reporter.state.get('failed_jobs', 0)}"
                ),
            )

            if stopped:
                return 130

            return (
                1
                if failures
                else 0
            )

        except (
            StopRequested,
            KeyboardInterrupt,
        ):

            reporter.close(
                "STOPPED",
                "Worker interrupted; "
                "saved XML can resume",
            )

            return 130

        except Exception as exc:

            reporter.event(
                f"WORKER ERROR: {exc}",
                "ERROR",
            )

            reporter.close(
                "ERROR",
                f"Worker crashed: {exc}",
            )

            return 1

        finally:

            if client is not None:
                client.close()

            try:

                save_worker_manifest(
                    output,
                    worker_id,
                    manifest,
                )

            except OSError as exc:

                log(
                    "Manifest save failed: "
                    f"{exc}",
                    worker_id,
                    "ERROR",
                )

            reporter.finished.set()


# ============================================================
# RICH MONITOR STATE
# ============================================================

def read_worker_states(
    output: Path,
    plan: dict,
    run_id: str | None = None,
) -> list[dict]:

    terminal_statuses = {
        "COMPLETED",
        "COMPLETED_WITH_ERRORS",
        "STOPPED",
        "ERROR",
        "CRASHED",
        "LOCKED",
    }

    states = []

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        path = (
            worker_dir(
                output,
                worker_id,
            )
            / "state.json"
        )

        state = read_json(
            path,
            {},
        )

        valid_state = (
            isinstance(state, dict)
            and (
                run_id is None
                or state.get("run_id") == run_id
            )
        )

        if not valid_state:

            state = {
                "worker_id": worker_id,
                "status": "STARTING",
                "assigned_jobs": sum(
                    job["worker_id"] == worker_id
                    for job in plan["jobs"]
                ),
            }

        else:

            status = state.get(
                "status",
                "STARTING",
            )

            if (
                status not in terminal_statuses
                and not pid_alive(
                    state.get("pid")
                )
            ):

                state["status"] = "CRASHED"

                state["last_message"] = (
                    "Process PID "
                    f"{state.get('pid')} "
                    "is no longer running"
                )

            elif status not in terminal_statuses:

                try:

                    last_update = (
                        datetime.fromisoformat(
                            state.get(
                                "updated_at",
                                "",
                            ).replace(
                                "Z",
                                "+00:00",
                            )
                        )
                    )

                    age = (
                        datetime.now(
                            timezone.utc
                        ) - last_update
                    ).total_seconds()

                    if age > 20:

                        state["status"] = "STALE"

                except ValueError:
                    pass

        states.append(
            state
        )

    return states


# ============================================================
# RICH MONITOR DASHBOARD
# ============================================================

def render_monitor(
    output: Path,
    plan: dict,
    run_id: str | None,
    started_at: float,
):

    states = read_worker_states(
        output,
        plan,
        run_id,
    )

    def count(key: str) -> int:

        return sum(
            int(
                state.get(key)
                or 0
            )
            for state in states
        )

    total = len(
        plan["jobs"]
    )

    completed = count(
        "completed_jobs"
    )

    partial = count(
        "partial_jobs"
    )

    failed = count(
        "failed_jobs"
    )

    finished = min(
        total,
        completed + partial + failed,
    )

    remaining = max(
        0,
        total - finished,
    )

    duration = int(
        time.monotonic()
        - started_at
    )

    elapsed = (
        f"{duration // 3600:02d}:"
        f"{duration % 3600 // 60:02d}:"
        f"{duration % 60:02d}"
    )

    # ========================================================
    # HEADER
    # ========================================================

    header = Panel(
        Align.center(
            "[bold bright_cyan]"
            "OIC PARALLEL COLLECTOR"
            "[/]\n"
            "[bold]"
            "LIVE EXTRACTION MONITOR"
            "[/]\n"
            "[dim]"
            "YEAR COLUMNS | "
            "DIMENSION + MEASURES ROWS | "
            "RAW XML + EXCEL"
            "[/]"
        ),
        border_style="cyan",
    )

    # ========================================================
    # OVERVIEW
    # ========================================================

    metrics = Table(
        box=rbox.SIMPLE,
        show_header=False,
        expand=True,
    )

    metrics.add_column(
        ratio=1
    )

    metrics.add_column(
        justify="right",
        ratio=1,
    )

    metrics.add_column(
        ratio=1
    )

    metrics.add_column(
        justify="right",
        ratio=1,
    )

    metrics.add_row(
        "TOTAL JOBS",
        str(total),
        "COMPLETED",
        f"[green]{completed}[/]",
    )

    metrics.add_row(
        "REMAINING",
        str(remaining),
        "PARTIAL / FAILED",
        (
            f"[yellow]{partial}[/]"
            " / "
            f"[red]{failed}[/]"
        ),
    )

    metrics.add_row(
        "CHUNKS OK",
        f"[green]{count('chunks_ok'):,}[/]",
        "CHUNKS FAILED",
        f"[red]{count('chunks_failed'):,}[/]",
    )

    metrics.add_row(
        "EXCEL ROWS",
        f"{count('excel_rows'):,}",
        "ELAPSED",
        elapsed,
    )

    # ========================================================
    # SHARED CONTROL
    # ========================================================

    shared = read_json(
        output / "shared_control.json",
        {},
    )

    if (
        not isinstance(shared, dict)
        or shared.get("run_id") != run_id
    ):
        shared = {}

    cooldown = max(
        0,
        int(
            float(
                shared.get(
                    "cooldown_until",
                    0,
                )
            ) - time.time()
        ),
    )

    metrics.add_row(
        "PARALLEL LIMIT",
        str(
            shared.get(
                "max_parallel",
                plan["workers"],
            )
        ),
        "SHARED COOLDOWN",
        (
            f"[yellow]{cooldown}s[/]"
            if cooldown
            else "—"
        ),
    )

    # ========================================================
    # TOTAL PROGRESS
    # ========================================================

    bar_width = 44

    filled = min(
        bar_width,
        round(
            bar_width
            * finished
            / max(total, 1)
        ),
    )

    progress = Text(
        "━" * filled,
        style="cyan",
    )

    progress.append(
        "─" * (bar_width - filled),
        style="grey35",
    )

    progress.append(
        f"  {finished / max(total, 1):.1%}",
        style="bold",
    )

    # ========================================================
    # WORKER TABLE
    # ========================================================

    table = Table(
        box=rbox.ROUNDED,
        expand=True,
        header_style="bold cyan",
    )

    columns = [
        "WORKER",
        "STATUS",
        "JOBS",
        "CHUNKS OK / FAIL",
        "CURRENT JOB",
        "WAIT",
    ]

    for column in columns:

        table.add_column(
            column,
            overflow="ellipsis",
            no_wrap=(
                column != "CURRENT JOB"
            ),
        )

    colors = {
        "RUNNING": "cyan",
        "RETRYING": "yellow",
        "COOLDOWN": "yellow",
        "WAITING": "dim",
        "COMPLETED": "green",
        "COMPLETED_WITH_ERRORS": "yellow",
        "STOPPED": "yellow",
        "CRASHED": "red",
        "ERROR": "red",
        "STALE": "red",
        "STARTING": "dim",
    }

    for state in states:

        status = state.get(
            "status",
            "STARTING",
        )

        current_job = state.get(
            "current_job"
        )

        if current_job:

            current = (
                f"[{current_job}] "
                f"{state.get('current_dimension') or ''}"
            )

        else:
            current = "—"

        request_started = state.get(
            "request_started_at"
        )

        wait = "—"

        if request_started:

            try:

                started = datetime.fromisoformat(
                    request_started
                )

                seconds = int(
                    (
                        datetime.now(
                            timezone.utc
                        ) - started
                    ).total_seconds()
                )

                wait = f"{seconds}s"

            except ValueError:
                pass

        worker_finished = (
            int(
                state.get(
                    "completed_jobs"
                ) or 0
            )
            + int(
                state.get(
                    "partial_jobs"
                ) or 0
            )
            + int(
                state.get(
                    "failed_jobs"
                ) or 0
            )
        )

        table.add_row(
            f"W{state['worker_id']:02d}",
            (
                f"[{colors.get(status, 'white')}]"
                f"{status}[/]"
            ),
            (
                f"{worker_finished}/"
                f"{state.get('assigned_jobs', '?')}"
            ),
            (
                f"{state.get('chunks_ok', 0)}"
                " / "
                f"{state.get('chunks_failed', 0)}"
            ),
            current,
            wait,
        )

    # ========================================================
    # ACTIVITY
    # ========================================================

    activity = Table(
        box=rbox.SIMPLE,
        expand=True,
        header_style="bold cyan",
    )

    activity.add_column(
        "WORKER",
        width=9,
    )

    activity.add_column(
        "LATEST ACTIVITY"
    )

    activity.add_column(
        "TIME",
        width=9,
    )

    for state in states:

        recent = (
            state.get("recent")
            or []
        )

        activity.add_row(
            f"W{state['worker_id']:02d}",
            str(
                state.get(
                    "last_message"
                ) or "Starting..."
            ),
            (
                recent[-1].get(
                    "time",
                    "",
                )
                if recent
                else ""
            ),
        )

    # ========================================================
    # FOOTER
    # ========================================================

    footer = (
        "[dim]"
        "CTRL+C = REQUEST GRACEFUL STOP ALL WORKERS"
        " | "
        "CTRL+C AGAIN = EXIT MONITOR"
        "[/]"
    )

    if shared.get(
        "stop_requested"
    ):

        footer = (
            "[yellow]"
            "STOP REQUESTED — "
            "waiting for in-flight requests "
            "and exports to finish"
            "[/]"
        )

    return Group(
        header,
        Panel(
            metrics,
            title="OVERVIEW",
        ),
        Panel(
            Align.center(progress),
            title="TOTAL PROGRESS",
        ),
        Panel(
            table,
            title="PARALLEL WORKERS",
        ),
        Panel(
            activity,
            title="LATEST ACTIVITY",
        ),
        Panel(
            Align.center(footer),
            border_style="grey35",
        ),
    )


# ============================================================
# MONITOR LOOP
# ============================================================

def monitor_workers(
    output: Path,
    plan: dict,
    run_id: str,
) -> int:

    control = SharedControl(
        output,
        run_id,
        plan["workers"],
    )

    started_at = time.monotonic()

    stop_sent = False

    terminal_states = {
        "COMPLETED",
        "COMPLETED_WITH_ERRORS",
        "STOPPED",
        "ERROR",
        "CRASHED",
        "LOCKED",
    }

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
    ) as live:

        while True:

            try:

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

                states = read_worker_states(
                    output,
                    plan,
                    run_id,
                )

                if all(
                    state.get("status")
                    in terminal_states
                    for state in states
                ):
                    break

                if (
                    time.monotonic() - started_at > 45
                    and any(
                        state.get("status") == "STARTING"
                        for state in states
                    )
                ):

                    log(
                        "One or more worker tabs "
                        "have not started. "
                        "Check Windows Terminal.",
                        level="WARN",
                    )

                    break

                time.sleep(
                    MONITOR_REFRESH_SECONDS
                )

            except KeyboardInterrupt:

                if stop_sent:
                    return 130

                stop_sent = True

                control.request_stop()

                log(
                    "Graceful stop requested "
                    "for ALL workers. "
                    "Press Ctrl+C again "
                    "to leave monitor.",
                    level="WARN",
                )

    merge_manifests(
        output,
        plan,
        quiet=True,
    )

    return 0


# ============================================================
# WINDOWS TERMINAL LAUNCHER
# ============================================================

def powershell_quote(
    value: str,
) -> str:

    return (
        "'"
        + str(value).replace(
            "'",
            "''",
        )
        + "'"
    )


def launch_worker_tabs(
    args: argparse.Namespace,
    output: Path,
    plan: dict,
) -> int:

    if os.name != "nt":

        raise CollectionError(
            "Windows Terminal tabs require Windows"
        )

    wt = (
        shutil.which("wt.exe")
        or shutil.which("wt")
    )

    if not wt:

        candidate = (
            Path(
                os.environ.get(
                    "LOCALAPPDATA",
                    "",
                )
            )
            / "Microsoft"
            / "WindowsApps"
            / "wt.exe"
        )

        if candidate.is_file():
            wt = str(candidate)

    if not wt:

        raise CollectionError(
            "Windows Terminal (wt.exe) not found"
        )

    # ========================================================
    # REFUSE DUPLICATE LIVE WORKERS
    # ========================================================

    busy = []

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        lockfile = (
            worker_dir(
                output,
                worker_id,
            )
            / "worker.lock"
        )

        if not lockfile.exists():
            continue

        owner = read_json(
            lockfile,
            {},
        )

        if (
            isinstance(owner, dict)
            and pid_alive(
                owner.get("pid")
            )
        ):

            busy.append(
                f"W{worker_id:02d} "
                f"(PID {owner.get('pid')})"
            )

    if busy:

        raise CollectionError(
            "Existing live workers detected: "
            + ", ".join(busy)
            + ". Stop them before starting "
            "another parallel run."
        )

    # ========================================================
    # NEW RUN
    # ========================================================

    run_id = (
        datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )
        + "-"
        + uuid.uuid4().hex[:6]
    )

    SharedControl(
        output,
        run_id,
        plan["workers"],
    ).reset()

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

        python_executable = powershell_quote(
            Path(
                sys.executable
            ).resolve()
        )

        script_path = powershell_quote(
            Path(
                __file__
            ).resolve()
        )

        ps_command = (
            f"& {python_executable} "
            f"{script_path} "
            f"--worker-id {worker_id} "
            f"--workers {plan['workers']} "
            f"--source {powershell_quote(args.source)} "
            f"--output-dir {powershell_quote(output)} "
            f"--years-per-request "
            f"{args.years_per_request} "
            f"--members-per-request "
            f"{args.members_per_request} "
            f"--timeout {args.timeout} "
            f"--delay {args.delay} "
            f"--retries {args.retries} "
            f"--max-response-mb "
            f"{args.max_response_mb} "
            f"--max-requests-per-job "
            f"{args.max_requests_per_job} "
            f"--max-years {args.max_years} "
            f"--max-jobs {args.max_jobs} "
            f"--run-id {powershell_quote(run_id)}"
        )

        if args.force:
            ps_command += " --force"

        command.extend([
            "new-tab",
            "--title",
            f"OIC Worker {worker_id:02d}",
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

    return monitor_workers(
        output,
        plan,
        run_id,
    )


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
        "workers": plan["workers"],
        "total_jobs": len(plan["jobs"]),
        "completed": 0,
        "partial": 0,
        "failed": 0,
        "interrupted": 0,
        "running": 0,
        "pending": 0,
        "successful_chunks": 0,
        "failed_chunks": 0,
        "excel_rows": 0,
        "jobs": [],
    }

    manifests = {
        worker_id: load_worker_manifest(
            output,
            worker_id,
        )
        for worker_id in range(
            1,
            plan["workers"] + 1,
        )
    }

    for job in plan["jobs"]:

        manifest = manifests[
            job["worker_id"]
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

        status_key = status.lower()

        if status_key not in summary:
            status_key = "pending"

        summary[status_key] += 1

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
            record.get(
                "excel_rows"
            ) or 0
        )

        summary["jobs"].append({
            "id": job["id"],
            "worker_id": job["worker_id"],
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

    atomic_json(
        destination,
        summary,
    )

    if not quiet:

        log(
            f"TOTAL {summary['total_jobs']} | "
            f"COMPLETED {summary['completed']} | "
            f"PARTIAL {summary['partial']} | "
            f"FAILED {summary['failed']} | "
            f"INTERRUPTED {summary['interrupted']} | "
            f"PENDING {summary['pending']}"
        )

    return summary


# ============================================================
# RE-EXPORT EXISTING XML
# ============================================================

def export_existing(
    output: Path,
    plan: dict,
) -> int:

    failed = 0

    for job in plan["jobs"]:

        manifest = load_worker_manifest(
            output,
            job["worker_id"],
        )

        status = manifest[
            "jobs"
        ].get(
            str(job["id"]),
            {},
        ).get(
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

                log(
                    f"[{job['id']}] "
                    f"{path.name} | "
                    f"{rows:,} rows",
                    level="OK",
                )

        except Exception as exc:

            failed += 1

            log(
                f"[{job['id']}] "
                f"EXPORT ERROR: {exc}",
                level="ERROR",
            )

    return (
        1
        if failed
        else 0
    )


# ============================================================
# PRINT JOB PLAN
# ============================================================

def print_plan(
    plan: dict,
) -> None:

    table = Table(
        title="OIC JOB PLAN",
        box=rbox.ROUNDED,
    )

    table.add_column(
        "WORKER"
    )

    table.add_column(
        "HIERARCHIES",
        justify="right",
    )

    table.add_column(
        "JOB IDS"
    )

    for worker_id in range(
        1,
        plan["workers"] + 1,
    ):

        ids = [
            str(job["id"])
            for job in plan["jobs"]
            if job["worker_id"] == worker_id
        ]

        table.add_row(
            f"W{worker_id:02d}",
            str(len(ids)),
            ", ".join(ids),
        )

    console.print(table)

    log(
        f"TOTAL: {len(plan['jobs'])} hierarchies"
    )


# ============================================================
# COMMAND LINE
# ============================================================

def build_parser() -> argparse.ArgumentParser:

    parser = argparse.ArgumentParser(
        description=(
            "OIC Parallel XMLA Collector "
            "with Rich Live Monitor "
            "and Windows Terminal tabs"
        ),
    )

    parser.add_argument(
        "--prepare",
        action="store_true",
    )

    parser.add_argument(
        "--launch",
        action="store_true",
    )

    parser.add_argument(
        "--worker-id",
        type=int,
        default=None,
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=DEFAULT_WORKERS,
    )

    parser.add_argument(
        "--status",
        action="store_true",
    )

    parser.add_argument(
        "--merge",
        action="store_true",
    )

    parser.add_argument(
        "--export-only",
        action="store_true",
    )

    parser.add_argument(
        "--source",
        type=Path,
        default=DEFAULT_SOURCE,
    )

    parser.add_argument(
        "--output-dir",
        type=Path,
        default=DEFAULT_OUTPUT,
    )

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
        "--retries",
        type=int,
        default=DEFAULT_RETRIES,
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
    )

    parser.add_argument(
        "--max-jobs",
        type=int,
        default=0,
    )

    parser.add_argument(
        "--force",
        action="store_true",
    )

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

    invalid = (
        args.workers < 1
        or args.years_per_request < 1
        or args.members_per_request < 1
        or args.timeout < 1
        or args.max_response_mb < 1
        or args.max_requests_per_job < 1
        or args.retries < 0
        or args.delay < 0
        or args.max_years < 0
        or args.max_jobs < 0
    )

    if invalid:

        parser.error(
            "Invalid worker, chunk, timeout, "
            "retry or limit value"
        )

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

    operation_selected = any([
        args.prepare,
        args.launch,
        args.worker_id is not None,
        args.status,
        args.merge,
        args.export_only,
    ])

    if not operation_selected:
        args.launch = True

    try:

        plan = prepare_plan(
            output,
            args.source,
            args.workers,
        )

        verify_chunk_configuration(
            output,
            plan,
            args,
        )

        # ====================================================
        # PREPARE
        # ====================================================

        if args.prepare:

            print_plan(
                plan
            )

            log(
                "PLAN: "
                f"{output / 'job_plan.json'}"
            )

            return 0

        # ====================================================
        # STATUS
        # ====================================================

        if args.status:

            console.print(
                render_monitor(
                    output,
                    plan,
                    None,
                    time.monotonic(),
                )
            )

            merge_manifests(
                output,
                plan,
                quiet=True,
            )

            return 0

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
        # WORKER
        # ====================================================

        if args.worker_id is not None:

            return run_worker(
                args,
                output,
                plan,
            )

        # ====================================================
        # LAUNCH + RICH MONITOR
        # ====================================================

        return launch_worker_tabs(
            args,
            output,
            plan,
        )

    except (
        CollectionError,
        OSError,
        ValueError,
    ) as exc:

        log(
            f"ERROR: {exc}",
            level="ERROR",
        )

        return 1

    except KeyboardInterrupt:

        log(
            "Interrupted",
            level="WARN",
        )

        return 130


if __name__ == "__main__":
    sys.exit(main())