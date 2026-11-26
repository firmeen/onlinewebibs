#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC HIGH-PERFORMANCE LOAD ENGINE v2.0.0
=======================================

INSTALL
-------
python -m pip install curl_cffi rich openpyxl

COMMANDS
--------
python oic_load_engine.py
python oic_load_engine.py --workers 4 --mode benchmark
python oic_load_engine.py --workers 4 --mode run

FEATURES
--------
- Standalone Python file
- Existing .env authentication
- Verified MV/FY field mappings
- Native curl_cffi callback
- Persistent HTTP connections
- HTTP/2 and compression negotiation
- Exact axis discovery
- Original MDX Subset engine
- Explicit member-set MDX optimization
- Reduced MDX metadata
- Multi-region benchmark sampling
- Chunk size optimization
- Concurrency optimization
- Bounded request scheduling
- Conservative request-rate control
- Adaptive concurrency reduction
- Detailed HTTP timing
- Benchmark recovery
- Dataset equivalence validation
- Rich live terminal monitor
- Worker relative progress bars
- Raw XML output
- Actual data CSV export
- Actual data Excel export
- Write-only Excel workbook
- Independent CSV/Excel export timing
- Partial dataset recovery on graceful stop

IMPORTANT
---------
This version preserves AllMembers, null values and all seven
selected measures.

No NON EMPTY filtering is applied automatically.

No unverified measure-unit conversion is performed.

Excel and data CSV exports occur after extraction finishes.
Export duration is excluded from request benchmark measurements.

Explicit member-set MDX and axis discovery must be validated
against the live OIC server before production use.

Server-side latency improvements are not guaranteed.

Local rate limits are conservative application limits,
not official OIC service limits.
"""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import json
import math
import os
import re
import statistics
import threading
import time
import uuid

from collections import deque
from concurrent.futures import (
    FIRST_COMPLETED,
    ThreadPoolExecutor,
    wait,
)
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from email.utils import parsedate_to_datetime
from pathlib import Path
from xml.etree import ElementTree as ET

from curl_cffi import CurlInfo
from curl_cffi import requests as curl_requests
from curl_cffi.requests.exceptions import RequestException

from openpyxl import Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill

from rich import box
from rich.bar import Bar
from rich.console import Console, Group
from rich.live import Live
from rich.panel import Panel
from rich.table import Table
from rich.text import Text


# ============================================================
# APPLICATION CONFIGURATION
# ============================================================

APP_NAME = "OIC HIGH-PERFORMANCE LOAD ENGINE"
APP_VERSION = "2.0.0"

ROOT = Path(__file__).resolve().parent

ENV_FILE = ROOT / ".env"
OUTPUT_ROOT = ROOT / "output" / "load_engine"

BASE_URL = "https://onlinewebibs.oic.or.th"
HOSTNAME = "onlinewebibs.oic.or.th"

XMLA_URL = BASE_URL + "/xmla-proxy/api/xmla-proxy"

PAGE_PATH = (
    "/WebForm/Dashboard/Pivot/"
    "DynamicPivotCube.aspx"
)

TOKEN_URL = (
    BASE_URL
    + "/WebForm/Dashboard/Pivot/"
    "PivotCube.asmx/GenToken"
)

CATEGORY = "MV"
CUBE_TYPE = "FY"
OIC_CLASS = "MO"

CATALOG = "PIVOT_202609_MV"
CUBE = "MV_AGGREGATE_FY"

REFERER = (
    BASE_URL
    + PAGE_PATH
    + f"?Class={OIC_CLASS}"
)

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)

DEFAULT_WORKERS = 4
MAX_WORKERS = 8

CONNECT_TIMEOUT = 15
REQUEST_TIMEOUT = 90
DISCOVERY_TIMEOUT = 120

TOKEN_REFRESH_MARGIN = 60

MAX_RESPONSE_BYTES = 100 * 1024 * 1024

MAX_REQUESTS_PER_SECOND = 4.0
WORKER_REQUEST_DELAY = 1.0

MAX_ATTEMPTS = 2

SLOW_REQUEST_MS = 5000
VERY_SLOW_REQUEST_MS = 15000

BENCHMARK_REGION_SIZE = 20
BENCHMARK_MAX_PROFILE_SECONDS = 240
BENCHMARK_PAUSE_SECONDS = 2

DEFAULT_X_CHUNK = 5
DEFAULT_Y_CHUNK = 4

MAX_TASKS_PER_RUN = 100_000
MAX_AXIS_TUPLES = 250_000
MAX_EXPORT_ROWS = 1_000_000

RECENT_HISTORY_LIMIT = 500
RECENT_WINDOW_SECONDS = 60

REFRESH_INTERVAL = 1.0

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"
MD_NS = "urn:schemas-microsoft-com:xml-analysis:mddataset"

ET.register_namespace("soap", SOAP_NS)

X_CAPTIONS = (
    "แบรนด์/ยี่ห้อ",
    "ประเภทกรมธรรม์",
)

Y_CAPTIONS = (
    "ปีบัญชี",
)

MEASURE_CAPTIONS = (
    "จำนวนเคลม",
    "เงินเอาประกันภัย (ล้าน)",
    "เบี้ยประกันภัย (ล้าน)",
    "จำนวนกรมธรรม์",
    "จำนวนเคลมที่ยังไม่ปิด",
    "จำนวนเคลมที่มีการรับคืน",
    "จำนวนกรมธรรม์ที่มีเคลม",
)

CATALOG_PATHS = (
    ROOT / "catalog" / "MV.json",
    ROOT / "output" / "catalog" / "MV.json",
    ROOT / "output" / "field_verification" / "MV_FY.json",
)

IDENTIFIER_PATTERN = re.compile(
    r"^\[[^\]\r\n]+\]\.\[[^\]\r\n]+\]$"
)

MEASURE_PATTERN = re.compile(
    r"^\[Measures\]\.\[[^\]\r\n]+\]$",
    re.IGNORECASE,
)

MEMBER_PART = r"\[(?:[^\]\r\n]|\]\])*\]"

MEMBER_PATTERN = re.compile(
    MEMBER_PART
    + r"(?:(?:\.?&|\.)"
    + MEMBER_PART
    + r")*"
)

XMLA_HEADERS = {
    "Accept": "text/plain, */*; q=0.01",
    "Content-Type": "text/xml",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE_URL,
    "Referer": REFERER,
}

TOKEN_HEADERS = {
    "Accept": "application/json, text/javascript, */*; q=0.01",
    "Content-Type": "application/json; charset=UTF-8",
    "X-Requested-With": "XMLHttpRequest",
    "Origin": BASE_URL,
    "Referer": REFERER,
}

CURL_INFO_NAMES = (
    "NAMELOOKUP_TIME",
    "CONNECT_TIME",
    "APPCONNECT_TIME",
    "PRETRANSFER_TIME",
    "STARTTRANSFER_TIME",
    "TOTAL_TIME",
    "SIZE_DOWNLOAD_T",
    "SIZE_UPLOAD_T",
    "NUM_CONNECTS",
    "HTTP_VERSION",
)

CURL_INFOS = [
    getattr(CurlInfo, name)
    for name in CURL_INFO_NAMES
    if hasattr(CurlInfo, name)
]

TIMING_FIELDS = (
    "queue_wait_ms",
    "rate_wait_ms",
    "token_lock_wait_ms",
    "token_refresh_ms",
    "dns_ms",
    "connect_ms",
    "tls_ms",
    "pretransfer_ms",
    "ttfb_ms",
    "response_wait_estimate_ms",
    "header_wait_ms",
    "first_body_ms",
    "receive_ms",
    "file_write_ms",
    "finalize_ms",
    "parse_ms",
    "transfer_ms",
    "execute_ms",
    "task_total_ms",
)

REQUEST_CSV_FIELDS = (
    "profile",
    "query_strategy",
    "worker_id",
    "request_id",
    "attempt",
    "x_offset",
    "y_offset",
    "x_chunk",
    "y_chunk",
    "started_at",
    "completed_at",
    "status",
    "http_status",
    "http_version",
    "content_type",
    "content_encoding",
    *TIMING_FIELDS,
    "response_bytes",
    "download_bytes",
    "upload_bytes",
    "new_connections",
    "x_members",
    "y_members",
    "axis0_tuples",
    "axis1_tuples",
    "cells",
    "non_null_cells",
    "chunk_kind",
    "response_file",
    "error",
)

DATA_HEADERS = (
    *X_CAPTIONS,
    *Y_CAPTIONS,
    *MEASURE_CAPTIONS,
    "X1 Unique Name",
    "X2 Unique Name",
    "Y1 Unique Name",
)

console = Console(highlight=False)

STOP_EVENT = threading.Event()
THREAD_LOCAL = threading.local()


# ============================================================
# EXCEPTIONS
# ============================================================

class EngineError(RuntimeError):
    pass


class AuthenticationError(EngineError):
    pass


class ResponseError(EngineError):
    pass


class ResponseTooLarge(EngineError):
    pass


class StopRequested(EngineError):
    pass


class ServiceBusyError(EngineError):
    pass


# ============================================================
# DATA MODELS
# ============================================================

@dataclass(frozen=True)
class QuerySpec:

    x: tuple[str, ...]
    y: tuple[str, ...]
    measures: tuple[str, ...]

    mapping_source: str
    query_hash: str


@dataclass(frozen=True)
class Member:

    unique_name: str
    caption: str
    hierarchy: str


@dataclass(frozen=True)
class Profile:

    name: str
    strategy: str

    workers: int

    x_chunk: int
    y_chunk: int


@dataclass(frozen=True)
class RequestTask:

    request_id: str

    x_offset: int
    y_offset: int

    x_chunk: int
    y_chunk: int


@dataclass
class AxisCatalog:

    x: list[tuple[Member, ...]]
    y: list[tuple[Member, ...]]

    snapshot: str | None
    digest: str

    @property
    def x_count(self):

        return len(self.x)

    @property
    def y_count(self):

        return len(self.y)


# ============================================================
# GENERAL UTILITIES
# ============================================================

def utc_now():

    return datetime.now(timezone.utc).isoformat(
        timespec="milliseconds"
    )


def local_time():

    return datetime.now().strftime("%H:%M:%S")


def elapsed_ms(started):

    return round(
        (time.perf_counter() - started) * 1000,
        3,
    )


def format_duration(seconds):

    seconds = max(0, int(seconds))

    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def format_ms(value):

    if value is None:
        return "—"

    if value >= 1000:
        return f"{value / 1000:,.2f}s"

    return f"{value:,.1f}ms"


def format_bytes(value):

    value = float(value)

    for unit in ("B", "KiB", "MiB", "GiB"):

        if value < 1024 or unit == "GiB":
            return f"{value:,.2f} {unit}"

        value /= 1024


def percentile(values, percent=95):

    if not values:
        return None

    ordered = sorted(values)

    index = min(
        len(ordered) - 1,
        max(
            0,
            math.ceil(
                len(ordered) * percent / 100
            ) - 1,
        ),
    )

    return round(ordered[index], 3)


def mean_or_none(values):

    return (
        round(statistics.mean(values), 3)
        if values
        else None
    )


def check_stop(stop_event=None):

    if STOP_EVENT.is_set():

        raise StopRequested(
            "Stop requested"
        )

    if stop_event is not None and stop_event.is_set():

        raise StopRequested(
            "Profile stopped"
        )


def interruptible_sleep(seconds, stop_event=None):

    deadline = time.monotonic() + max(0, seconds)

    while True:

        check_stop(stop_event)

        remaining = deadline - time.monotonic()

        if remaining <= 0:
            return

        STOP_EVENT.wait(
            min(remaining, 0.25)
        )


def read_json(path):

    try:

        result = json.loads(
            path.read_text(
                encoding="utf-8-sig"
            )
        )

        return result if isinstance(result, dict) else {}

    except (OSError, ValueError, TypeError):

        return {}


def write_json(path, data):

    path.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary = path.with_name(
        path.name + ".partial"
    )

    try:

        temporary.write_text(
            json.dumps(
                data,
                ensure_ascii=False,
                indent=2,
                allow_nan=False,
                default=str,
            ),
            encoding="utf-8",
        )

        temporary.replace(path)

    finally:

        temporary.unlink(
            missing_ok=True
        )


def digest_json(value):

    return hashlib.sha256(
        json.dumps(
            value,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8")
    ).hexdigest()


def safe_error(exc):

    if isinstance(exc, EngineError):

        return str(exc)[:300]

    if isinstance(exc, RequestException):

        return (
            type(exc).__name__
            + " | CURL="
            + str(
                getattr(exc, "code", "UNKNOWN")
            )
        )

    if isinstance(exc, OSError):

        return (
            type(exc).__name__
            + " | errno="
            + str(exc.errno)
        )

    return type(exc).__name__


def normalize_caption(value):

    return re.sub(
        r"\s+",
        "",
        str(value or ""),
    ).casefold()


def local_name(tag):

    return tag.rsplit("}", 1)[-1]


def canonical_value(value):

    if value is None:
        return None

    value = str(value).strip()

    if not value:
        return None

    try:

        number = Decimal(value)

        if not number.is_finite():
            return value

        if number == 0:
            return "0"

        sign, digits, exponent = number.as_tuple()

        digits = list(digits)

        while len(digits) > 1 and digits[-1] == 0:

            digits.pop()
            exponent += 1

        return str(
            Decimal(
                (
                    sign,
                    tuple(digits),
                    exponent,
                )
            )
        )

    except InvalidOperation:

        return value


def retry_after_seconds(value, default=15.0):

    if not value:
        return default

    try:

        seconds = float(value)

        if math.isfinite(seconds):
            return max(0.0, seconds)

    except (ValueError, TypeError):

        pass

    try:

        retry_at = parsedate_to_datetime(value)

        if retry_at.tzinfo is None:

            retry_at = retry_at.replace(
                tzinfo=timezone.utc
            )

        return max(
            0.0,
            (
                retry_at
                - datetime.now(timezone.utc)
            ).total_seconds(),
        )

    except Exception:

        return default


# ============================================================
# COMMAND LINE
# ============================================================

def parse_args():

    parser = argparse.ArgumentParser(
        description=APP_NAME
    )

    parser.add_argument(
        "--workers",
        type=int,
        default=DEFAULT_WORKERS,
        help=f"Maximum workers: 1-{MAX_WORKERS}",
    )

    parser.add_argument(
        "--mode",
        choices=("benchmark", "run"),
        default="benchmark",
    )

    args = parser.parse_args()

    if not 1 <= args.workers <= MAX_WORKERS:

        parser.error(
            f"--workers must be between 1 and {MAX_WORKERS}"
        )

    return args


# ============================================================
# ENVIRONMENT
# ============================================================

def load_env():

    values = dict(os.environ)

    if ENV_FILE.is_file():

        for line in ENV_FILE.read_text(
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

    if not values.get(
        "OIC_COOKIE",
        "",
    ).strip():

        raise AuthenticationError(
            "OIC_COOKIE missing from .env"
        )

    return values


def jwt_exp(token):

    try:

        token = token.removeprefix(
            "Bearer "
        ).strip()

        parts = token.split(".")

        if len(parts) != 3:
            return 0

        payload = parts[1]

        payload += "=" * (
            (-len(payload)) % 4
        )

        data = json.loads(
            base64.urlsafe_b64decode(
                payload.encode("ascii")
            )
        )

        return int(data.get("exp", 0))

    except Exception:

        return 0


# ============================================================
# VERIFIED FIELD MAPPING
# ============================================================

def valid_catalog(data):

    return (
        data.get(
            "code",
            data.get("insurance_code"),
        ) == CATEGORY

        and data.get(
            "cube_type",
            "FY",
        ) == CUBE_TYPE

        and data.get("catalog") == CATALOG

        and data.get("cube") == CUBE

        and isinstance(
            data.get("fields"),
            list,
        )

        and bool(data["fields"])

        and data.get("status") not in (
            "ERROR",
            "UNAVAILABLE",
        )
    )


def selectable_verified_field(
    field,
    expected_type,
):

    if field.get("type") != expected_type:
        return False

    if field.get("status") != "VERIFIED":
        return False

    if field.get("source") == "PROFILE":
        return False

    if field.get("verification_method") != "XMLA_EXECUTE":
        return False

    if not field.get("verified_at"):
        return False

    mdx = field.get("mdx")

    if not isinstance(mdx, str):
        return False

    if expected_type == "MEASURE":

        if not MEASURE_PATTERN.fullmatch(mdx):
            return False

        details = field.get("details") or {}

        datafield = str(
            details.get("DATAFIELD")
            or field.get("data_field")
            or ""
        ).upper()

        calculation = str(
            details.get("CALCULATE") or ""
        ).strip().lower()

        summary = str(
            details.get("SUMMARYTYPE") or ""
        ).strip().lower()

        if (
            datafield == "CALCULATE"
            or calculation not in (
                "",
                "0",
                "false",
                "none",
                "null",
            )
            or summary not in ("", "sum")
        ):

            return False

    elif not IDENTIFIER_PATTERN.fullmatch(mdx):

        return False

    elif mdx.casefold().startswith("[measures]."):

        return False

    return True


def resolve_field(
    fields,
    caption,
    expected_type,
):

    target = normalize_caption(caption)

    matches = [
        field
        for field in fields
        if normalize_caption(
            field.get(
                "caption",
                field.get("display", ""),
            )
        ) == target

        and selectable_verified_field(
            field,
            expected_type,
        )
    ]

    original = [
        field
        for field in matches
        if field.get("source") == "GETSTRUCTURE"
    ]

    if original:
        matches = original

    identifiers = {
        field["mdx"]
        for field in matches
    }

    if not identifiers:

        raise EngineError(
            f"Missing verified {expected_type}: {caption}"
        )

    if len(identifiers) != 1:

        raise EngineError(
            f"Ambiguous field mapping: {caption}"
        )

    return next(iter(identifiers))


def load_fixed_query():

    recovery_dir = (
        ROOT
        / "output"
        / "field_verification"
    )

    paths = [
        *CATALOG_PATHS,
        *recovery_dir.glob(
            "MV_FY.recovery.*.json"
        ),
    ]

    candidates = []

    for path in paths:

        if not path.is_file():
            continue

        data = read_json(path)

        if not valid_catalog(data):
            continue

        updated_at = str(
            data.get("report_updated_at")
            or data.get("updated_at")
            or data.get("finished_at")
            or data.get("refreshed_at")
            or ""
        )

        candidates.append(
            (
                updated_at,
                path.stat().st_mtime_ns,
                path,
                data,
            )
        )

    if not candidates:

        raise EngineError(
            "Verified MV/FY catalog not found"
        )

    _, _, path, catalog = max(
        candidates,
        key=lambda item: item[:2],
    )

    fields = catalog["fields"]

    x = tuple(
        resolve_field(
            fields,
            caption,
            "DIMENSION",
        )
        for caption in X_CAPTIONS
    )

    y = tuple(
        resolve_field(
            fields,
            caption,
            "DIMENSION",
        )
        for caption in Y_CAPTIONS
    )

    measures = tuple(
        resolve_field(
            fields,
            caption,
            "MEASURE",
        )
        for caption in MEASURE_CAPTIONS
    )

    if len(set(x + y)) != len(x + y):

        raise EngineError(
            "Duplicate dimension hierarchy"
        )

    if len(set(measures)) != len(measures):

        raise EngineError(
            "Duplicate measures"
        )

    query_hash = digest_json({
        "category": CATEGORY,
        "cube_type": CUBE_TYPE,
        "catalog": CATALOG,
        "cube": CUBE,
        "x": x,
        "y": y,
        "measures": measures,
        "schema_version": 2,
    })[:16]

    return QuerySpec(
        x=x,
        y=y,
        measures=measures,
        mapping_source=str(path),
        query_hash=query_hash,
    )


# ============================================================
# GLOBAL REQUEST CONTROL
# ============================================================

class RequestControl:

    def __init__(self):

        self.lock = threading.Lock()

        self.next_request_at = 0.0
        self.cooldown_until = 0.0

        self.cooldown_reason = ""
        self.service_busy = False

    def wait_slot(self, stop_event=None):

        started = time.perf_counter()

        while True:

            check_stop(stop_event)

            with self.lock:

                if self.service_busy:

                    raise ServiceBusyError(
                        "Server requested backoff; "
                        "current run stopped"
                    )

                now = time.monotonic()

                allowed_at = max(
                    self.next_request_at,
                    self.cooldown_until,
                )

                remaining = allowed_at - now

                if remaining <= 0:

                    self.next_request_at = (
                        now
                        + 1.0 / MAX_REQUESTS_PER_SECOND
                    )

                    return elapsed_ms(started)

            interruptible_sleep(
                min(remaining, 0.25),
                stop_event,
            )

    def set_cooldown(
        self,
        seconds,
        reason,
        *,
        halt=False,
    ):

        with self.lock:

            self.cooldown_until = max(
                self.cooldown_until,
                time.monotonic()
                + max(0, seconds),
            )

            self.cooldown_reason = reason

            if halt:
                self.service_busy = True

    def remaining(self):

        with self.lock:

            return max(
                0.0,
                self.cooldown_until
                - time.monotonic(),
            )


# ============================================================
# HTTP SESSION
# ============================================================

def create_session(cookie):

    session = curl_requests.Session(
        verify=True,
        curl_infos=CURL_INFOS,
    )

    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept-Language": "th-TH,th;q=0.9",
    })

    cookie = cookie.removeprefix(
        "Cookie:"
    ).strip()

    for part in cookie.split(";"):

        name, separator, value = (
            part.strip().partition("=")
        )

        if not separator or not name:
            continue

        session.cookies.set(
            name.strip(),
            value.strip(),
            domain=HOSTNAME,
            path="/",
            secure=True,
        )

    return session


# ============================================================
# AUTHENTICATION MANAGER
# ============================================================

class AuthenticationManager:

    def __init__(
        self,
        env,
        control,
    ):

        self.cookie = env["OIC_COOKIE"]

        self.token = env.get(
            "OIC_BEARER_TOKEN",
            "",
        ).removeprefix(
            "Bearer "
        ).strip()

        self.control = control

        self.lock = threading.Lock()

        self.refresh_count = 0
        self.refresh_total_ms = 0.0

    def refresh_token(
        self,
        session,
        stop_event=None,
    ):

        started = time.perf_counter()

        self.control.wait_slot(
            stop_event
        )

        response = session.post(
            TOKEN_URL,

            json={
                "_tk": "Init",

                "_pageUrl": (
                    PAGE_PATH
                    + f"?Class={OIC_CLASS}"
                ),
            },

            headers=TOKEN_HEADERS,

            timeout=(
                CONNECT_TIMEOUT,
                45,
            ),

            allow_redirects=False,
        )

        try:

            status = response.status_code

            if status in (429, 503):

                self.control.set_cooldown(
                    retry_after_seconds(
                        response.headers.get(
                            "Retry-After"
                        ),
                        30,
                    ),
                    f"GenToken HTTP {status}",
                    halt=True,
                )

                raise ServiceBusyError(
                    f"GenToken HTTP {status}"
                )

            if status != 200:

                raise AuthenticationError(
                    f"GenToken HTTP {status}"
                )

            try:

                token = response.json().get(
                    "d",
                    "",
                )

            except Exception as exc:

                raise AuthenticationError(
                    "Invalid GenToken response"
                ) from exc

            if token == "SESSION_EXPIRED":

                raise AuthenticationError(
                    "Cookie session expired"
                )

            if not isinstance(token, str):

                raise AuthenticationError(
                    "GenToken did not return JWT"
                )

            token = token.removeprefix(
                "Bearer "
            ).strip()

            if jwt_exp(token) <= time.time() + 30:

                raise AuthenticationError(
                    "Invalid or expired JWT"
                )

            self.token = token

            duration = elapsed_ms(started)

            self.refresh_count += 1
            self.refresh_total_ms += duration

            return duration

        finally:

            response.close()

    def get_token(
        self,
        session,
        rejected_token=None,
        stop_event=None,
    ):

        started = time.perf_counter()

        with self.lock:

            check_stop(stop_event)

            lock_wait_ms = elapsed_ms(
                started
            )

            refresh_ms = 0.0

            expired = (
                jwt_exp(self.token)
                <= time.time()
                + TOKEN_REFRESH_MARGIN
            )

            rejected = (
                rejected_token is not None
                and self.token == rejected_token
            )

            if expired or rejected:

                refresh_ms = self.refresh_token(
                    session,
                    stop_event,
                )

            return (
                self.token,
                lock_wait_ms,
                refresh_ms,
            )


# ============================================================
# MDX QUERY OPTIMIZATION
# ============================================================

def build_axis_set(identifiers):

    members = [
        identifier + ".AllMembers"
        for identifier in identifiers
    ]

    if not members:

        raise EngineError(
            "Empty dimension set"
        )

    result = members[0]

    for member in members[1:]:

        result = (
            f"CrossJoin({result},{member})"
        )

    return result


def validate_member_identifier(value):

    if not MEMBER_PATTERN.fullmatch(value):

        raise ResponseError(
            "Unsupported member identifier syntax"
        )

    return value


def build_explicit_set(members):

    items = []

    for member_tuple in members:

        identifiers = [
            validate_member_identifier(
                member.unique_name
            )
            for member in member_tuple
        ]

        if len(identifiers) == 1:

            items.append(
                identifiers[0]
            )

        else:

            items.append(
                "("
                + ",".join(identifiers)
                + ")"
            )

    if not items:

        raise EngineError(
            "Cannot build empty MDX member set"
        )

    return (
        "{"
        + ",".join(items)
        + "}"
    )


def build_count_mdx(query):

    x_set = build_axis_set(query.x)
    y_set = build_axis_set(query.y)

    return (
        "WITH "
        "MEMBER [Measures].[__LoadEngineXCount] "
        f"AS Count({x_set}) "

        "MEMBER [Measures].[__LoadEngineYCount] "
        f"AS Count({y_set}) "

        "SELECT {"
        "[Measures].[__LoadEngineXCount],"
        "[Measures].[__LoadEngineYCount]"
        "} ON COLUMNS "

        f"FROM [{CUBE}] "

        "CELL PROPERTIES VALUE"
    )


def build_axis_discovery_mdx(identifiers):

    axis_set = build_axis_set(
        identifiers
    )

    return (
        "WITH "
        "MEMBER [Measures].[__LoadEngineAxis] AS 0 "

        "SELECT "
        "{[Measures].[__LoadEngineAxis]} "
        "ON COLUMNS, "

        f"{axis_set} "

        "DIMENSION PROPERTIES "
        "HIERARCHY_UNIQUE_NAME "
        "ON ROWS "

        f"FROM [{CUBE}] "

        "CELL PROPERTIES VALUE"
    )


def build_mdx(
    query,
    axes,
    task,
    strategy,
):

    measures = (
        "{"
        + ",".join(query.measures)
        + "}"
    )

    if strategy == "subset":

        x_set = (
            f"Subset("
            f"{build_axis_set(query.x)},"
            f"{task.x_offset},"
            f"{task.x_chunk})"
        )

        y_set = (
            f"Subset("
            f"{build_axis_set(query.y)},"
            f"{task.y_offset},"
            f"{task.y_chunk})"
        )

        return (
            f"SELECT {y_set} "

            "DIMENSION PROPERTIES "
            "PARENT_UNIQUE_NAME,"
            "HIERARCHY_UNIQUE_NAME,"
            "MEMBER_VALUE ON COLUMNS, "

            f"CrossJoin({x_set},{measures}) "

            "DIMENSION PROPERTIES "
            "PARENT_UNIQUE_NAME,"
            "HIERARCHY_UNIQUE_NAME,"
            "MEMBER_VALUE ON ROWS "

            f"FROM [{CUBE}] "

            "CELL PROPERTIES "
            "VALUE,FORMAT_STRING,LANGUAGE"
        )

    if strategy != "explicit":

        raise EngineError(
            f"Unknown MDX strategy: {strategy}"
        )

    x_members = axes.x[
        task.x_offset:
        task.x_offset + task.x_chunk
    ]

    y_members = axes.y[
        task.y_offset:
        task.y_offset + task.y_chunk
    ]

    x_set = build_explicit_set(
        x_members
    )

    y_set = build_explicit_set(
        y_members
    )

    # Avoid repeatedly evaluating the entire
    # CrossJoin and Subset expression.
    #
    # Preserve original members and measures.
    # Do not apply NON EMPTY automatically.
    #
    # Reduced properties must pass the
    # dataset equivalence benchmark.

    return (
        f"SELECT {y_set} ON COLUMNS, "

        f"CrossJoin({x_set},{measures}) "
        "ON ROWS "

        f"FROM [{CUBE}] "

        "CELL PROPERTIES VALUE"
    )


def build_soap(mdx):

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

    for name, value in {
        "Catalog": CATALOG,
        "ShowHiddenCubes": "True",
        "Timeout": "3600",
        "LocaleIdentifier": "9",
    }.items():

        ET.SubElement(
            properties,
            f"{{{XMLA_NS}}}{name}",
        ).text = value

    return ET.tostring(
        envelope,
        encoding="utf-8",
    )


# ============================================================
# CURL TIMING INFORMATION
# ============================================================

def get_curl_info(response, name):

    enum = getattr(
        CurlInfo,
        name,
        None,
    )

    if enum is None:
        return None

    return (
        getattr(response, "infos", {})
        or {}
    ).get(enum)


def extract_curl_timings(response):

    def milliseconds(name):

        value = get_curl_info(
            response,
            name,
        )

        if not isinstance(value, (int, float)):
            return None

        return round(
            value * 1000,
            3,
        )

    def difference(end, start):

        if end is None or start is None:
            return None

        return round(
            max(0.0, end - start),
            3,
        )

    dns = milliseconds(
        "NAMELOOKUP_TIME"
    )

    connect = milliseconds(
        "CONNECT_TIME"
    )

    tls = milliseconds(
        "APPCONNECT_TIME"
    )

    pretransfer = milliseconds(
        "PRETRANSFER_TIME"
    )

    ttfb = milliseconds(
        "STARTTRANSFER_TIME"
    )

    total = milliseconds(
        "TOTAL_TIME"
    )

    connections = get_curl_info(
        response,
        "NUM_CONNECTS",
    )

    downloaded = get_curl_info(
        response,
        "SIZE_DOWNLOAD_T",
    )

    uploaded = get_curl_info(
        response,
        "SIZE_UPLOAD_T",
    )

    reused = connections == 0

    return {
        "dns_ms": (
            0.0 if reused else dns
        ),

        "connect_ms": (
            0.0
            if reused
            else difference(connect, dns)
        ),

        "tls_ms": (
            0.0
            if reused or tls == 0
            else difference(tls, connect)
        ),

        "pretransfer_ms": pretransfer,

        "ttfb_ms": ttfb,

        "response_wait_estimate_ms": difference(
            ttfb,
            pretransfer,
        ),

        "transfer_ms": total,

        "download_bytes": (
            int(downloaded)
            if isinstance(
                downloaded,
                (int, float),
            )
            else None
        ),

        "upload_bytes": (
            int(uploaded)
            if isinstance(
                uploaded,
                (int, float),
            )
            else None
        ),

        "new_connections": (
            int(connections)
            if isinstance(
                connections,
                (int, float),
            )
            else None
        ),
    }


def http_version_name(value):

    versions = {
        0: "UNKNOWN",
        1: "HTTP/1.0",
        2: "HTTP/1.1",
        3: "HTTP/2",
        30: "HTTP/3",
    }

    try:

        return versions.get(
            int(value),
            str(value),
        )

    except (TypeError, ValueError):

        return "UNKNOWN"


# ============================================================
# NATIVE CALLBACK REQUEST ENGINE
# ============================================================

def execute_request(
    session,
    token,
    payload,
    destination,
    *,
    timeout=REQUEST_TIMEOUT,
    stop_event=None,
):

    started = time.perf_counter()

    destination.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    temporary = destination.with_name(
        destination.name + ".partial"
    )

    headers = dict(XMLA_HEADERS)

    headers["Authorization"] = (
        "Bearer " + token
    )

    response = None

    response_bytes = 0

    first_body_at = None
    last_body_at = None

    file_write_ms = 0.0
    finalize_ms = None

    limit_exceeded = False
    callback_error = None

    try:

        with temporary.open(
            "wb",
            buffering=256 * 1024,
        ) as handle:

            transfer_started = time.perf_counter()

            def content_callback(chunk):

                nonlocal response_bytes
                nonlocal first_body_at
                nonlocal last_body_at
                nonlocal file_write_ms
                nonlocal limit_exceeded
                nonlocal callback_error

                if STOP_EVENT.is_set():

                    return 0

                if (
                    stop_event is not None
                    and stop_event.is_set()
                ):

                    return 0

                if not chunk:

                    return 0

                if first_body_at is None:

                    first_body_at = (
                        time.perf_counter()
                    )

                response_bytes += len(chunk)

                if response_bytes > MAX_RESPONSE_BYTES:

                    limit_exceeded = True

                    return 0

                write_started = (
                    time.perf_counter()
                )

                try:

                    written = handle.write(
                        chunk
                    )

                except OSError as exc:

                    callback_error = exc

                    return 0

                finally:

                    file_write_ms += (
                        time.perf_counter()
                        - write_started
                    ) * 1000

                last_body_at = (
                    time.perf_counter()
                )

                return written

            response = session.post(
                XMLA_URL,

                data=payload,

                headers=headers,

                content_callback=content_callback,

                timeout=(
                    CONNECT_TIMEOUT,
                    timeout,
                ),

                allow_redirects=False,

                accept_encoding=(
                    "gzip, deflate, br"
                ),
            )

            transfer_wall_ms = elapsed_ms(
                transfer_started
            )

            timings = extract_curl_timings(
                response
            )

            flush_started = time.perf_counter()

            handle.flush()

            file_write_ms += (
                time.perf_counter()
                - flush_started
            ) * 1000

        check_stop(stop_event)

        if limit_exceeded:

            raise ResponseTooLarge(
                "XMLA response exceeded local size limit"
            )

        if callback_error:

            raise ResponseError(
                "Unable to save XML response"
            )

        status = int(
            response.status_code
        )

        if status == 200:

            finalize_started = (
                time.perf_counter()
            )

            temporary.replace(
                destination
            )

            finalize_ms = elapsed_ms(
                finalize_started
            )

        first_body_ms = (
            round(
                (
                    first_body_at
                    - transfer_started
                ) * 1000,
                3,
            )
            if first_body_at is not None
            else None
        )

        receive_ms = (
            round(
                (
                    last_body_at
                    - first_body_at
                ) * 1000,
                3,
            )
            if (
                first_body_at is not None
                and last_body_at is not None
            )
            else None
        )

        return {
            **timings,

            "http_status": status,

            "http_version": http_version_name(
                getattr(
                    response,
                    "http_version",
                    None,
                )
            ),

            "content_type": response.headers.get(
                "Content-Type",
                "",
            ),

            "content_encoding": response.headers.get(
                "Content-Encoding",
                "",
            ),

            "retry_after": response.headers.get(
                "Retry-After"
            ),

            # Native callback exposes exact libcurl
            # first-byte timing, not an independent
            # headers-complete event.
            "header_wait_ms": None,

            "first_body_ms": first_body_ms,

            "receive_ms": receive_ms,

            "file_write_ms": round(
                file_write_ms,
                3,
            ),

            "finalize_ms": finalize_ms,

            "transfer_wall_ms": transfer_wall_ms,

            "execute_ms": elapsed_ms(started),

            "response_bytes": response_bytes,

            "response_file": (
                str(destination)
                if status == 200
                else None
            ),
        }

    except RequestException:

        check_stop(stop_event)

        if limit_exceeded:

            raise ResponseTooLarge(
                "Response exceeded local byte limit"
            )

        if callback_error:

            raise ResponseError(
                "XML file write failed"
            )

        raise

    finally:

        if response is not None:

            try:

                response.close()

            except Exception:

                pass

        temporary.unlink(
            missing_ok=True
        )


# ============================================================
# XMLA PARSER
# ============================================================

def load_xml_dataset(path):

    try:

        root = ET.parse(
            path
        ).getroot()

    except (
        OSError,
        ET.ParseError,
    ) as exc:

        raise ResponseError(
            "Invalid XMLA response"
        ) from exc

    for element in root.iter():

        if local_name(element.tag) in (
            "Fault",
            "Error",
        ):

            raise ResponseError(
                "XMLA response contains SOAP or OLAP error"
            )

    dataset = root.find(
        f".//{{{MD_NS}}}root"
    )

    if dataset is None:

        raise ResponseError(
            "XMLA multidimensional dataset missing"
        )

    return root, dataset


def read_axis(dataset, axis_name):

    axis = dataset.find(
        f"./{{{MD_NS}}}Axes/"
        f"{{{MD_NS}}}Axis"
        f"[@name='{axis_name}']"
    )

    if axis is None:

        raise ResponseError(
            f"Missing XMLA {axis_name}"
        )

    tuples = []

    for item in axis.findall(
        f"{{{MD_NS}}}Tuples/"
        f"{{{MD_NS}}}Tuple"
    ):

        members = []

        for member in item.findall(
            f"{{{MD_NS}}}Member"
        ):

            unique_name = member.findtext(
                f"{{{MD_NS}}}UName"
            )

            if not unique_name:

                raise ResponseError(
                    "XMLA member has no unique name"
                )

            members.append(
                Member(
                    unique_name=unique_name,

                    caption=member.findtext(
                        f"{{{MD_NS}}}Caption",
                        default=unique_name,
                    ),

                    hierarchy=member.get(
                        "Hierarchy",
                        "",
                    ),
                )
            )

        tuples.append(
            tuple(members)
        )

    return tuples


def read_cells(dataset):

    cells = {}

    for cell in dataset.findall(
        f"./{{{MD_NS}}}CellData/"
        f"{{{MD_NS}}}Cell"
    ):

        try:

            ordinal = int(
                cell.get(
                    "CellOrdinal",
                    "-1",
                )
            )

        except ValueError as exc:

            raise ResponseError(
                "Invalid CellOrdinal"
            ) from exc

        if ordinal < 0 or ordinal in cells:

            raise ResponseError(
                "Invalid or duplicate CellOrdinal"
            )

        value = cell.find(
            f"{{{MD_NS}}}Value"
        )

        cells[ordinal] = (
            value.text
            if value is not None
            else None
        )

    return cells


def read_snapshot(root):

    return next(
        (
            element.text.strip()
            for element in root.iter()

            if (
                local_name(element.tag)
                == "LastDataUpdate"

                and element.text
            )
        ),
        None,
    )


def tuple_key(members):

    return tuple(
        member.unique_name
        for member in members
    )


def is_measure(member):

    return (
        member.hierarchy.casefold()
        == "[measures]"

        or member.unique_name.casefold().startswith(
            "[measures]."
        )
    )


def parse_response(
    path,
    query,
    axes,
    task,
):

    started = time.perf_counter()

    root, dataset = load_xml_dataset(
        path
    )

    axis0 = read_axis(
        dataset,
        "Axis0",
    )

    axis1 = read_axis(
        dataset,
        "Axis1",
    )

    cells = read_cells(
        dataset
    )

    snapshot = read_snapshot(
        root
    )

    if (
        axes.snapshot
        and snapshot
        and snapshot != axes.snapshot
    ):

        raise ResponseError(
            "Cube snapshot changed during extraction"
        )

    expected_x = {
        tuple_key(item)
        for item in axes.x[
            task.x_offset:
            task.x_offset + task.x_chunk
        ]
    }

    expected_y = {
        tuple_key(item)
        for item in axes.y[
            task.y_offset:
            task.y_offset + task.y_chunk
        ]
    }

    measure_indexes = {
        name.casefold(): index

        for index, name in enumerate(
            query.measures
        )
    }

    row_keys = []
    returned_x = set()
    measures_by_x = {}

    for row in axis1:

        dimensions = tuple(
            member
            for member in row
            if not is_measure(member)
        )

        measures = [
            member
            for member in row
            if is_measure(member)
        ]

        if len(dimensions) != len(query.x):

            raise ResponseError(
                "Unexpected row dimension count"
            )

        if len(measures) != 1:

            raise ResponseError(
                "Expected one measure per row tuple"
            )

        x_key = tuple_key(
            dimensions
        )

        measure_name = (
            measures[0]
            .unique_name
            .casefold()
        )

        if measure_name not in measure_indexes:

            raise ResponseError(
                "Unexpected measure in XMLA response"
            )

        measure_index = measure_indexes[
            measure_name
        ]

        row_keys.append(
            (
                x_key,
                measure_index,
            )
        )

        returned_x.add(
            x_key
        )

        measures_by_x.setdefault(
            x_key,
            set(),
        ).add(
            measure_index
        )

    column_keys = []

    for column in axis0:

        if len(column) != len(query.y):

            raise ResponseError(
                "Unexpected column dimension count"
            )

        column_keys.append(
            tuple_key(column)
        )

    if len(set(row_keys)) != len(row_keys):

        raise ResponseError(
            "Duplicate row tuple"
        )

    if len(set(column_keys)) != len(column_keys):

        raise ResponseError(
            "Duplicate column tuple"
        )

    returned_y = set(
        column_keys
    )

    if returned_x != expected_x:

        raise ResponseError(
            "X axis member coverage mismatch"
        )

    if returned_y != expected_y:

        raise ResponseError(
            "Y axis member coverage mismatch"
        )

    expected_measures = set(
        range(
            len(query.measures)
        )
    )

    for returned_measures in measures_by_x.values():

        if returned_measures != expected_measures:

            raise ResponseError(
                "Incomplete measure set"
            )

    ordinal_limit = (
        len(row_keys)
        * len(column_keys)
    )

    if any(
        ordinal >= ordinal_limit
        for ordinal in cells
    ):

        raise ResponseError(
            "CellOrdinal outside returned axes"
        )

    rows = {}

    non_null_cells = 0

    column_count = len(
        column_keys
    )

    for row_index, (
        x_key,
        measure_index,
    ) in enumerate(row_keys):

        for column_index, y_key in enumerate(
            column_keys
        ):

            ordinal = (
                row_index
                * column_count
                + column_index
            )

            value = canonical_value(
                cells.get(ordinal)
            )

            if value is not None:

                non_null_cells += 1

            key = x_key + y_key

            if key not in rows:

                rows[key] = [
                    None
                ] * len(query.measures)

            rows[key][measure_index] = value

    if not returned_x or not returned_y:

        chunk_kind = "EMPTY_PAGE"

    elif non_null_cells == 0:

        chunk_kind = "NO_CELLS"

    else:

        chunk_kind = "DATA"

    return {
        "rows": rows,

        "x_members": len(returned_x),
        "y_members": len(returned_y),

        "axis0_tuples": len(axis0),
        "axis1_tuples": len(axis1),

        "cells": len(cells),

        "non_null_cells": non_null_cells,

        "chunk_kind": chunk_kind,

        "last_data_update": snapshot,

        "parse_ms": elapsed_ms(started),
    }


# ============================================================
# AUTHENTICATED DISCOVERY REQUEST
# ============================================================

def discovery_request(
    session,
    auth,
    control,
    mdx,
    destination,
):

    rejected_token = None

    for attempt in range(
        1,
        MAX_ATTEMPTS + 1,
    ):

        check_stop()

        token, lock_wait, refresh_ms = (
            auth.get_token(
                session,
                rejected_token,
            )
        )

        rate_wait = control.wait_slot()

        result = execute_request(
            session,

            token,

            build_soap(mdx),

            destination,

            timeout=DISCOVERY_TIMEOUT,
        )

        result.update({
            "token_lock_wait_ms": lock_wait,

            "token_refresh_ms": refresh_ms,

            "rate_wait_ms": rate_wait,

            "attempt": attempt,
        })

        status = result["http_status"]

        if status == 200:

            return result

        if status in (401, 403):

            if attempt >= MAX_ATTEMPTS:

                raise AuthenticationError(
                    f"XMLA authentication HTTP {status}"
                )

            rejected_token = token

            continue

        if status in (429, 503):

            control.set_cooldown(
                retry_after_seconds(
                    result.get("retry_after"),
                    30,
                ),

                f"Discovery HTTP {status}",

                halt=True,
            )

            raise ServiceBusyError(
                f"Discovery HTTP {status}"
            )

        raise ResponseError(
            f"Discovery HTTP {status}"
        )

    raise ResponseError(
        "Discovery attempts exhausted"
    )


# ============================================================
# EXACT AXIS DISCOVERY
# ============================================================

def discover_axes(
    query,
    auth,
    control,
    run_dir,
):

    session = create_session(
        auth.cookie
    )

    discovery_dir = (
        run_dir
        / "discovery"
    )

    discovery_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    results = {}

    snapshots = set()

    try:

        count_path = (
            discovery_dir
            / "axis_counts.xml"
        )

        results["counts"] = discovery_request(
            session,

            auth,

            control,

            build_count_mdx(query),

            count_path,
        )

        root, dataset = load_xml_dataset(
            count_path
        )

        snapshot = read_snapshot(
            root
        )

        if snapshot:

            snapshots.add(snapshot)

        count_axis = read_axis(
            dataset,
            "Axis0",
        )

        count_cells = read_cells(
            dataset
        )

        counts = {}

        for ordinal, member_tuple in enumerate(
            count_axis
        ):

            if len(member_tuple) != 1:

                raise ResponseError(
                    "Invalid axis-count tuple"
                )

            raw = count_cells.get(
                ordinal
            )

            if raw is None:

                raise ResponseError(
                    "Axis count is null"
                )

            try:

                number = Decimal(raw)

            except InvalidOperation as exc:

                raise ResponseError(
                    "Invalid axis-count value"
                ) from exc

            if (
                not number.is_finite()

                or number < 0

                or number != number.to_integral_value()
            ):

                raise ResponseError(
                    "Invalid axis cardinality"
                )

            counts[
                member_tuple[0]
                .unique_name
                .casefold()
            ] = int(number)

        x_count = counts.get(
            "[measures].[__loadenginexcount]"
        )

        y_count = counts.get(
            "[measures].[__loadengineycount]"
        )

        if x_count is None or y_count is None:

            raise ResponseError(
                "Unable to discover X/Y cardinality"
            )

        if not x_count or not y_count:

            raise ResponseError(
                "Fixed query returned empty axes"
            )

        if (
            x_count > MAX_AXIS_TUPLES
            or y_count > MAX_AXIS_TUPLES
        ):

            raise EngineError(
                "Axis cardinality exceeded local limit"
            )

        if x_count * y_count > MAX_EXPORT_ROWS:

            raise EngineError(
                "Dataset exceeds local export row limit"
            )

        discovered = {}

        for axis_name, identifiers, expected in (
            ("x", query.x, x_count),
            ("y", query.y, y_count),
        ):

            console.print(
                f"[cyan]DISCOVER {axis_name.upper()}[/cyan]"
                f" | {expected:,} expected tuples"
            )

            path = (
                discovery_dir
                / f"axis_{axis_name}.xml"
            )

            results[axis_name] = discovery_request(
                session,

                auth,

                control,

                build_axis_discovery_mdx(
                    identifiers
                ),

                path,
            )

            root, dataset = load_xml_dataset(
                path
            )

            snapshot = read_snapshot(
                root
            )

            if snapshot:

                snapshots.add(snapshot)

            axis = read_axis(
                dataset,
                "Axis1",
            )

            if len(axis) != expected:

                raise ResponseError(
                    f"{axis_name.upper()} axis discovery "
                    f"returned {len(axis):,}/{expected:,} tuples"
                )

            for members in axis:

                if len(members) != len(identifiers):

                    raise ResponseError(
                        "Axis dimension count mismatch"
                    )

                for member in members:

                    validate_member_identifier(
                        member.unique_name
                    )

            keys = [
                tuple_key(members)
                for members in axis
            ]

            if len(set(keys)) != len(keys):

                raise ResponseError(
                    "Duplicate discovered axis tuples"
                )

            discovered[axis_name] = axis

        if len(snapshots) > 1:

            raise ResponseError(
                "Cube snapshot changed during discovery"
            )

        snapshot = next(
            iter(snapshots),
            None,
        )

        axis_digest = digest_json({
            "x": [
                tuple_key(item)
                for item in discovered["x"]
            ],

            "y": [
                tuple_key(item)
                for item in discovered["y"]
            ],
        })

        axes = AxisCatalog(
            x=discovered["x"],

            y=discovered["y"],

            snapshot=snapshot,

            digest=axis_digest,
        )

        write_json(
            run_dir / "axis_discovery.json",
            {
                "query_hash": query.query_hash,

                "axis_digest": axis_digest,

                "snapshot": snapshot,

                "x_members": axes.x_count,

                "y_members": axes.y_count,

                "measures": len(
                    query.measures
                ),

                "expected_rows": (
                    axes.x_count
                    * axes.y_count
                ),

                "expected_coordinates": (
                    axes.x_count
                    * axes.y_count
                    * len(query.measures)
                ),

                "requests": results,
            },
        )

        write_json(
            run_dir / "axis_members.json",
            {
                "x": [
                    [
                        asdict(member)
                        for member in members
                    ]
                    for members in axes.x
                ],

                "y": [
                    [
                        asdict(member)
                        for member in members
                    ]
                    for members in axes.y
                ],
            },
        )

        return axes

    finally:

        session.close()


# ============================================================
# WORKER STATE
# ============================================================

class EngineState:

    def __init__(
        self,
        profile,
        total_tasks,
    ):

        self.profile = profile

        self.total_tasks = total_tasks

        self.started = time.monotonic()

        self.lock = threading.Lock()

        self.workers = {}

        for worker_id in range(
            1,
            profile.workers + 1,
        ):

            self.workers[worker_id] = {
                "worker_id": worker_id,

                "status": "STARTING",
                "phase": "IDLE",

                "x_offset": None,
                "y_offset": None,

                "completed": 0,
                "failed": 0,
                "requests": 0,

                "cells": 0,

                "last_http": None,

                "last_request_ms": None,
                "last_ttfb_ms": None,

                "request_started": None,

                "last_event": "Waiting for task",
                "event_time": local_time(),
            }

        self.events = deque(
            maxlen=20
        )

        self.recent = deque(
            maxlen=RECENT_HISTORY_LIMIT
        )

        self.durations = []

        self.timing_sums = {}
        self.timing_counts = {}

        self.requests = 0

        self.successful_requests = 0
        self.failed_requests = 0
        self.stopped_requests = 0

        self.completed = 0
        self.failed = 0

        self.cells = 0
        self.non_null_cells = 0

        self.response_bytes = 0
        self.download_bytes = 0

        self.data_chunks = 0
        self.no_cell_chunks = 0
        self.empty_pages = 0

        self.slow_requests = 0

        self.connection_samples = 0
        self.reused_transfers = 0
        self.new_connections = 0

        self.fatal_error = None

        self.active_limit = profile.workers

    def update_worker(
        self,
        worker_id,
        **values,
    ):

        with self.lock:

            worker = self.workers[
                worker_id
            ]

            worker.update(
                values
            )

            if "last_event" in values:

                worker["event_time"] = (
                    local_time()
                )

                self.events.append({
                    "time": worker["event_time"],

                    "worker": worker_id,

                    "message": worker[
                        "last_event"
                    ],
                })

    def record_request(
        self,
        record,
    ):

        with self.lock:

            self.requests += 1

            status = record.get(
                "status"
            )

            if status == "OK":

                self.successful_requests += 1

            elif status == "STOPPED":

                self.stopped_requests += 1

            else:

                self.failed_requests += 1

            worker = self.workers[
                record["worker_id"]
            ]

            worker["requests"] += 1

            duration = record.get(
                "execute_ms"
            )

            if duration is not None:

                self.durations.append(
                    duration
                )

                if duration >= SLOW_REQUEST_MS:

                    self.slow_requests += 1

            worker["last_request_ms"] = duration

            worker["last_ttfb_ms"] = record.get(
                "ttfb_ms"
            )

            worker["request_started"] = None

            if record.get("http_status") is not None:

                worker["last_http"] = record[
                    "http_status"
                ]

            for field in TIMING_FIELDS:

                value = record.get(
                    field
                )

                if not isinstance(
                    value,
                    (int, float),
                ):

                    continue

                self.timing_sums[field] = (
                    self.timing_sums.get(
                        field,
                        0.0,
                    )
                    + value
                )

                self.timing_counts[field] = (
                    self.timing_counts.get(
                        field,
                        0,
                    )
                    + 1
                )

            connections = record.get(
                "new_connections"
            )

            if connections is not None:

                self.connection_samples += 1

                self.new_connections += connections

                if connections == 0:

                    self.reused_transfers += 1

            self.recent.append({
                "time": time.monotonic(),

                "execute_ms": duration,

                "ttfb_ms": record.get(
                    "ttfb_ms"
                ),

                "status": status,
            })

    def finish_task(
        self,
        worker_id,
        result,
    ):

        with self.lock:

            worker = self.workers[
                worker_id
            ]

            if result is None:

                self.failed += 1

                worker["failed"] += 1

                worker["phase"] = "IDLE"

                worker["status"] = "FAILED"

                worker["request_started"] = None

                return

            self.completed += 1

            worker["completed"] += 1

            worker["cells"] += result[
                "non_null_cells"
            ]

            self.cells += result[
                "cells"
            ]

            self.non_null_cells += result[
                "non_null_cells"
            ]

            self.response_bytes += result.get(
                "response_bytes",
                0,
            )

            self.download_bytes += (
                result.get(
                    "download_bytes"
                )
                or 0
            )

            kind = result[
                "chunk_kind"
            ]

            if kind == "DATA":

                self.data_chunks += 1

            elif kind == "NO_CELLS":

                self.no_cell_chunks += 1

            elif kind == "EMPTY_PAGE":

                self.empty_pages += 1

            worker["phase"] = "IDLE"
            worker["status"] = "RUNNING"

            worker["request_started"] = None

    def set_fatal(self, message):

        with self.lock:

            if self.fatal_error is None:

                self.fatal_error = message

    def set_active_limit(self, value):

        with self.lock:

            self.active_limit = value

    def snapshot(self):

        with self.lock:

            return {
                "elapsed": max(
                    0.001,
                    time.monotonic()
                    - self.started,
                ),

                "total_tasks": self.total_tasks,

                "completed_tasks": self.completed,

                "failed_tasks": self.failed,

                "requests": self.requests,

                "successful_requests": (
                    self.successful_requests
                ),

                "failed_requests": (
                    self.failed_requests
                ),

                "stopped_requests": (
                    self.stopped_requests
                ),

                "cells": self.cells,

                "non_null_cells": (
                    self.non_null_cells
                ),

                "response_bytes": (
                    self.response_bytes
                ),

                "download_bytes": (
                    self.download_bytes
                ),

                "data_chunks": self.data_chunks,

                "no_cell_chunks": (
                    self.no_cell_chunks
                ),

                "empty_pages": self.empty_pages,

                "slow_requests": self.slow_requests,

                "durations": list(
                    self.durations
                ),

                "recent": list(
                    self.recent
                ),

                "timing": {
                    field: round(
                        self.timing_sums[field]
                        / self.timing_counts[field],
                        3,
                    )

                    for field in self.timing_sums
                },

                "new_connections": (
                    self.new_connections
                ),

                "connection_samples": (
                    self.connection_samples
                ),

                "connection_reuse_percent": (
                    round(
                        self.reused_transfers
                        / self.connection_samples
                        * 100,
                        2,
                    )

                    if self.connection_samples
                    else None
                ),

                "active_limit": self.active_limit,

                "workers": [
                    dict(worker)

                    for worker in (
                        self.workers.values()
                    )
                ],

                "events": list(
                    self.events
                ),

                "fatal_error": self.fatal_error,
            }


# ============================================================
# METRICS AGGREGATION
# ============================================================

def aggregate_metrics(snapshot):

    durations = snapshot.pop(
        "durations"
    )

    recent_records = snapshot.pop(
        "recent"
    )

    now = time.monotonic()

    recent = [
        record["execute_ms"]

        for record in recent_records

        if (
            now - record["time"]
            <= RECENT_WINDOW_SECONDS
        )

        and record["execute_ms"] is not None
    ]

    elapsed = snapshot[
        "elapsed"
    ]

    workers = snapshot[
        "workers"
    ]

    max_completed = max(
        (
            worker["completed"]
            for worker in workers
        ),
        default=0,
    )

    slow_inflight = 0

    for worker in workers:

        worker["relative_progress"] = (
            worker["completed"]
            / max_completed

            if max_completed
            else 0.0
        )

        started = worker.get(
            "request_started"
        )

        current_ms = None

        if (
            worker["phase"] == "REQUESTING"
            and started is not None
        ):

            current_ms = round(
                max(
                    0.0,
                    now - started,
                ) * 1000,
                3,
            )

            if current_ms >= SLOW_REQUEST_MS:

                slow_inflight += 1

        worker["current_request_ms"] = (
            current_ms
        )

    completed = snapshot[
        "completed_tasks"
    ]

    total = snapshot[
        "total_tasks"
    ]

    remaining = max(
        0,
        total - completed
        - snapshot["failed_tasks"],
    )

    chunk_rate = completed / elapsed

    snapshot.update({
        "average_ms": mean_or_none(
            durations
        ),

        "median_ms": (
            round(
                statistics.median(durations),
                3,
            )
            if durations
            else None
        ),

        "p95_ms": percentile(
            durations
        ),

        "max_ms": (
            max(durations)
            if durations
            else None
        ),

        "recent_average_ms": mean_or_none(
            recent
        ),

        "recent_p95_ms": percentile(
            recent
        ),

        "recent_samples": len(
            recent
        ),

        "slow_inflight_workers": (
            slow_inflight
        ),

        "requests_per_second": round(
            snapshot["requests"] / elapsed,
            3,
        ),

        "cells_per_second": round(
            snapshot["non_null_cells"]
            / elapsed,
            3,
        ),

        "chunks_per_second": round(
            chunk_rate,
            3,
        ),

        "eta_seconds": (
            round(
                remaining / chunk_rate,
                2,
            )

            if completed >= 10
            and chunk_rate > 0

            else None
        ),

        "completion_percent": (
            round(
                completed / total * 100,
                3,
            )
            if total
            else 100.0
        ),
    })

    return snapshot


# ============================================================
# BOTTLENECK DIAGNOSIS
# ============================================================

def identify_bottleneck(metrics):

    timing = metrics.get(
        "timing",
        {},
    )

    response_wait = timing.get(
        "response_wait_estimate_ms"
    )

    if response_wait is None:

        response_wait = timing.get(
            "ttfb_ms"
        )

    if response_wait is None:

        response_wait = timing.get(
            "header_wait_ms"
        )

    if response_wait is None:

        response_wait = timing.get(
            "first_body_ms"
        )

    candidates = {
        "RESPONSE_WAIT": response_wait,

        "RESPONSE_RECEIVE": timing.get(
            "receive_ms"
        ),

        "DNS_LOOKUP": timing.get(
            "dns_ms"
        ),

        "TCP_CONNECT": timing.get(
            "connect_ms"
        ),

        "TLS_HANDSHAKE": timing.get(
            "tls_ms"
        ),

        "FILE_WRITE": timing.get(
            "file_write_ms"
        ),

        "XML_PARSE": timing.get(
            "parse_ms"
        ),

        "TOKEN_REFRESH": timing.get(
            "token_refresh_ms"
        ),
    }

    candidates = {
        key: value

        for key, value in candidates.items()

        if value is not None
    }

    if not candidates:

        return {
            "name": "UNKNOWN",

            "message": (
                "Insufficient timing samples"
            ),
        }

    name, duration = max(
        candidates.items(),

        key=lambda item: item[1],
    )

    execution = timing.get(
        "execute_ms"
    )

    return {
        "name": name,

        "average_ms": round(
            duration,
            3,
        ),

        "relative_execute_percent": (
            round(
                duration / execution * 100,
                2,
            )
            if execution
            else None
        ),

        "scheduler_rate_wait_ms": timing.get(
            "rate_wait_ms"
        ),

        "message": (
            "Response waiting dominates. "
            "OLAP execution, upstream queuing and "
            "network transit remain possible causes."

            if name == "RESPONSE_WAIT"

            else f"Largest measured stage: {name}"
        ),

        "note": (
            "Timing stages can overlap. "
            "Response waiting is not pure server CPU time. "
            "Excel and data CSV export are measured separately."
        ),
    }


# ============================================================
# THREAD-LOCAL PERSISTENT HTTP SESSIONS
# ============================================================

def initialize_worker(
    auth,
    worker_ids,
    sessions,
    session_lock,
):

    worker_id = worker_ids.popleft()

    session = create_session(
        auth.cookie
    )

    THREAD_LOCAL.worker_id = worker_id

    THREAD_LOCAL.session = session

    THREAD_LOCAL.last_request_at = 0.0

    with session_lock:

        sessions.append(
            session
        )


# ============================================================
# REQUEST WORKER
# ============================================================

def request_worker(
    query,
    axes,
    profile,
    task,
    queued_at,
    auth,
    control,
    state,
    stop_event,
    profile_dir,
):

    worker_id = THREAD_LOCAL.worker_id

    session = THREAD_LOCAL.session

    records = []

    rejected_token = None

    result = None

    fatal_error = None
    stopped = False

    task_started = time.perf_counter()

    queue_wait_ms = round(
        (
            task_started
            - queued_at
        ) * 1000,
        3,
    )

    try:

        payload = build_soap(
            build_mdx(
                query,
                axes,
                task,
                profile.strategy,
            )
        )

        destination = (
            profile_dir
            / "responses"
            / f"{task.request_id}.xml"
        )

        for attempt in range(
            1,
            MAX_ATTEMPTS + 1,
        ):

            check_stop(stop_event)

            state.update_worker(
                worker_id,

                status="RUNNING",
                phase="AUTHENTICATING",

                x_offset=task.x_offset,
                y_offset=task.y_offset,

                request_started=None,

                last_event=(
                    f"Preparing X={task.x_offset} "
                    f"Y={task.y_offset}"
                ),
            )

            token, token_lock_ms, refresh_ms = (
                auth.get_token(
                    session,

                    rejected_token,

                    stop_event,
                )
            )

            rejected_token = None

            delay = max(
                0.0,

                THREAD_LOCAL.last_request_at
                + WORKER_REQUEST_DELAY

                - time.monotonic(),
            )

            delay_started = (
                time.perf_counter()
            )

            if delay:

                state.update_worker(
                    worker_id,
                    phase="DELAY",
                )

                interruptible_sleep(
                    delay,
                    stop_event,
                )

            delay_ms = elapsed_ms(
                delay_started
            )

            state.update_worker(
                worker_id,

                phase=(
                    "COOLDOWN"
                    if control.remaining() > 0
                    else "RATE_WAIT"
                ),
            )

            rate_wait_ms = control.wait_slot(
                stop_event
            )

            check_stop(stop_event)

            started = time.perf_counter()

            THREAD_LOCAL.last_request_at = (
                time.monotonic()
            )

            record = {
                "profile": profile.name,

                "query_strategy": profile.strategy,

                "worker_id": worker_id,

                "request_id": task.request_id,

                "attempt": attempt,

                "x_offset": task.x_offset,
                "y_offset": task.y_offset,

                "x_chunk": task.x_chunk,
                "y_chunk": task.y_chunk,

                "started_at": utc_now(),

                "queue_wait_ms": (
                    queue_wait_ms
                    if attempt == 1
                    else 0.0
                ),

                "rate_wait_ms": round(
                    rate_wait_ms + delay_ms,
                    3,
                ),

                "token_lock_wait_ms": (
                    token_lock_ms
                ),

                "token_refresh_ms": (
                    refresh_ms
                ),
            }

            state.update_worker(
                worker_id,

                status="RUNNING",
                phase="REQUESTING",

                request_started=(
                    time.monotonic()
                ),

                last_event=(
                    f"REQUEST "
                    f"X={task.x_offset} "
                    f"Y={task.y_offset} "
                    f"SIZE={task.x_chunk}/{task.y_chunk}"
                ),
            )

            retry = False

            try:

                response = execute_request(
                    session,

                    token,

                    payload,

                    destination,

                    stop_event=stop_event,
                )

                record.update(
                    response
                )

                http_status = response[
                    "http_status"
                ]

                state.update_worker(
                    worker_id,

                    last_http=http_status,

                    request_started=None,
                )

                if http_status == 200:

                    state.update_worker(
                        worker_id,

                        phase="PARSING",
                    )

                    parsed = parse_response(
                        destination,

                        query,

                        axes,

                        task,
                    )

                    rows = parsed.pop(
                        "rows"
                    )

                    record.update(
                        parsed
                    )

                    record["status"] = "OK"

                    result = {
                        **record,

                        "rows": rows,
                    }

                elif http_status in (401, 403):

                    record["status"] = "ERROR"

                    record["error"] = (
                        f"Authentication HTTP {http_status}"
                    )

                    if attempt < MAX_ATTEMPTS:

                        rejected_token = token

                        retry = True

                    else:

                        fatal_error = record[
                            "error"
                        ]

                elif http_status in (429, 503):

                    record["status"] = "ERROR"

                    record["error"] = (
                        f"Server HTTP {http_status}"
                    )

                    control.set_cooldown(
                        retry_after_seconds(
                            response.get(
                                "retry_after"
                            ),
                            30,
                        ),

                        record["error"],

                        halt=True,
                    )

                    fatal_error = (
                        record["error"]
                        + " | Server backoff required"
                    )

                elif http_status in (
                    408,
                    425,
                    500,
                    502,
                    504,
                ):

                    record["status"] = "ERROR"

                    record["error"] = (
                        f"Transient HTTP {http_status}"
                    )

                    retry = (
                        attempt < MAX_ATTEMPTS
                    )

                    if not retry:

                        fatal_error = record[
                            "error"
                        ]

                else:

                    record["status"] = "ERROR"

                    record["error"] = (
                        f"HTTP {http_status}"
                    )

                    fatal_error = record[
                        "error"
                    ]

            except StopRequested:

                record["status"] = "STOPPED"

                record["error"] = "Stopped"

                stopped = True

            except RequestException as exc:

                record["status"] = "ERROR"

                record["error"] = safe_error(
                    exc
                )

                retry = (
                    attempt < MAX_ATTEMPTS
                )

                if not retry:

                    fatal_error = (
                        record["error"]
                    )

            except (
                ResponseError,
                ResponseTooLarge,
            ) as exc:

                record["status"] = "ERROR"

                record["error"] = safe_error(
                    exc
                )

                # Do not blindly repeat expensive
                # or structurally invalid queries.

                fatal_error = record[
                    "error"
                ]

            except Exception as exc:

                record["status"] = "ERROR"

                record["error"] = safe_error(
                    exc
                )

                fatal_error = record[
                    "error"
                ]

            finally:

                record.setdefault(
                    "execute_ms",

                    elapsed_ms(
                        started
                    ),
                )

                record["completed_at"] = utc_now()

                record["task_total_ms"] = round(
                    (
                        time.perf_counter()
                        - queued_at
                    ) * 1000,
                    3,
                )

                state.record_request(
                    record
                )

                records.append(
                    dict(record)
                )

            if result is not None:

                duration = result.get(
                    "execute_ms",
                    0.0,
                )

                prefix = (
                    "SLOW"
                    if duration >= SLOW_REQUEST_MS
                    else "OK"
                )

                state.update_worker(
                    worker_id,

                    phase="IDLE",

                    last_event=(
                        f"{prefix} HTTP 200 | "
                        f"{format_ms(duration)} | "
                        f"TTFB "
                        f"{format_ms(result.get('ttfb_ms'))} | "
                        f"{result['chunk_kind']}"
                    ),
                )

                break

            if stopped or fatal_error:

                break

            if not retry:

                break

            state.update_worker(
                worker_id,

                status="RETRYING",

                phase="RETRYING",

                last_event=(
                    f"Retry {attempt + 1}/"
                    f"{MAX_ATTEMPTS}"
                ),
            )

            interruptible_sleep(
                min(
                    10,
                    2 ** attempt,
                ),
                stop_event,
            )

    except StopRequested:

        stopped = True

    except (
        AuthenticationError,
        ServiceBusyError,
    ) as exc:

        fatal_error = safe_error(
            exc
        )

    except Exception as exc:

        fatal_error = safe_error(
            exc
        )

    if fatal_error:

        state.set_fatal(
            fatal_error
        )

        stop_event.set()

    if stopped:

        state.update_worker(
            worker_id,

            status="STOPPED",

            phase="IDLE",

            request_started=None,

            last_event="STOPPED",
        )

    else:

        state.finish_task(
            worker_id,
            result,
        )

    return {
        "worker_id": worker_id,

        "task": task,

        "records": records,

        "result": result,

        "fatal_error": fatal_error,

        "stopped": stopped,
    }


# ============================================================
# BENCHMARK REGIONS
# ============================================================

def build_benchmark_regions(x_total):

    size = min(
        BENCHMARK_REGION_SIZE,
        x_total,
    )

    candidates = [
        (
            0,
            size,
        ),

        (
            max(
                0,
                x_total // 2
                - size // 2,
            ),
            size,
        ),

        (
            max(
                0,
                x_total - size,
            ),
            size,
        ),
    ]

    intervals = sorted(
        (
            start,
            min(
                start + count,
                x_total,
            ),
        )

        for start, count in candidates
    )

    merged = []

    for start, end in intervals:

        if merged and start <= merged[-1][1]:

            merged[-1][1] = max(
                merged[-1][1],
                end,
            )

        else:

            merged.append([
                start,
                end,
            ])

    return [
        (
            start,
            end - start,
        )
        for start, end in merged
    ]


def build_benchmark_profiles(
    max_workers,
    axes,
):

    all_years = axes.y_count

    profiles = [
        Profile(
            "01_REFERENCE_SUBSET",

            "subset",

            1,

            5,

            all_years,
        ),

        Profile(
            "02_EXPLICIT_SMALL",

            "explicit",

            1,

            5,

            all_years,
        ),

        Profile(
            "03_EXPLICIT_BALANCED",

            "explicit",

            1,

            10,

            all_years,
        ),

        Profile(
            "04_EXPLICIT_MEDIUM",

            "explicit",

            1,

            20,

            all_years,
        ),
    ]

    if max_workers >= 2:

        profiles.append(
            Profile(
                "05_EXPLICIT_PARALLEL_2",

                "explicit",

                2,

                10,

                all_years,
            )
        )

    if max_workers > 2:

        profiles.append(
            Profile(
                "06_EXPLICIT_PARALLEL_MAX",

                "explicit",

                max_workers,

                10,

                all_years,
            )
        )

    return profiles


def generate_tasks(
    query,
    axes,
    profile,
    regions,
):

    for region_start, region_size in regions:

        region_end = (
            region_start
            + region_size
        )

        for x_offset in range(
            region_start,
            region_end,
            profile.x_chunk,
        ):

            x_chunk = min(
                profile.x_chunk,

                region_end
                - x_offset,
            )

            for y_offset in range(
                0,

                axes.y_count,

                profile.y_chunk,
            ):

                y_chunk = min(
                    profile.y_chunk,

                    axes.y_count
                    - y_offset,
                )

                request_id = (
                    f"{query.query_hash}"
                    f"_x{x_offset:06d}"
                    f"_y{y_offset:06d}"
                    f"_xs{x_chunk:04d}"
                    f"_ys{y_chunk:04d}"
                )

                yield RequestTask(
                    request_id=request_id,

                    x_offset=x_offset,

                    y_offset=y_offset,

                    x_chunk=x_chunk,

                    y_chunk=y_chunk,
                )


def count_tasks(
    axes,
    profile,
    regions,
):

    return sum(
        math.ceil(
            count / profile.x_chunk
        )

        * math.ceil(
            axes.y_count
            / profile.y_chunk
        )

        for _, count in regions
    )


# ============================================================
# RICH UI UTILITIES
# ============================================================

def latency_color(value):

    if value is None:

        return "dim"

    if value >= VERY_SLOW_REQUEST_MS:

        return "bright_red"

    if value >= SLOW_REQUEST_MS:

        return "yellow"

    if value < 1000:

        return "green"

    return "white"


def status_color(status):

    return {
        "STARTING": "dim",

        "RUNNING": "bright_cyan",

        "RETRYING": "yellow",

        "PAUSED": "yellow",

        "COMPLETED": "green",

        "FAILED": "red",

        "STOPPED": "dim",
    }.get(
        status,
        "white",
    )


def make_panel(content, title, height=None):

    return Panel(
        content,

        title=title,

        border_style="cyan",

        padding=(0, 1),

        height=height,
    )


def kv_table(values):

    table = Table.grid(
        expand=True
    )

    table.add_column(
        style="dim",

        ratio=1,
    )

    table.add_column(
        ratio=1,

        justify="right",
    )

    for name, value in values:

        table.add_row(
            Text(str(name)),

            (
                value
                if isinstance(
                    value,
                    Text,
                )

                else Text(str(value))
            ),
        )

    return table


def paired_panels(
    first,
    first_title,
    second,
    second_title,
    height,
):

    if console.size.width < 100:

        return Group(
            make_panel(
                first,
                first_title,
            ),

            make_panel(
                second,
                second_title,
            ),
        )

    table = Table.grid(
        expand=True
    )

    table.add_column(
        ratio=1
    )

    table.add_column(
        ratio=1
    )

    table.add_row(
        make_panel(
            first,
            first_title,
            height,
        ),

        make_panel(
            second,
            second_title,
            height,
        ),
    )

    return table


# ============================================================
# LIVE TERMINAL MONITOR
# ============================================================

def render_monitor(
    profile,
    axes,
    state,
    control,
    mode,
    benchmark_results,
):

    metrics = aggregate_metrics(
        state.snapshot()
    )

    width = console.size.width
    height = console.size.height

    timing = metrics[
        "timing"
    ]

    completed = metrics[
        "completed_tasks"
    ]

    total = metrics[
        "total_tasks"
    ]

    workers = metrics[
        "workers"
    ]

    header = Panel(
        Text(
            APP_NAME

            + f" | v{APP_VERSION}"

            + "\n"

            + "MV/FY | XMLA / MDX | NATIVE CURL CALLBACK",

            justify="center",

            style="bold bright_cyan",
        ),

        border_style="cyan",
    )

    overview = kv_table([
        (
            "Mode",

            mode.upper(),
        ),

        (
            "Profile",

            profile.name,
        ),

        (
            "MDX strategy",

            profile.strategy.upper(),
        ),

        (
            "Parallel limit",

            f"{metrics['active_limit']}/"
            f"{profile.workers}",
        ),

        (
            "Elapsed",

            format_duration(
                metrics["elapsed"]
            ),
        ),

        (
            "ETA",

            format_duration(
                metrics["eta_seconds"]
            )

            if metrics["eta_seconds"] is not None
            else "—",
        ),
    ])

    configuration = kv_table([
        (
            "Cube",

            "MV / FY",
        ),

        (
            "X1",

            X_CAPTIONS[0],
        ),

        (
            "X2",

            X_CAPTIONS[1],
        ),

        (
            "Y1",

            Y_CAPTIONS[0],
        ),

        (
            "Measures",

            len(MEASURE_CAPTIONS),
        ),

        (
            "Chunk X/Y",

            f"{profile.x_chunk}/"
            f"{profile.y_chunk}",
        ),

        (
            "Axis X/Y",

            f"{axes.x_count:,}/"
            f"{axes.y_count:,}",
        ),
    ])

    top = paired_panels(
        overview,

        "01 / ENGINE OVERVIEW",

        configuration,

        "02 / FIXED QUERY",

        height=9,
    )

    progress = make_panel(
        Group(
            Text.assemble(
                (
                    f"{completed:,}/{total:,} COMPLETED",

                    "bold green",
                ),

                " | ",

                (
                    f"{metrics['failed_tasks']:,} FAILED",

                    "red"
                    if metrics["failed_tasks"]
                    else "dim",
                ),

                " | ",

                (
                    f"{metrics['completion_percent']:.2f}%",

                    "cyan",
                ),
            ),

            Bar(
                size=max(
                    1,
                    total,
                ),

                begin=0,

                end=completed,

                color="cyan",

                bgcolor="grey23",
            ),

            Text(
                f"Data {metrics['data_chunks']:,}"

                f" | No cells {metrics['no_cell_chunks']:,}"

                f" | Empty {metrics['empty_pages']:,}"

                " | Excel export: AFTER EXTRACTION",

                style="dim",
            ),
        ),

        "03 / EXTRACTION PROGRESS",
    )

    performance = kv_table([
        (
            "Requests",

            metrics["requests"],
        ),

        (
            "Requests/sec",

            f"{metrics['requests_per_second']:.2f}",
        ),

        (
            "Cells/sec",

            Text(
                f"{metrics['cells_per_second']:.2f}",

                style="bold green",
            ),
        ),

        (
            "Average",

            Text(
                format_ms(
                    metrics["average_ms"]
                ),

                style=latency_color(
                    metrics["average_ms"]
                ),
            ),
        ),

        (
            "P95",

            Text(
                format_ms(
                    metrics["p95_ms"]
                ),

                style=latency_color(
                    metrics["p95_ms"]
                ),
            ),
        ),

        (
            "Recent average",

            format_ms(
                metrics["recent_average_ms"]
            ),
        ),

        (
            "Recent P95",

            format_ms(
                metrics["recent_p95_ms"]
            ),
        ),

        (
            "Slow requests",

            metrics["slow_requests"],
        ),

        (
            "Non-null cells",

            f"{metrics['non_null_cells']:,}",
        ),

        (
            "Saved XML",

            format_bytes(
                metrics["response_bytes"]
            ),
        ),

        (
            "Downloaded",

            format_bytes(
                metrics["download_bytes"]
            ),
        ),

        (
            "Connection reuse",

            (
                f"{metrics['connection_reuse_percent']:.1f}%"

                if metrics["connection_reuse_percent"]
                is not None

                else "—"
            ),
        ),
    ])

    timing_table = kv_table([
        (
            label,

            Text(
                format_ms(
                    timing.get(key)
                ),

                style=latency_color(
                    timing.get(key)
                ),
            ),
        )

        for label, key in (
            ("Queue wait", "queue_wait_ms"),

            ("Rate wait", "rate_wait_ms"),

            ("Token refresh", "token_refresh_ms"),

            ("DNS", "dns_ms"),

            ("TCP connect", "connect_ms"),

            ("TLS", "tls_ms"),

            ("TTFB", "ttfb_ms"),

            (
                "Response wait",
                "response_wait_estimate_ms",
            ),

            ("Receive", "receive_ms"),

            ("XML write", "file_write_ms"),

            ("Finalize XML", "finalize_ms"),

            ("XML parse", "parse_ms"),
        )
    ])

    metric_panels = paired_panels(
        performance,

        "04 / PERFORMANCE",

        timing_table,

        "05 / TIMING BREAKDOWN",

        height=14,
    )

    worker_blocks = []

    compact = (
        height < 45 + len(workers) * 4
        or width < 100
    )

    for worker in workers:

        worker_id = worker[
            "worker_id"
        ]

        status = worker[
            "status"
        ]

        phase = worker[
            "phase"
        ]

        last_ms = worker[
            "last_request_ms"
        ]

        current_ms = worker[
            "current_request_ms"
        ]

        color = status_color(
            status
        )

        bar = Bar(
            size=1000,

            begin=0,

            end=int(
                worker["relative_progress"]
                * 1000
            ),

            color=color,

            bgcolor="grey23",
        )

        offset = (
            f"{worker['x_offset']}/"
            f"{worker['y_offset']}"

            if worker["x_offset"] is not None

            else "—"
        )

        if compact:

            table = Table.grid(
                expand=True
            )

            table.add_column(
                width=5
            )

            table.add_column(
                width=14
            )

            table.add_column(
                ratio=1
            )

            table.add_column(
                justify="right"
            )

            table.add_row(
                Text(
                    f"W{worker_id:02d}",
                    style="bold",
                ),

                Text(
                    phase,
                    style=color,
                ),

                bar,

                Text(
                    f"{worker['completed']:,} OK"

                    f" | {worker['failed']} ERR"

                    f" | Last {format_ms(last_ms)}"

                    f" | Now {format_ms(current_ms)}",

                    style=latency_color(
                        current_ms
                        if current_ms is not None
                        else last_ms
                    ),
                ),
            )

            worker_blocks.append(
                table
            )

        else:

            row = Table.grid(
                expand=True
            )

            row.add_column(
                ratio=1
            )

            row.add_column(
                justify="right"
            )

            row.add_row(
                Text.assemble(
                    (
                        f"W{worker_id:02d}  ",
                        "bold white",
                    ),

                    (
                        status,
                        color,
                    ),
                ),

                Text(
                    f"Last {format_ms(last_ms)}",

                    style=latency_color(
                        last_ms
                    ),
                ),
            )

            details = Table.grid(
                expand=True
            )

            details.add_column(
                ratio=1
            )

            details.add_column(
                justify="right"
            )

            details.add_row(
                Text(
                    f"Offset {offset} | {phase}",

                    style="dim",
                ),

                Text(
                    f"{worker['completed']:,} chunks"

                    f" | {worker['failed']} failed",
                ),
            )

            extra = Text.assemble(
                (
                    f"HTTP {worker['last_http'] or '—'}",

                    "green"

                    if worker["last_http"] == 200

                    else "dim",
                ),

                (
                    f" | TTFB "
                    f"{format_ms(worker['last_ttfb_ms'])}"
                ),

                (
                    f" | Current "
                    f"{format_ms(current_ms)}",

                    latency_color(
                        current_ms
                    ),
                ),

                f" | Cells {worker['cells']:,}",
            )

            worker_blocks.extend([
                row,

                bar,

                details,

                extra,
            ])

    worker_blocks.append(
        Text(
            "Relative completed chunks, "
            "not worker completion percentage.",

            style="dim",
        )
    )

    worker_panel = make_panel(
        Group(*worker_blocks),

        "06 / PARALLEL WORKERS",
    )

    diagnosis = identify_bottleneck(
        metrics
    )

    alerts = []

    if metrics["slow_inflight_workers"]:

        alerts.append(
            Text(
                "SLOW IN-FLIGHT: "

                f"{metrics['slow_inflight_workers']} workers",

                style="yellow",
            )
        )

    if diagnosis["name"] != "UNKNOWN":

        alerts.append(
            Text(
                f"BOTTLENECK: {diagnosis['name']}"

                f" | {format_ms(diagnosis.get('average_ms'))}",

                style=(
                    "yellow"

                    if diagnosis.get(
                        "average_ms",
                        0,
                    ) >= SLOW_REQUEST_MS

                    else "cyan"
                ),
            )
        )

    if metrics["active_limit"] < profile.workers:

        alerts.append(
            Text(
                "CONCURRENCY REDUCED: "

                f"{profile.workers}"

                f" → {metrics['active_limit']}",

                style="yellow",
            )
        )

    if control.remaining() > 0:

        alerts.append(
            Text(
                "SERVER BACKOFF: "

                f"{control.remaining():.1f}s",

                style="yellow",
            )
        )

    if metrics.get("fatal_error"):

        alerts.append(
            Text(
                metrics["fatal_error"],

                style="red",
            )
        )

    if not alerts:

        alerts.append(
            Text(
                "No active performance warnings",

                style="green",
            )
        )

    alert_panel = make_panel(
        Group(*alerts),

        "07 / PERFORMANCE ANALYSIS",
    )

    sections = [
        header,

        top,

        progress,

        metric_panels,

        worker_panel,

        alert_panel,
    ]

    if benchmark_results:

        table = Table(
            box=box.SIMPLE,

            expand=True,
        )

        for column in (
            "Profile",

            "W",

            "Chunk",

            "Elapsed",

            "P95",

            "Cells/s",

            "Match",
        ):

            table.add_column(
                column
            )

        for result in benchmark_results[-6:]:

            equivalent = result.get(
                "equivalent"
            )

            table.add_row(
                result["profile"],

                str(
                    result["workers_count"]
                ),

                f"{result['x_chunk']}/"
                f"{result['y_chunk']}",

                format_duration(
                    result["elapsed"]
                ),

                format_ms(
                    result.get("p95_ms")
                ),

                f"{result['cells_per_second']:.2f}",

                Text(
                    "YES"

                    if equivalent is True

                    else "NO"

                    if equivalent is False

                    else "—",

                    style=(
                        "green"

                        if equivalent is True

                        else "red"

                        if equivalent is False

                        else "dim"
                    ),
                ),
            )

        if height >= 50:

            sections.append(
                make_panel(
                    table,

                    "08 / BENCHMARK RESULTS",
                )
            )

    if height >= 42:

        events = [
            Text.assemble(
                (
                    event["time"] + " ",
                    "dim",
                ),

                (
                    f"W{event['worker']:02d} ",
                    "cyan",
                ),

                event["message"],
            )

            for event in metrics["events"][-3:]
        ]

        sections.append(
            make_panel(
                Group(
                    *events
                ),

                "LATEST ACTIVITY",
            )
        )

    sections.append(
        Text(
            "LOCAL TIME | REFRESH 1s | CTRL+C TO STOP",

            style="dim",
        )
    )

    return Group(
        *sections
    )


# ============================================================
# DATASET VALIDATION
# ============================================================

def dataset_digest(rows):

    digest = hashlib.sha256()

    for key in sorted(rows):

        encoded = json.dumps(
            [
                key,

                rows[key],
            ],

            ensure_ascii=False,

            separators=(",", ":"),
        ).encode("utf-8")

        digest.update(encoded)

        digest.update(b"\n")

    return digest.hexdigest()


def expected_dataset_keys(
    axes,
    regions,
):

    return {
        tuple_key(x_members)
        + tuple_key(y_members)

        for start, count in regions

        for x_members in axes.x[
            start:
            start + count
        ]

        for y_members in axes.y
    }


# ============================================================
# ADAPTIVE CONCURRENCY
# ============================================================

def evaluate_concurrency(
    state,
    active_limit,
    last_adjustment,
):

    if active_limit <= 1:

        return (
            active_limit,
            last_adjustment,
        )

    if (
        time.monotonic()
        - last_adjustment
        < 15
    ):

        return (
            active_limit,
            last_adjustment,
        )

    snapshot = state.snapshot()

    recent = [
        item

        for item in snapshot["recent"][-6:]

        if item["execute_ms"] is not None
    ]

    if len(recent) < 3:

        return (
            active_limit,
            last_adjustment,
        )

    slow = sum(
        item["execute_ms"]
        >= VERY_SLOW_REQUEST_MS

        for item in recent
    )

    errors = sum(
        item["status"] == "ERROR"

        for item in recent
    )

    if slow >= 3 or errors >= 2:

        active_limit = max(
            1,

            active_limit // 2,
        )

        state.set_active_limit(
            active_limit
        )

        return (
            active_limit,

            time.monotonic(),
        )

    return (
        active_limit,
        last_adjustment,
    )


# ============================================================
# BOUNDED DYNAMIC REQUEST SCHEDULER
# ============================================================

def run_scheduler(
    query,
    axes,
    profile,
    regions,
    auth,
    control,
    run_dir,
    csv_writer,
    csv_handle,
    *,
    mode,
    benchmark_results,
):

    task_total = count_tasks(
        axes,

        profile,

        regions,
    )

    if task_total > MAX_TASKS_PER_RUN:

        raise EngineError(
            "Planned tasks exceeded local limit"
        )

    profile_dir = (
        run_dir
        / profile.name
    )

    profile_dir.mkdir(
        parents=True,

        exist_ok=True,
    )

    state = EngineState(
        profile,

        task_total,
    )

    stop_event = threading.Event()

    worker_ids = deque(
        range(
            1,

            profile.workers + 1,
        )
    )

    sessions = []

    session_lock = threading.Lock()

    task_iterator = iter(
        generate_tasks(
            query,

            axes,

            profile,

            regions,
        )
    )

    rows = {}

    snapshots = set()

    missing_snapshot_chunks = 0

    duplicate_keys = 0

    exhausted = False

    pending = {}

    active_limit = profile.workers

    last_adjustment = time.monotonic()

    next_refresh = 0.0

    interrupted = False

    profile_deadline_reached = False

    def consume_future(future):

        nonlocal duplicate_keys
        nonlocal missing_snapshot_chunks

        try:

            completed = future.result()

        except Exception as exc:

            state.set_fatal(
                safe_error(exc)
            )

            stop_event.set()

            return

        for record in completed[
            "records"
        ]:

            csv_writer.writerow({
                name: record.get(name)

                for name in REQUEST_CSV_FIELDS
            })

        result = completed.get(
            "result"
        )

        if result is None:

            if completed.get(
                "fatal_error"
            ):

                state.set_fatal(
                    completed["fatal_error"]
                )

                stop_event.set()

            return

        snapshot = result.get(
            "last_data_update"
        )

        if snapshot:

            snapshots.add(snapshot)

        else:

            missing_snapshot_chunks += 1

        if len(snapshots) > 1:

            state.set_fatal(
                "Cube snapshot changed"
            )

            stop_event.set()

        for key, values in result[
            "rows"
        ].items():

            if key in rows:

                duplicate_keys += 1

                state.set_fatal(
                    "Duplicate dataset row detected"
                )

                stop_event.set()

            else:

                rows[key] = values

    executor = ThreadPoolExecutor(
        max_workers=profile.workers,

        thread_name_prefix="OIC",

        initializer=initialize_worker,

        initargs=(
            auth,

            worker_ids,

            sessions,

            session_lock,
        ),
    )

    try:

        with Live(
            console=console,

            auto_refresh=False,

            refresh_per_second=1,

            screen=True,

            vertical_overflow="crop",
        ) as live:

            while True:

                if STOP_EVENT.is_set():

                    stop_event.set()

                if control.service_busy:

                    stop_event.set()

                if (
                    mode == "benchmark"

                    and time.monotonic()
                    - state.started

                    >= BENCHMARK_MAX_PROFILE_SECONDS
                ):

                    profile_deadline_reached = True

                    stop_event.set()

                if mode == "run":

                    (
                        active_limit,

                        last_adjustment,

                    ) = evaluate_concurrency(
                        state,

                        active_limit,

                        last_adjustment,
                    )

                while (
                    not exhausted

                    and not stop_event.is_set()

                    and len(pending) < active_limit
                ):

                    try:

                        task = next(
                            task_iterator
                        )

                    except StopIteration:

                        exhausted = True

                        break

                    future = executor.submit(
                        request_worker,

                        query,

                        axes,

                        profile,

                        task,

                        time.perf_counter(),

                        auth,

                        control,

                        state,

                        stop_event,

                        profile_dir,
                    )

                    pending[future] = task

                if pending:

                    completed, _ = wait(
                        tuple(pending),

                        timeout=0.2,

                        return_when=FIRST_COMPLETED,
                    )

                    for future in completed:

                        pending.pop(
                            future,
                            None,
                        )

                        consume_future(
                            future
                        )

                    if completed:

                        csv_handle.flush()

                now = time.monotonic()

                if now >= next_refresh:

                    live.update(
                        render_monitor(
                            profile,

                            axes,

                            state,

                            control,

                            mode,

                            benchmark_results,
                        ),

                        refresh=True,
                    )

                    next_refresh = (
                        now
                        + REFRESH_INTERVAL
                    )

                if (
                    exhausted
                    or stop_event.is_set()
                ) and not pending:

                    break

    except KeyboardInterrupt:

        interrupted = True

        STOP_EVENT.set()

        stop_event.set()

    except Exception as exc:

        state.set_fatal(
            safe_error(exc)
        )

        stop_event.set()

    finally:

        if STOP_EVENT.is_set():

            stop_event.set()

        executor.shutdown(
            wait=True,

            cancel_futures=True,
        )

        for future in list(pending):

            if (
                future.done()
                and not future.cancelled()
            ):

                consume_future(
                    future
                )

        pending.clear()

        csv_handle.flush()

        for session in sessions:

            try:

                session.close()

            except Exception:

                pass

    # Freeze benchmark time before CSV/XLSX data export.
    extraction_metrics = aggregate_metrics(
        state.snapshot()
    )

    expected_keys = expected_dataset_keys(
        axes,

        regions,
    )

    complete_coverage = (
        set(rows) == expected_keys
    )

    if (
        STOP_EVENT.is_set()
        or interrupted
    ):

        status = "STOPPED"

    elif profile_deadline_reached:

        status = "PARTIAL"

    elif (
        extraction_metrics["fatal_error"]

        or extraction_metrics["failed_tasks"]

        or extraction_metrics["completed_tasks"]
        != task_total

        or not complete_coverage

        or duplicate_keys

        or len(snapshots) > 1
    ):

        status = (
            "PARTIAL"

            if rows

            else "FAILED"
        )

    else:

        status = "COMPLETED"

    for worker in extraction_metrics[
        "workers"
    ]:

        worker["request_started"] = None

        worker["current_request_ms"] = None

        worker["phase"] = "FINISHED"

        worker["status"] = (
            "FAILED"

            if worker["failed"]

            else "STOPPED"

            if status != "COMPLETED"

            else "COMPLETED"
        )

    report = {
        **extraction_metrics,

        "status": status,

        "profile": profile.name,

        "strategy": profile.strategy,

        "engine": "curl_cffi_callback",

        "workers_count": profile.workers,

        "x_chunk": profile.x_chunk,

        "y_chunk": profile.y_chunk,

        "query_hash": query.query_hash,

        "axis_digest": axes.digest,

        "regions": [
            list(region)

            for region in regions
        ],

        "expected_rows": len(
            expected_keys
        ),

        "returned_rows": len(
            rows
        ),

        "missing_rows": len(
            expected_keys - set(rows)
        ),

        "unexpected_rows": len(
            set(rows) - expected_keys
        ),

        "duplicate_keys": duplicate_keys,

        "complete_coverage": (
            complete_coverage
        ),

        "expected_coordinates": (
            len(expected_keys)
            * len(query.measures)
        ),

        "returned_coordinates": (
            len(rows)
            * len(query.measures)
        ),

        "last_data_updates": sorted(
            snapshots
        ),

        "missing_snapshot_chunks": (
            missing_snapshot_chunks
        ),

        "snapshot_verified": (
            bool(axes.snapshot)

            and snapshots == {
                axes.snapshot
            }

            and missing_snapshot_chunks == 0
        ),

        "profile_deadline_reached": (
            profile_deadline_reached
        ),

        "dataset_digest": dataset_digest(
            rows
        ),

        "output_directory": str(
            profile_dir
        ),

        "completed_at": utc_now(),
    }

    report["bottleneck"] = (
        identify_bottleneck(report)
    )

    # Save the extraction report before export.
    # A failed export must not destroy benchmark timings.

    write_json(
        profile_dir / "summary.json",

        report,
    )

    return report, rows


# ============================================================
# CSV / EXCEL VALUE CONVERSION
# ============================================================

def csv_safe_text(value):

    value = str(
        value or ""
    )

    if value.lstrip().startswith(
        (
            "=",
            "+",
            "-",
            "@",
        )
    ):

        return "'" + value

    return value


def csv_measure_value(value):

    if value is None:

        return ""

    try:

        number = Decimal(value)

        if number.is_finite():

            return value

    except InvalidOperation:

        pass

    return csv_safe_text(
        value
    )


def excel_measure_value(value):

    if value is None:

        return None

    try:

        number = Decimal(value)

        if not number.is_finite():

            return value

        # Excel supports approximately 15 significant
        # digits. Preserve larger values as text.

        digits = len(
            number.as_tuple().digits
        )

        if digits > 15:

            return value

        if number == number.to_integral_value():

            if abs(number) < Decimal(
                "1000000000000000"
            ):

                return int(number)

            return value

        converted = float(number)

        if not math.isfinite(converted):

            return value

        if converted == 0 and number != 0:

            return value

        return converted

    except (
        InvalidOperation,
        ValueError,
        OverflowError,
    ):

        return value


def make_excel_text(
    sheet,
    value,
):

    cell = WriteOnlyCell(
        sheet,

        value=str(
            value or ""
        ),
    )

    # Prevent text beginning with "=" from being
    # interpreted as an Excel formula.

    cell.data_type = "s"

    return cell


# ============================================================
# EXPORT ROW GENERATOR
# ============================================================

def iter_export_rows(
    axes,
    rows,
):

    y_index = {
        tuple_key(members): members

        for members in axes.y
    }

    for x_members in axes.x:

        x_key = tuple_key(
            x_members
        )

        for y_key, y_members in y_index.items():

            key = (
                x_key
                + y_key
            )

            if key not in rows:

                continue

            captions = [
                member.caption
                or member.unique_name

                for member in (
                    *x_members,

                    *y_members,
                )
            ]

            identifiers = [
                member.unique_name

                for member in (
                    *x_members,

                    *y_members,
                )
            ]

            yield (
                captions,

                rows[key],

                identifiers,
            )


# ============================================================
# HIGH-PERFORMANCE DATA CSV EXPORT
# ============================================================

def export_data_csv(
    path,
    axes,
    rows,
):

    started = time.perf_counter()

    temporary = path.with_name(
        path.name + ".partial"
    )

    row_count = 0

    try:

        with temporary.open(
            "w",

            encoding="utf-8-sig",

            newline="",

            buffering=1024 * 1024,
        ) as handle:

            writer = csv.writer(
                handle
            )

            writer.writerow(
                DATA_HEADERS
            )

            for (
                captions,

                values,

                identifiers,

            ) in iter_export_rows(
                axes,

                rows,
            ):

                writer.writerow([
                    *[
                        csv_safe_text(value)

                        for value in captions
                    ],

                    *[
                        csv_measure_value(value)

                        for value in values
                    ],

                    *[
                        csv_safe_text(value)

                        for value in identifiers
                    ],
                ])

                row_count += 1

        temporary.replace(
            path
        )

        return {
            "status": "COMPLETED",

            "path": str(path),

            "rows": row_count,

            "bytes": path.stat().st_size,

            "duration_ms": elapsed_ms(
                started
            ),
        }

    finally:

        temporary.unlink(
            missing_ok=True
        )


# ============================================================
# HIGH-PERFORMANCE EXCEL EXPORT
# ============================================================

def export_data_excel(
    path,
    axes,
    rows,
    report,
):

    started = time.perf_counter()

    temporary = path.with_name(
        path.stem + ".partial.xlsx"
    )

    workbook = Workbook(
        write_only=True
    )

    sheet = workbook.create_sheet(
        title="MV_FY"
    )

    sheet.freeze_panes = "D2"

    sheet.column_dimensions["A"].width = 30
    sheet.column_dimensions["B"].width = 28
    sheet.column_dimensions["C"].width = 15

    for column in "DEFGHIJ":

        sheet.column_dimensions[
            column
        ].width = 25

    for column in "KLM":

        sheet.column_dimensions[
            column
        ].width = 42

    header_fill = PatternFill(
        fill_type="solid",

        fgColor="18354A",
    )

    header_font = Font(
        bold=True,

        color="FFFFFF",
    )

    headers = []

    for name in DATA_HEADERS:

        cell = WriteOnlyCell(
            sheet,

            value=name,
        )

        cell.font = header_font

        cell.fill = header_fill

        headers.append(
            cell
        )

    sheet.append(
        headers
    )

    row_count = 0

    try:

        for (
            captions,

            values,

            identifiers,

        ) in iter_export_rows(
            axes,

            rows,
        ):

            output_row = [
                make_excel_text(
                    sheet,

                    value,
                )

                for value in captions
            ]

            for value in values:

                converted = excel_measure_value(
                    value
                )

                if isinstance(
                    converted,
                    str,
                ):

                    output_row.append(
                        make_excel_text(
                            sheet,

                            converted,
                        )
                    )

                else:

                    output_row.append(
                        converted
                    )

            output_row.extend(
                make_excel_text(
                    sheet,

                    value,
                )

                for value in identifiers
            )

            sheet.append(
                output_row
            )

            row_count += 1

        metadata = workbook.create_sheet(
            title="METADATA"
        )

        metadata.column_dimensions[
            "A"
        ].width = 35

        metadata.column_dimensions[
            "B"
        ].width = 100

        metadata.append([
            "Field",

            "Value",
        ])

        metadata_rows = (
            (
                "Category",
                CATEGORY,
            ),

            (
                "Cube type",
                CUBE_TYPE,
            ),

            (
                "Catalog",
                CATALOG,
            ),

            (
                "Cube",
                CUBE,
            ),

            (
                "Profile",
                report["profile"],
            ),

            (
                "Extraction status",
                report["status"],
            ),

            (
                "Complete coverage",
                report["complete_coverage"],
            ),

            (
                "Expected rows",
                report["expected_rows"],
            ),

            (
                "Returned rows",
                row_count,
            ),

            (
                "Missing rows",
                report["missing_rows"],
            ),

            (
                "Dataset digest",
                report["dataset_digest"],
            ),

            (
                "Snapshot",
                axes.snapshot or "UNKNOWN",
            ),

            (
                "Measure units",
                "Raw OIC values; no unit conversion applied.",
            ),

            (
                "Monetary captions",
                "Original OIC captions preserved. "
                "Units must be independently verified.",
            ),

            (
                "Null values",
                "Preserved as blank cells.",
            ),

            (
                "All Members",
                "Preserved without automatic filtering.",
            ),

            (
                "Numeric precision",
                "Values exceeding Excel numeric precision "
                "are stored as text.",
            ),
        )

        for key, value in metadata_rows:

            metadata.append([
                make_excel_text(
                    metadata,

                    key,
                ),

                (
                    make_excel_text(
                        metadata,

                        value,
                    )

                    if isinstance(
                        value,
                        str,
                    )

                    else value
                ),
            ])

        workbook.save(
            temporary
        )

        temporary.replace(
            path
        )

        return {
            "status": "COMPLETED",

            "path": str(path),

            "rows": row_count,

            "bytes": path.stat().st_size,

            "duration_ms": elapsed_ms(
                started
            ),
        }

    finally:

        workbook.close()

        temporary.unlink(
            missing_ok=True
        )


# ============================================================
# SEPARATE EXPORT PIPELINE
# ============================================================

def export_profile_data(
    axes,
    rows,
    report,
):

    profile_dir = Path(
        report["output_directory"]
    )

    profile_dir.mkdir(
        parents=True,

        exist_ok=True,
    )

    started = time.perf_counter()

    prefix = (
        "data"

        if report["status"] == "COMPLETED"

        else "data_partial"
    )

    exports = {}

    console.print()

    console.print(
        Panel(
            Text(
                f"{report['profile']}"

                "\n"

                f"EXPORTING {len(rows):,} DATA ROWS"

                "\n"

                "HTTP extraction measurements frozen",

                justify="center",

                style="bold cyan",
            ),

            border_style="cyan",
        )
    )

    csv_path = (
        profile_dir
        / f"{prefix}.csv"
    )

    xlsx_path = (
        profile_dir
        / f"{prefix}.xlsx"
    )

    with console.status(
        "Writing data CSV...",

        spinner="dots",
    ):

        try:

            exports["csv"] = export_data_csv(
                csv_path,

                axes,

                rows,
            )

        except Exception as exc:

            exports["csv"] = {
                "status": "FAILED",

                "error": safe_error(exc),
            }

    with console.status(
        "Writing Excel workbook...",

        spinner="dots",
    ):

        try:

            exports["excel"] = export_data_excel(
                xlsx_path,

                axes,

                rows,

                report,
            )

        except Exception as exc:

            exports["excel"] = {
                "status": "FAILED",

                "error": safe_error(exc),
            }

    report["exports"] = exports

    report["csv_export_ms"] = (
        exports["csv"].get(
            "duration_ms"
        )
    )

    report["excel_export_ms"] = (
        exports["excel"].get(
            "duration_ms"
        )
    )

    report["export_total_ms"] = (
        elapsed_ms(started)
    )

    report["export_status"] = (
        "COMPLETED"

        if all(
            item["status"] == "COMPLETED"

            for item in exports.values()
        )

        else "PARTIAL"

        if any(
            item["status"] == "COMPLETED"

            for item in exports.values()
        )

        else "FAILED"
    )

    report["elapsed_with_export"] = round(
        report["elapsed"]

        + report["export_total_ms"]
        / 1000,

        3,
    )

    write_json(
        profile_dir / "summary.json",

        report,
    )

    console.print(
        make_panel(
            kv_table([
                (
                    "Extraction",

                    format_duration(
                        report["elapsed"]
                    ),
                ),

                (
                    "CSV export",

                    format_ms(
                        report["csv_export_ms"]
                    ),
                ),

                (
                    "Excel export",

                    format_ms(
                        report["excel_export_ms"]
                    ),
                ),

                (
                    "Export status",

                    report["export_status"],
                ),

                (
                    "Rows",

                    f"{len(rows):,}",
                ),
            ]),

            "EXPORT RESULT",
        )
    )

    return report


# ============================================================
# BENCHMARK PROFILE SELECTION
# ============================================================

def recommend_profile(reports):

    eligible = [
        report

        for report in reports

        if report.get("status") == "COMPLETED"

        and report.get("equivalent") is True

        and report.get("complete_coverage")

        and not report.get("duplicate_keys")

        and not report.get("failed_requests")

        and report.get("returned_rows", 0) > 0

        and report.get("non_null_cells", 0) > 0
    ]

    if not eligible:

        return None

    fastest = min(
        eligible,

        key=lambda report: (
            report["elapsed"],

            report.get("p95_ms")
            or float("inf"),

            report["workers_count"],
        ),
    )

    # Prefer lower concurrency when the measured
    # total extraction duration is within 10%
    # of the fastest eligible profile.

    comparable = [
        report

        for report in eligible

        if report["elapsed"]
        <= fastest["elapsed"] * 1.10
    ]

    best = min(
        comparable,

        key=lambda report: (
            report["workers_count"],

            report.get("p95_ms")
            or float("inf"),

            report["elapsed"],
        ),
    )

    baseline = next(
        (
            report

            for report in reports

            if report["profile"]
            == "01_REFERENCE_SUBSET"

            and report["status"]
            == "COMPLETED"
        ),

        None,
    )

    speedup = (
        baseline["elapsed"]
        / max(
            best["elapsed"],
            0.001,
        )

        if baseline

        else None
    )

    return {
        "profile": best["profile"],

        "strategy": best["strategy"],

        "engine": best["engine"],

        "workers": best["workers_count"],

        "x_chunk": best["x_chunk"],

        "y_chunk": best["y_chunk"],

        "sample_elapsed": best["elapsed"],

        "sample_cells_per_second": (
            best["cells_per_second"]
        ),

        "sample_p95_ms": best.get(
            "p95_ms"
        ),

        "sample_speedup": (
            round(speedup, 3)

            if speedup is not None

            else None
        ),

        "snapshot_verified": best.get(
            "snapshot_verified",
            False,
        ),

        "bottleneck": best.get(
            "bottleneck"
        ),

        "selection_policy": (
            "Equivalent complete datasets only; "
            "prefer fewer workers within 10% "
            "of fastest measured extraction time."
        ),
    }


# ============================================================
# BENCHMARK CHECKPOINT AND RECOVERY
# ============================================================

def benchmark_plan_key(
    query,
    axes,
    profiles,
    regions,
):

    return digest_json({
        "version": APP_VERSION,

        "query_hash": query.query_hash,

        "axis_digest": axes.digest,

        "snapshot": axes.snapshot,

        "profiles": [
            asdict(profile)

            for profile in profiles
        ],

        "regions": regions,

        "rate_limit": MAX_REQUESTS_PER_SECOND,

        "worker_delay": WORKER_REQUEST_DELAY,
    })


def load_benchmark_checkpoint(
    plan_key,
    axes,
    run_dir,
):

    # Reuse completed benchmark profiles only when
    # the source snapshot and complete plan match.
    #
    # Never combine unknown or changed snapshots.

    if not axes.snapshot:

        return []

    candidates = sorted(
        OUTPUT_ROOT.glob(
            "*/benchmark_progress.json"
        ),

        key=lambda path: path.stat().st_mtime,

        reverse=True,
    )

    for path in candidates:

        if path.parent == run_dir:

            continue

        checkpoint = read_json(
            path
        )

        if checkpoint.get(
            "plan_key"
        ) != plan_key:

            continue

        if checkpoint.get(
            "snapshot"
        ) != axes.snapshot:

            continue

        valid_reports = []

        for report in checkpoint.get(
            "profiles",
            [],
        ):

            if (
                report.get("status")
                != "COMPLETED"

                or report.get("equivalent")
                is not True

                or not report.get(
                    "snapshot_verified"
                )

                or not report.get(
                    "complete_coverage"
                )

                or not report.get(
                    "dataset_digest"
                )
            ):

                break

            directory = Path(
                report["output_directory"]
            )

            if not directory.is_dir():

                break

            exported_csv = (
                report.get(
                    "exports",
                    {},
                )

                .get(
                    "csv",
                    {},
                )

                .get(
                    "path"
                )
            )

            if (
                not exported_csv

                or not Path(
                    exported_csv
                ).is_file()
            ):

                break

            exported_excel = (
                report.get(
                    "exports",
                    {},
                )

                .get(
                    "excel",
                    {},
                )

                .get(
                    "path"
                )
            )

            if (
                not exported_excel

                or not Path(
                    exported_excel
                ).is_file()
            ):

                break

            recovered = dict(
                report
            )

            recovered["recovered"] = True

            recovered["recovered_from"] = str(
                path.parent
            )

            valid_reports.append(
                recovered
            )

        if valid_reports:

            return valid_reports

    return []


def save_benchmark_checkpoint(
    run_dir,
    query,
    axes,
    plan_key,
    reports,
    profiles,
    regions,
    status,
):

    summary = {
        "app": APP_NAME,

        "version": APP_VERSION,

        "mode": "benchmark",

        "status": status,

        "query_hash": query.query_hash,

        "axis_digest": axes.digest,

        "snapshot": axes.snapshot,

        "plan_key": plan_key,

        "regions": [
            list(region)

            for region in regions
        ],

        "planned_profiles": [
            asdict(profile)

            for profile in profiles
        ],

        "profiles": reports,

        "recommended_profile": (
            recommend_profile(reports)
        ),

        "updated_at": utc_now(),

        "limitations": [
            "Performance varies with server load and cache.",

            "Benchmark profiles cover selected "
            "beginning, middle and ending regions.",

            "Equivalence is verified within benchmark scope.",

            "Export timing is excluded from extraction timing.",

            "No automatic NON EMPTY filtering.",

            "No measure-unit conversion.",

            "Rate limits and server backoff are respected.",
        ],
    }

    write_json(
        run_dir / "benchmark_progress.json",

        summary,
    )

    return summary


# ============================================================
# MULTI-REGION PERFORMANCE BENCHMARK
# ============================================================

def run_benchmark(
    query,
    axes,
    max_workers,
    auth,
    control,
    run_dir,
    csv_writer,
    csv_handle,
):

    regions = build_benchmark_regions(
        axes.x_count
    )

    profiles = build_benchmark_profiles(
        max_workers,

        axes,
    )

    plan_key = benchmark_plan_key(
        query,

        axes,

        profiles,

        regions,
    )

    reports = load_benchmark_checkpoint(
        plan_key,

        axes,

        run_dir,
    )

    valid_profile_names = {
        profile.name

        for profile in profiles
    }

    reports = [
        report

        for report in reports

        if report["profile"]
        in valid_profile_names
    ]

    completed_names = {
        report["profile"]

        for report in reports
    }

    baseline_report = next(
        (
            report

            for report in reports

            if report["profile"]
            == "01_REFERENCE_SUBSET"
        ),

        None,
    )

    baseline_digest = (
        baseline_report["dataset_digest"]

        if baseline_report

        else None
    )

    console.print(
        make_panel(
            kv_table([
                (
                    "Regions",

                    str(regions),
                ),

                (
                    "Sample X tuples",

                    sum(
                        count

                        for _, count in regions
                    ),
                ),

                (
                    "Sample Y tuples",

                    axes.y_count,
                ),

                (
                    "Profiles",

                    len(profiles),
                ),

                (
                    "Recovered profiles",

                    len(reports),
                ),
            ]),

            "BENCHMARK PLAN",
        )
    )

    status = "PARTIAL"

    save_benchmark_checkpoint(
        run_dir,

        query,

        axes,

        plan_key,

        reports,

        profiles,

        regions,

        "RUNNING",
    )

    try:

        for profile in profiles:

            if STOP_EVENT.is_set():

                status = "STOPPED"

                break

            if control.service_busy:

                status = "PARTIAL"

                break

            if profile.name in completed_names:

                continue

            report, rows = run_scheduler(
                query,

                axes,

                profile,

                regions,

                auth,

                control,

                run_dir,

                csv_writer,

                csv_handle,

                mode="benchmark",

                benchmark_results=reports,
            )

            complete = (
                report["status"] == "COMPLETED"

                and report["complete_coverage"]

                and not report["duplicate_keys"]
            )

            if profile.name == "01_REFERENCE_SUBSET":

                report["equivalent"] = (
                    True

                    if complete

                    else None
                )

                if complete:

                    baseline_report = dict(
                        report
                    )

                    baseline_digest = report[
                        "dataset_digest"
                    ]

            else:

                report["equivalent"] = (
                    complete

                    and baseline_digest is not None

                    and report["dataset_digest"]
                    == baseline_digest
                )

            report["speedup"] = (
                round(
                    baseline_report["elapsed"]

                    / max(
                        report["elapsed"],
                        0.001,
                    ),

                    3,
                )

                if baseline_report

                else None
            )

            # Preserve results before any data export.

            reports.append(
                report
            )

            save_benchmark_checkpoint(
                run_dir,

                query,

                axes,

                plan_key,

                reports,

                profiles,

                regions,

                "RUNNING",
            )

            if rows:

                export_profile_data(
                    axes,

                    rows,

                    report,
                )

            else:

                report["export_status"] = "SKIPPED"

            write_json(
                Path(
                    report["output_directory"]
                ) / "summary.json",

                report,
            )

            save_benchmark_checkpoint(
                run_dir,

                query,

                axes,

                plan_key,

                reports,

                profiles,

                regions,

                "RUNNING",
            )

            if STOP_EVENT.is_set():

                status = "STOPPED"

                break

            if (
                not complete

                or report.get("fatal_error")

                or report.get("equivalent")
                is not True

                or control.service_busy
            ):

                status = "PARTIAL"

                break

            if (
                len(reports)
                < len(profiles)
            ):

                interruptible_sleep(
                    BENCHMARK_PAUSE_SECONDS
                )

        if (
            len(reports) == len(profiles)

            and all(
                report.get("equivalent") is True

                for report in reports
            )
        ):

            status = "COMPLETED"

    except (
        KeyboardInterrupt,
        StopRequested,
    ):

        STOP_EVENT.set()

        status = "STOPPED"

    finally:

        if STOP_EVENT.is_set():

            status = "STOPPED"

        summary = save_benchmark_checkpoint(
            run_dir,

            query,

            axes,

            plan_key,

            reports,

            profiles,

            regions,

            status,
        )

    return summary


# ============================================================
# LOAD VERIFIED PERFORMANCE CONFIGURATION
# ============================================================

def load_recommended_profile(
    query,
    axes,
    max_workers,
):

    candidates = sorted(
        [
            *OUTPUT_ROOT.glob(
                "*/summary.json"
            ),

            *OUTPUT_ROOT.glob(
                "*/benchmark_progress.json"
            ),
        ],

        key=lambda path: path.stat().st_mtime,

        reverse=True,
    )

    for path in candidates:

        report = read_json(
            path
        )

        if report.get("version") != APP_VERSION:

            continue

        if report.get("mode") != "benchmark":

            continue

        if report.get("query_hash") != query.query_hash:

            continue

        if report.get("axis_digest") != axes.digest:

            continue

        # Never reuse old tuning results when the
        # source snapshot cannot be verified.

        if not axes.snapshot:

            continue

        if report.get("snapshot") != axes.snapshot:

            continue

        matching_reports = [
            item

            for item in report.get(
                "profiles",
                [],
            )

            if item.get(
                "snapshot_verified"
            )
        ]

        recommendation = recommend_profile(
            matching_reports
        )

        if not recommendation:

            continue

        try:

            workers = int(
                recommendation["workers"]
            )

            x_chunk = int(
                recommendation["x_chunk"]
            )

            y_chunk = int(
                recommendation["y_chunk"]
            )

            strategy = recommendation[
                "strategy"
            ]

        except (
            KeyError,
            ValueError,
            TypeError,
        ):

            continue

        if not 1 <= workers <= max_workers:

            continue

        if not 1 <= x_chunk <= 20:

            continue

        if not 1 <= y_chunk <= axes.y_count:

            continue

        if strategy not in (
            "subset",
            "explicit",
        ):

            continue

        return (
            Profile(
                name="FULL_EXTRACTION",

                strategy=strategy,

                workers=workers,

                x_chunk=x_chunk,

                y_chunk=y_chunk,
            ),

            {
                "source": str(path),

                "recommendation": recommendation,

                "verified": True,
            },
        )

    # Conservative fallback:
    # use the original query semantics at low concurrency
    # until an optimized profile has been validated.

    return (
        Profile(
            name="FULL_EXTRACTION",

            strategy="subset",

            workers=1,

            x_chunk=DEFAULT_X_CHUNK,

            y_chunk=min(
                DEFAULT_Y_CHUNK,

                axes.y_count,
            ),
        ),

        {
            "source": "CONSERVATIVE_DEFAULT",

            "verified": False,

            "note": (
                "No matching validated benchmark found. "
                "Using one worker and original Subset MDX."
            ),
        },
    )


# ============================================================
# FULL DATA EXTRACTION
# ============================================================

def run_full_extraction(
    query,
    axes,
    max_workers,
    auth,
    control,
    run_dir,
    csv_writer,
    csv_handle,
):

    profile, selection = load_recommended_profile(
        query,

        axes,

        max_workers,
    )

    console.print(
        make_panel(
            kv_table([
                (
                    "Strategy",

                    profile.strategy,
                ),

                (
                    "Workers",

                    profile.workers,
                ),

                (
                    "X chunk",

                    profile.x_chunk,
                ),

                (
                    "Y chunk",

                    profile.y_chunk,
                ),

                (
                    "Benchmark verified",

                    selection["verified"],
                ),

                (
                    "Source",

                    selection["source"],
                ),
            ]),

            "SELECTED EXTRACTION PROFILE",
        )
    )

    report, rows = run_scheduler(
        query,

        axes,

        profile,

        [
            (
                0,

                axes.x_count,
            ),
        ],

        auth,

        control,

        run_dir,

        csv_writer,

        csv_handle,

        mode="run",

        benchmark_results=[],
    )

    report["profile_selection"] = (
        selection
    )

    if rows:

        export_profile_data(
            axes,

            rows,

            report,
        )

    else:

        report["export_status"] = "SKIPPED"

    write_json(
        Path(
            report["output_directory"]
        ) / "summary.json",

        report,
    )

    return {
        "mode": "run",

        "status": report["status"],

        "query_hash": query.query_hash,

        "axis_digest": axes.digest,

        "snapshot": axes.snapshot,

        "profile": report,

        "expected_rows": (
            axes.x_count

            * axes.y_count
        ),

        "returned_rows": len(
            rows
        ),

        "expected_coordinates": (
            axes.x_count

            * axes.y_count

            * len(query.measures)
        ),

        "exports": report.get(
            "exports",
            {},
        ),

        "export_status": report.get(
            "export_status"
        ),
    }


# ============================================================
# BENCHMARK COMPARISON CSV
# ============================================================

def export_benchmark_comparison(
    run_dir,
    summary,
):

    profiles = summary.get(
        "profiles",
        [],
    )

    if not profiles:

        return

    path = (
        run_dir
        / "benchmark_comparison.csv"
    )

    fields = (
        "profile",

        "strategy",

        "workers_count",

        "x_chunk",

        "y_chunk",

        "status",

        "equivalent",

        "requests",

        "failed_requests",

        "elapsed",

        "average_ms",

        "median_ms",

        "p95_ms",

        "max_ms",

        "requests_per_second",

        "cells_per_second",

        "slow_requests",

        "returned_rows",

        "non_null_cells",

        "response_bytes",

        "download_bytes",

        "speedup",

        "ttfb_ms",

        "response_wait_estimate_ms",

        "receive_ms",

        "file_write_ms",

        "parse_ms",

        "csv_export_ms",

        "excel_export_ms",

        "export_total_ms",

        "export_status",

        "snapshot_verified",

        "recovered",

        "dataset_digest",
    )

    with path.open(
        "w",

        encoding="utf-8-sig",

        newline="",
    ) as handle:

        writer = csv.DictWriter(
            handle,

            fieldnames=fields,

            extrasaction="ignore",
        )

        writer.writeheader()

        for profile in profiles:

            flattened = {
                **profile,

                **profile.get(
                    "timing",
                    {},
                ),
            }

            writer.writerow(
                flattened
            )


# ============================================================
# STARTUP DISPLAY
# ============================================================

def render_startup(
    query,
    args,
    run_dir,
):

    console.clear()

    console.print(
        Panel(
            Text(
                APP_NAME

                + f"\nVERSION {APP_VERSION}"

                + "\n"

                + "XMLA / MDX / CURL_CFFI / CSV / EXCEL",

                justify="center",

                style="bold bright_cyan",
            ),

            border_style="cyan",
        )
    )

    console.print(
        make_panel(
            kv_table([
                (
                    "Mode",

                    args.mode.upper(),
                ),

                (
                    "Category",

                    "MV — รถยนต์",
                ),

                (
                    "Cube",

                    "FY — ปีบัญชี",
                ),

                (
                    "Catalog",

                    CATALOG,
                ),

                (
                    "Max workers",

                    args.workers,
                ),

                (
                    "HTTP",

                    "NATIVE CALLBACK",
                ),

                (
                    "Request rate cap",

                    f"{MAX_REQUESTS_PER_SECOND:.1f}/s",
                ),

                (
                    "Query hash",

                    query.query_hash,
                ),

                (
                    "Output",

                    str(run_dir),
                ),
            ]),

            "ENGINE CONFIGURATION",
        )
    )

    fields = Table(
        box=box.SIMPLE,

        expand=True,
    )

    fields.add_column(
        "Axis",

        style="cyan",
    )

    fields.add_column(
        "Caption"
    )

    fields.add_column(
        "Verified MDX"
    )

    for prefix, captions, identifiers in (
        (
            "X",

            X_CAPTIONS,

            query.x,
        ),

        (
            "Y",

            Y_CAPTIONS,

            query.y,
        ),

        (
            "M",

            MEASURE_CAPTIONS,

            query.measures,
        ),
    ):

        for index, (
            caption,

            identifier,

        ) in enumerate(
            zip(
                captions,

                identifiers,
            ),

            start=1,
        ):

            fields.add_row(
                f"{prefix}{index}",

                Text(caption),

                Text(identifier),
            )

    console.print(
        make_panel(
            fields,

            "VERIFIED FIXED QUERY",
        )
    )


# ============================================================
# FINAL PERFORMANCE REPORT
# ============================================================

def render_final_report(
    summary,
    run_dir,
):

    status = summary.get(
        "status",

        "UNKNOWN",
    )

    console.print()

    console.print(
        Panel(
            Text(
                APP_NAME

                + "\n"

                + status,

                justify="center",

                style=(
                    "bold green"

                    if status == "COMPLETED"

                    else "bold yellow"
                ),
            ),

            border_style=(
                "green"

                if status == "COMPLETED"

                else "yellow"
            ),
        )
    )

    if summary.get(
        "mode"
    ) == "benchmark":

        table = Table(
            box=box.SIMPLE,

            expand=True,

            title="BENCHMARK RESULTS",
        )

        for name in (
            "Profile",

            "W",

            "Chunk",

            "Requests",

            "Extract",

            "P95",

            "Cells/s",

            "CSV",

            "Excel",

            "Match",
        ):

            table.add_column(
                name
            )

        for profile in summary.get(
            "profiles",
            [],
        ):

            equivalent = profile.get(
                "equivalent"
            )

            table.add_row(
                profile["profile"],

                str(
                    profile["workers_count"]
                ),

                f"{profile['x_chunk']}/"
                f"{profile['y_chunk']}",

                str(
                    profile["requests"]
                ),

                format_duration(
                    profile["elapsed"]
                ),

                format_ms(
                    profile.get("p95_ms")
                ),

                f"{profile['cells_per_second']:.2f}",

                format_ms(
                    profile.get(
                        "csv_export_ms"
                    )
                ),

                format_ms(
                    profile.get(
                        "excel_export_ms"
                    )
                ),

                Text(
                    "YES"

                    if equivalent is True

                    else "NO"

                    if equivalent is False

                    else "—",

                    style=(
                        "green"

                        if equivalent is True

                        else "yellow"
                    ),
                ),
            )

        console.print(
            table
        )

        recommendation = summary.get(
            "recommended_profile"
        )

        if recommendation:

            console.print(
                make_panel(
                    kv_table([
                        (
                            "Profile",

                            recommendation[
                                "profile"
                            ],
                        ),

                        (
                            "MDX strategy",

                            recommendation[
                                "strategy"
                            ],
                        ),

                        (
                            "Workers",

                            recommendation[
                                "workers"
                            ],
                        ),

                        (
                            "Chunk X/Y",

                            f"{recommendation['x_chunk']}/"
                            f"{recommendation['y_chunk']}",
                        ),

                        (
                            "Sample elapsed",

                            format_duration(
                                recommendation[
                                    "sample_elapsed"
                                ]
                            ),
                        ),

                        (
                            "Cells/sec",

                            recommendation[
                                "sample_cells_per_second"
                            ],
                        ),

                        (
                            "Sample speedup",

                            (
                                f"{recommendation['sample_speedup']:.2f}x"

                                if recommendation.get(
                                    "sample_speedup"
                                ) is not None

                                else "—"
                            ),
                        ),

                        (
                            "Snapshot verified",

                            recommendation[
                                "snapshot_verified"
                            ],
                        ),
                    ]),

                    "RECOMMENDED PROFILE",
                )
            )

    elif summary.get(
        "mode"
    ) == "run":

        report = summary.get(
            "profile",
            {},
        )

        if report:

            console.print(
                make_panel(
                    kv_table([
                        (
                            "Strategy",

                            report["strategy"],
                        ),

                        (
                            "Rows",

                            f"{report['returned_rows']:,}/"
                            f"{report['expected_rows']:,}",
                        ),

                        (
                            "Missing rows",

                            report["missing_rows"],
                        ),

                        (
                            "Duplicate rows",

                            report["duplicate_keys"],
                        ),

                        (
                            "Requests",

                            report["requests"],
                        ),

                        (
                            "Extraction",

                            format_duration(
                                report["elapsed"]
                            ),
                        ),

                        (
                            "Average latency",

                            format_ms(
                                report["average_ms"]
                            ),
                        ),

                        (
                            "P95 latency",

                            format_ms(
                                report["p95_ms"]
                            ),
                        ),

                        (
                            "CSV export",

                            format_ms(
                                report.get(
                                    "csv_export_ms"
                                )
                            ),
                        ),

                        (
                            "Excel export",

                            format_ms(
                                report.get(
                                    "excel_export_ms"
                                )
                            ),
                        ),

                        (
                            "Export status",

                            report.get(
                                "export_status",
                                "SKIPPED",
                            ),
                        ),

                        (
                            "Bottleneck",

                            report["bottleneck"][
                                "name"
                            ],
                        ),
                    ]),

                    "FULL EXTRACTION RESULT",
                )
            )

    console.print(
        make_panel(
            kv_table([
                (
                    "Run directory",

                    str(run_dir),
                ),

                (
                    "Main summary",

                    str(
                        run_dir
                        / "summary.json"
                    ),
                ),

                (
                    "Request metrics",

                    str(
                        run_dir
                        / "requests.csv"
                    ),
                ),

                (
                    "Axis discovery",

                    str(
                        run_dir
                        / "axis_discovery.json"
                    ),
                ),

                (
                    "Data output",

                    "Profile/data.csv + data.xlsx",
                ),

                (
                    "Partial output",

                    "Profile/data_partial.csv + data_partial.xlsx",
                ),
            ]),

            "OUTPUT FILES",
        )
    )

    if summary.get(
        "error"
    ):

        console.print(
            Text(
                summary["error"],

                style="bold red",
            )
        )


# ============================================================
# MAIN
# ============================================================

def main():

    args = parse_args()

    STOP_EVENT.clear()

    started = time.monotonic()

    run_id = (
        datetime.now().strftime(
            "%Y%m%d_%H%M%S"
        )

        + "_"

        + uuid.uuid4().hex[:6]
    )

    run_dir = (
        OUTPUT_ROOT
        / run_id
    )

    run_dir.mkdir(
        parents=True,

        exist_ok=True,
    )

    summary = {
        "app": APP_NAME,

        "version": APP_VERSION,

        "run_id": run_id,

        "mode": args.mode,

        "status": "STARTING",

        "started_at": utc_now(),

        "category": CATEGORY,

        "cube_type": CUBE_TYPE,
    }

    auth = None
    control = None

    try:

        env = load_env()

        query = load_fixed_query()

        control = RequestControl()

        auth = AuthenticationManager(
            env,

            control,
        )

        summary["query"] = {
            "hash": query.query_hash,

            "catalog": CATALOG,

            "cube": CUBE,

            "x": list(query.x),

            "y": list(query.y),

            "measures": list(
                query.measures
            ),

            "x_captions": list(
                X_CAPTIONS
            ),

            "y_captions": list(
                Y_CAPTIONS
            ),

            "measure_captions": list(
                MEASURE_CAPTIONS
            ),

            "mapping_source": (
                query.mapping_source
            ),
        }

        render_startup(
            query,

            args,

            run_dir,
        )

        discovery_started = (
            time.perf_counter()
        )

        axes = discover_axes(
            query,

            auth,

            control,

            run_dir,
        )

        summary["discovery_ms"] = (
            elapsed_ms(
                discovery_started
            )
        )

        summary["axis_counts"] = {
            "x_members": axes.x_count,

            "y_members": axes.y_count,

            "measures": len(
                query.measures
            ),

            "expected_rows": (
                axes.x_count
                * axes.y_count
            ),

            "expected_coordinates": (
                axes.x_count

                * axes.y_count

                * len(query.measures)
            ),
        }

        summary["axis_digest"] = axes.digest

        summary["snapshot"] = axes.snapshot

        console.print(
            make_panel(
                kv_table([
                    (
                        "Authentication",

                        Text(
                            "VALID",

                            style="green",
                        ),
                    ),

                    (
                        "X tuples",

                        f"{axes.x_count:,}",
                    ),

                    (
                        "Y tuples",

                        f"{axes.y_count:,}",
                    ),

                    (
                        "Measures",

                        len(query.measures),
                    ),

                    (
                        "Potential data rows",

                        f"{axes.x_count * axes.y_count:,}",
                    ),

                    (
                        "Potential cells",

                        f"{axes.x_count * axes.y_count * len(query.measures):,}",
                    ),

                    (
                        "Discovery time",

                        format_ms(
                            summary["discovery_ms"]
                        ),
                    ),

                    (
                        "Snapshot",

                        axes.snapshot or "UNKNOWN",
                    ),
                ]),

                "AXIS DISCOVERY COMPLETED",
            )
        )

        with (
            run_dir / "requests.csv"
        ).open(
            "w",

            encoding="utf-8-sig",

            newline="",

            buffering=1024 * 1024,
        ) as csv_handle:

            csv_writer = csv.DictWriter(
                csv_handle,

                fieldnames=REQUEST_CSV_FIELDS,

                extrasaction="ignore",
            )

            csv_writer.writeheader()

            csv_handle.flush()

            if args.mode == "benchmark":

                result = run_benchmark(
                    query,

                    axes,

                    args.workers,

                    auth,

                    control,

                    run_dir,

                    csv_writer,

                    csv_handle,
                )

            else:

                result = run_full_extraction(
                    query,

                    axes,

                    args.workers,

                    auth,

                    control,

                    run_dir,

                    csv_writer,

                    csv_handle,
                )

            summary.update(
                result
            )

        if args.mode == "benchmark":

            export_benchmark_comparison(
                run_dir,

                summary,
            )

    except (
        KeyboardInterrupt,
        StopRequested,
    ):

        STOP_EVENT.set()

        summary["status"] = "STOPPED"

    except (
        AuthenticationError,
        ServiceBusyError,
        EngineError,
    ) as exc:

        summary["status"] = "FAILED"

        summary["error"] = safe_error(
            exc
        )

    except Exception as exc:

        summary["status"] = "FAILED"

        summary["error"] = safe_error(
            exc
        )

    finally:

        # Preserve completed benchmark profiles even
        # when the user stops between profiles or
        # interrupts the export stage.

        if args.mode == "benchmark":

            checkpoint = read_json(
                run_dir
                / "benchmark_progress.json"
            )

            if checkpoint:

                for key in (
                    "profiles",

                    "recommended_profile",

                    "plan_key",

                    "regions",

                    "planned_profiles",

                    "limitations",
                ):

                    if key in checkpoint:

                        summary[key] = (
                            checkpoint[key]
                        )

                if STOP_EVENT.is_set():

                    summary["status"] = "STOPPED"

        summary["finished_at"] = utc_now()

        summary["total_elapsed_seconds"] = round(
            time.monotonic() - started,

            3,
        )

        if auth is not None:

            summary["authentication"] = {
                "token_refresh_count": (
                    auth.refresh_count
                ),

                "token_refresh_total_ms": round(
                    auth.refresh_total_ms,

                    3,
                ),
            }

        if control is not None:

            summary["request_control"] = {
                "rate_cap": (
                    MAX_REQUESTS_PER_SECOND
                ),

                "worker_delay": (
                    WORKER_REQUEST_DELAY
                ),

                "service_busy": (
                    control.service_busy
                ),

                "cooldown_remaining_seconds": round(
                    control.remaining(),

                    3,
                ),

                "cooldown_reason": (
                    control.cooldown_reason
                ),
            }

        write_json(
            run_dir
            / "summary.json",

            summary,
        )

    render_final_report(
        summary,

        run_dir,
    )

    if summary["status"] != "COMPLETED":

        return 1

    if (
        args.mode == "run"

        and summary.get("export_status")
        != "COMPLETED"
    ):

        return 2

    if args.mode == "benchmark":

        if any(
            profile.get("export_status")
            != "COMPLETED"

            for profile in summary.get(
                "profiles",
                [],
            )
        ):

            return 2

    return 0


if __name__ == "__main__":

    raise SystemExit(
        main()
    )