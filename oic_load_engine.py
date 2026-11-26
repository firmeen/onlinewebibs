#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC HIGH-PERFORMANCE REQUEST LOAD ENGINE
========================================

Standalone MV/FY XMLA Request Engine.

INSTALL
-------
python -m pip install curl_cffi rich

COMMANDS
--------
python oic_load_engine.py
python oic_load_engine.py --workers 4
python oic_load_engine.py --workers 4 --mode benchmark
python oic_load_engine.py --workers 4 --mode run

AUTHENTICATION
--------------
Reuse existing .env:

    OIC_COOKIE=...
    OIC_BEARER_TOKEN=...

FIXED QUERY
-----------
Category : MV
Cube     : FY

X / ROWS:
    1. แบรนด์/ยี่ห้อ
    2. ประเภทกรมธรรม์

Y / COLUMNS:
    1. ปีบัญชี

VALUES:
    1. จำนวนเคลม
    2. เงินเอาประกันภัย (ล้าน)
    3. เบี้ยประกันภัย (ล้าน)
    4. จำนวนกรมธรรม์
    5. จำนวนเคลมที่ยังไม่ปิด
    6. จำนวนเคลมที่มีการรับคืน
    7. จำนวนกรมธรรม์ที่มีเคลม

The fixed field mappings are resolved from the existing local
verified MV/FY catalog. Missing or ambiguous mappings stop execution.

No existing collector modules are imported or modified.
No credentials are written to reports.
No automatic changes are made to .env.

Performance improvements must be verified by actual benchmarks.
The engine cannot guarantee lower OLAP server processing time.

Benchmark results depend on server load, cache and network conditions.
Raw measure values are preserved without unverified unit conversions.
"""

from __future__ import annotations

import argparse
import base64
import csv
import hashlib
import json
import math
import os
import queue
import re
import statistics
import threading
import time
import uuid

from dataclasses import dataclass
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from email.utils import parsedate_to_datetime
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from curl_cffi import CurlInfo
from curl_cffi import requests as curl_requests
from curl_cffi.requests.exceptions import RequestException

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
APP_VERSION = "1.0.0"

ROOT = Path(__file__).resolve().parent

ENV_FILE = ROOT / ".env"

OUTPUT_ROOT = ROOT / "output" / "load_engine"

BASE_URL = "https://onlinewebibs.oic.or.th"

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

HOSTNAME = "onlinewebibs.oic.or.th"

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)

DEFAULT_WORKERS = 4
MAX_WORKERS = 8

CONNECT_TIMEOUT = 15
REQUEST_TIMEOUT = 90

TOKEN_REFRESH_MARGIN = 60

MAX_RESPONSE_BYTES = 100 * 1024 * 1024

# Conservative local limits, not official OIC service limits.
# Do not increase beyond the service's authorized usage policy.
MAX_REQUESTS_PER_SECOND = 4.0
WORKER_REQUEST_DELAY = 1.0

MAX_ATTEMPTS = 2
SLOW_REQUEST_MS = 5000

BENCHMARK_X_MEMBERS = 10
BENCHMARK_Y_MEMBERS = 4

DEFAULT_X_CHUNK = 5
DEFAULT_Y_CHUNK = 10

MAX_TASKS_PER_RUN = 100_000

RECENT_HISTORY_LIMIT = 100
RECENT_WINDOW_SECONDS = 60

REFRESH_INTERVAL = 1.0

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"
MD_NS = "urn:schemas-microsoft-com:xml-analysis:mddataset"

ET.register_namespace("soap", SOAP_NS)

IDENTIFIER_PATTERN = re.compile(
    r"^\[[^\]\r\n]+\]\.\[[^\]\r\n]+\]$"
)

MEASURE_PATTERN = re.compile(
    r"^\[Measures\]\.\[[^\]\r\n]+\]$"
)

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

CURL_TIMING_NAMES = (
    "NAMELOOKUP_TIME",
    "CONNECT_TIME",
    "APPCONNECT_TIME",
    "PRETRANSFER_TIME",
    "STARTTRANSFER_TIME",
    "TOTAL_TIME",
    "SIZE_DOWNLOAD_T",
    "NUM_CONNECTS",
    "HTTP_VERSION",
)

CURL_INFOS = [
    getattr(CurlInfo, name)
    for name in CURL_TIMING_NAMES
    if hasattr(CurlInfo, name)
]

CSV_FIELDS = (
    "profile",
    "engine",
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
    "response_bytes",
    "download_bytes",
    "download_kib_per_sec",
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

console = Console(highlight=False)

STOP_EVENT = threading.Event()


# ============================================================
# EXCEPTIONS
# ============================================================

class EngineError(RuntimeError):
    pass


class AuthenticationError(EngineError):
    pass


class ResponseError(EngineError):
    pass


class StopRequested(EngineError):
    pass


class ResponseTooLarge(EngineError):
    pass


# ============================================================
# DATA MODELS
# ============================================================

@dataclass(frozen=True)
class QuerySpec:

    x: tuple[str, ...]
    y: tuple[str, ...]
    measures: tuple[str, ...]

    x_captions: tuple[str, ...]
    y_captions: tuple[str, ...]
    measure_captions: tuple[str, ...]

    mapping_source: str
    query_hash: str


@dataclass(frozen=True)
class Profile:

    name: str
    engine: str
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

    queued_at: float


# ============================================================
# GENERAL UTILITIES
# ============================================================

def utc_now() -> str:

    return datetime.now(timezone.utc).isoformat(
        timespec="milliseconds"
    )


def local_time() -> str:

    return datetime.now().strftime("%H:%M:%S")


def elapsed_ms(started: float) -> float:

    return round(
        (time.perf_counter() - started) * 1000,
        3,
    )


def format_duration(seconds: float | int) -> str:

    seconds = max(0, int(seconds))

    hours, remainder = divmod(seconds, 3600)
    minutes, seconds = divmod(remainder, 60)

    return f"{hours:02d}:{minutes:02d}:{seconds:02d}"


def format_ms(value: Any) -> str:

    if value is None:

        return "—"

    value = float(value)

    if value >= 1000:

        return f"{value / 1000:,.2f}s"

    return f"{value:,.0f}ms"


def format_bytes(value: int | float) -> str:

    value = float(value)

    for unit in ("B", "KiB", "MiB", "GiB"):

        if value < 1024 or unit == "GiB":

            return f"{value:,.2f} {unit}"

        value /= 1024

    return f"{value:,.2f} GiB"


def percentile(values: list[float], percent: float):

    if not values:

        return None

    ordered = sorted(values)

    position = max(
        0,
        min(
            len(ordered) - 1,
            math.ceil(
                len(ordered) * percent / 100
            ) - 1,
        ),
    )

    return round(ordered[position], 3)


def mean_or_none(values: list[float]):

    return round(statistics.mean(values), 3) if values else None


def normalize_caption(value: str) -> str:

    return re.sub(
        r"\s+",
        "",
        str(value or ""),
    ).casefold()


def read_json(path: Path) -> dict:

    try:

        data = json.loads(
            path.read_text(
                encoding="utf-8-sig"
            )
        )

        return data if isinstance(data, dict) else {}

    except (OSError, ValueError, TypeError):

        return {}


def write_json(path: Path, data: dict) -> None:

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
            ),
            encoding="utf-8",
        )

        temporary.replace(path)

    finally:

        temporary.unlink(
            missing_ok=True
        )


def check_stop() -> None:

    if STOP_EVENT.is_set():

        raise StopRequested(
            "Stop requested"
        )


def interruptible_sleep(seconds: float) -> None:

    if STOP_EVENT.wait(max(0, seconds)):

        raise StopRequested(
            "Stop requested"
        )


def parse_timestamp(value: str | None):

    if not value:

        return None

    try:

        return datetime.fromisoformat(
            value.replace("Z", "+00:00")
        )

    except (ValueError, TypeError):

        return None


def safe_error(exc: Exception) -> str:

    # Do not include HTTP headers, request dumps,
    # cookies, JWTs or complete server responses.

    if isinstance(
        exc,
        (
            AuthenticationError,
            ResponseError,
            ResponseTooLarge,
            StopRequested,
        ),
    ):

        return str(exc)[:200]

    if isinstance(exc, RequestException):

        code = getattr(exc, "code", None)

        return (
            f"{type(exc).__name__}"
            + (f" | curl code={code}" if code else "")
        )

    return type(exc).__name__


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
        help=(
            "Maximum parallel workers "
            f"(1-{MAX_WORKERS})"
        ),
    )

    parser.add_argument(
        "--mode",
        choices=("benchmark", "run"),
        default="benchmark",
        help=(
            "Benchmark the fixed MV/FY query "
            "or extract its full dataset"
        ),
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

def load_env() -> dict[str, str]:

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

    cookie = values.get(
        "OIC_COOKIE",
        "",
    ).strip()

    if not cookie:

        raise AuthenticationError(
            "OIC_COOKIE is missing from .env"
        )

    return values


def jwt_exp(token: str) -> int:

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

        decoded = base64.urlsafe_b64decode(
            payload.encode("ascii")
        )

        data = json.loads(decoded)

        return int(data.get("exp", 0))

    except Exception:

        return 0


# ============================================================
# VERIFIED FIXED QUERY
# ============================================================

def catalog_timestamp(data: dict) -> str:

    return str(
        data.get("report_updated_at")
        or data.get("updated_at")
        or data.get("finished_at")
        or data.get("refreshed_at")
        or ""
    )


def valid_catalog(data: dict) -> bool:

    return (
        data.get(
            "code",
            data.get("insurance_code"),
        ) == CATEGORY
        and data.get("cube_type", "FY") == CUBE_TYPE
        and data.get("catalog") == CATALOG
        and data.get("cube") == CUBE
        and isinstance(data.get("fields"), list)
        and bool(data["fields"])
        and data.get("status") not in (
            "ERROR",
            "UNAVAILABLE",
        )
    )


def selectable_verified_field(
    field: dict,
    expected_type: str,
) -> bool:

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

    elif mdx.startswith("[Measures]."):

        return False

    return True


def resolve_field(
    fields: list[dict],
    caption: str,
    expected_type: str,
) -> str:

    target = normalize_caption(caption)

    matches = []

    for field in fields:

        name = field.get(
            "caption",
            field.get("display", ""),
        )

        if normalize_caption(name) != target:

            continue

        if selectable_verified_field(
            field,
            expected_type,
        ):

            matches.append(field)

    # Prefer original OIC captions over additional XMLA metadata.
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
            f"Missing verified {expected_type} mapping: "
            f"{caption}"
        )

    if len(identifiers) != 1:

        raise EngineError(
            f"Ambiguous verified mapping: {caption}"
        )

    return next(iter(identifiers))


def load_fixed_query() -> QuerySpec:

    candidates = []

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

    for path in paths:

        if not path.is_file():

            continue

        data = read_json(path)

        if not valid_catalog(data):

            continue

        candidates.append(
            (
                catalog_timestamp(data),
                path.stat().st_mtime_ns,
                path,
                data,
            )
        )

    if not candidates:

        raise EngineError(
            "No valid local MV/FY verification catalog found. "
            "Export the verified catalog or retain "
            "output/field_verification/MV_FY.json."
        )

    # Use the newest valid catalog.
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
            "Duplicate measure identifiers"
        )

    canonical = {
        "category": CATEGORY,
        "cube_type": CUBE_TYPE,
        "catalog": CATALOG,
        "cube": CUBE,
        "x": x,
        "y": y,
        "measures": measures,
        "schema_version": 2,
    }

    query_hash = hashlib.sha256(
        json.dumps(
            canonical,
            sort_keys=True,
            ensure_ascii=False,
        ).encode("utf-8")
    ).hexdigest()[:16]

    return QuerySpec(
        x=x,
        y=y,
        measures=measures,
        x_captions=X_CAPTIONS,
        y_captions=Y_CAPTIONS,
        measure_captions=MEASURE_CAPTIONS,
        mapping_source=str(path),
        query_hash=query_hash,
    )


# ============================================================
# CURL SESSION
# ============================================================

def create_session(cookie: str):

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
# GLOBAL REQUEST CONTROL
# ============================================================

class RequestControl:

    def __init__(self):

        self.lock = threading.Lock()

        self.next_request_at = 0.0
        self.cooldown_until = 0.0

        self.cooldown_reason = ""

    def wait_slot(self) -> float:

        started = time.perf_counter()

        while True:

            check_stop()

            with self.lock:

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
                min(0.25, remaining)
            )

    def set_cooldown(
        self,
        seconds: float,
        reason: str,
    ):

        with self.lock:

            self.cooldown_until = max(
                self.cooldown_until,
                time.monotonic() + max(0, seconds),
            )

            self.cooldown_reason = reason

    def remaining(self) -> float:

        with self.lock:

            return max(
                0.0,
                self.cooldown_until
                - time.monotonic(),
            )


def retry_after_seconds(
    value: str | None,
    default: float,
) -> float:

    if not value:

        return default

    try:

        return max(
            0,
            float(value),
        )

    except ValueError:

        pass

    try:

        retry_at = parsedate_to_datetime(value)

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
# AUTHENTICATION
# ============================================================

class AuthenticationManager:

    def __init__(
        self,
        env: dict,
        control: RequestControl,
    ):

        self.cookie = env["OIC_COOKIE"]

        self.token = env.get(
            "OIC_BEARER_TOKEN",
            "",
        ).removeprefix(
            "Bearer "
        ).strip()

        self.lock = threading.Lock()

        self.control = control

        self.refresh_count = 0
        self.refresh_total_ms = 0.0

    def refresh_token(self, session) -> float:

        started = time.perf_counter()

        self.control.wait_slot()

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

            if response.status_code != 200:

                raise AuthenticationError(
                    "GenToken HTTP "
                    f"{response.status_code}"
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
                    "GenToken returned invalid or expired JWT"
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
        *,
        rejected_token: str | None = None,
    ) -> tuple[str, float, float]:

        waiting = time.perf_counter()

        with self.lock:

            lock_wait_ms = elapsed_ms(waiting)

            refresh_ms = 0.0

            expired = (
                jwt_exp(self.token)
                <= time.time() + TOKEN_REFRESH_MARGIN
            )

            rejected = (
                rejected_token is not None
                and self.token == rejected_token
            )

            if expired or rejected:

                check_stop()

                refresh_ms = self.refresh_token(
                    session
                )

            return (
                self.token,
                lock_wait_ms,
                refresh_ms,
            )


# ============================================================
# MDX BUILDER
# ============================================================

def build_axis_set(
    identifiers: tuple[str, ...],
) -> str:

    members = [
        identifier + ".AllMembers"
        for identifier in identifiers
    ]

    result = members[0]

    for member in members[1:]:

        result = (
            f"CrossJoin({result},{member})"
        )

    return result


def build_mdx(
    query: QuerySpec,
    task: RequestTask,
) -> str:

    x_set = build_axis_set(query.x)
    y_set = build_axis_set(query.y)

    x = (
        f"Subset({x_set},"
        f"{task.x_offset},"
        f"{task.x_chunk})"
    )

    y = (
        f"Subset({y_set},"
        f"{task.y_offset},"
        f"{task.y_chunk})"
    )

    measures = (
        "{"
        + ",".join(query.measures)
        + "}"
    )

    return (
        f"SELECT {y} "
        "DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,"
        "HIERARCHY_UNIQUE_NAME,"
        "MEMBER_VALUE ON COLUMNS, "
        f"CrossJoin({x},{measures}) "
        "DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,"
        "HIERARCHY_UNIQUE_NAME,"
        "MEMBER_VALUE ON ROWS "
        f"FROM [{CUBE}] "
        "CELL PROPERTIES "
        "VALUE,FORMAT_STRING,LANGUAGE"
    )


def build_count_mdx(
    query: QuerySpec,
) -> str:

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
# CURL TIMING EXTRACTION
# ============================================================

def get_curl_info(
    response,
    name: str,
):

    enum = getattr(
        CurlInfo,
        name,
        None,
    )

    if enum is None:

        return None

    infos = getattr(
        response,
        "infos",
        {},
    ) or {}

    return infos.get(enum)


def extract_curl_timings(
    response,
) -> dict:

    def milliseconds(name):

        value = get_curl_info(
            response,
            name,
        )

        if not isinstance(value, (int, float)):

            return None

        return round(value * 1000, 3)

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

    downloaded = get_curl_info(
        response,
        "SIZE_DOWNLOAD_T",
    )

    connections = get_curl_info(
        response,
        "NUM_CONNECTS",
    )

    def difference(a, b):

        if a is None or b is None:

            return None

        return round(
            max(0.0, a - b),
            3,
        )

    return {
        "dns_ms": dns,

        "connect_ms": difference(
            connect,
            dns,
        ),

        "tls_ms": difference(
            tls,
            connect,
        ),

        "pretransfer_ms": pretransfer,

        "ttfb_ms": ttfb,

        # Includes network transit and upstream processing.
        # This is not a measurement of pure OLAP CPU time.
        "response_wait_estimate_ms": difference(
            ttfb,
            pretransfer,
        ),

        "transfer_ms": total,

        "download_bytes": (
            int(downloaded)
            if isinstance(downloaded, (int, float))
            else None
        ),

        "new_connections": (
            int(connections)
            if isinstance(connections, (int, float))
            else None
        ),
    }


def http_version_name(value) -> str:

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
# HIGH-PERFORMANCE HTTP EXECUTION
# ============================================================

def execute_request(
    session,
    token: str,
    payload: bytes,
    destination: Path,
    engine: str,
) -> dict:

    started = time.perf_counter()

    temporary = destination.with_name(
        destination.name + ".partial"
    )

    destination.parent.mkdir(
        parents=True,
        exist_ok=True,
    )

    headers = dict(XMLA_HEADERS)

    headers["Authorization"] = (
        "Bearer " + token
    )

    response = None

    response_bytes = 0

    file_write_ms = 0.0

    first_body_at = None
    first_body_ms = None
    last_body_at = None

    header_wait_ms = None
    finalize_ms = None

    limit_exceeded = False
    callback_error = None

    status = None
    content_type = ""
    content_encoding = ""

    metrics = {}

    try:

        with temporary.open(
            "wb",
            buffering=256 * 1024,
        ) as file:

            def content_callback(chunk: bytes):

                nonlocal response_bytes
                nonlocal file_write_ms
                nonlocal first_body_at
                nonlocal first_body_ms
                nonlocal last_body_at
                nonlocal limit_exceeded
                nonlocal callback_error

                if STOP_EVENT.is_set():

                    return 0

                if not chunk:

                    return 0

                now = time.perf_counter()

                if first_body_at is None:

                    first_body_at = now

                    first_body_ms = round(
                        (
                            now
                            - transfer_started
                        ) * 1000,
                        3,
                    )

                response_bytes += len(chunk)

                if response_bytes > MAX_RESPONSE_BYTES:

                    limit_exceeded = True

                    return 0

                write_started = time.perf_counter()

                try:

                    written = file.write(chunk)

                except OSError as exc:

                    callback_error = exc

                    return 0

                finally:

                    file_write_ms += elapsed_ms(
                        write_started
                    )

                last_body_at = time.perf_counter()

                return written

            transfer_started = time.perf_counter()

            if engine == "callback":

                # Native libcurl callback avoids the
                # streaming iterator's intermediate queue.
                #
                # Curl timing reference:
                # https://curl-cffi.readthedocs.io/en/stable/api.html

                response = session.post(
                    XMLA_URL,
                    data=payload,
                    headers=headers,
                    content_callback=content_callback,
                    timeout=(
                        CONNECT_TIMEOUT,
                        REQUEST_TIMEOUT,
                    ),
                    allow_redirects=False,
                    accept_encoding="gzip, deflate, br",
                )

                # curl_infos are collected after the
                # synchronous transfer completes.
                metrics = extract_curl_timings(
                    response
                )

            elif engine == "iterator":

                # Reference engine: matches the existing
                # collector's streaming iterator approach.

                with session.stream(
                    "POST",
                    XMLA_URL,
                    data=payload,
                    headers=headers,
                    timeout=(
                        CONNECT_TIMEOUT,
                        REQUEST_TIMEOUT,
                    ),
                    allow_redirects=False,
                    accept_encoding="gzip, deflate, br",
                ) as stream_response:

                    response = stream_response

                    header_wait_ms = elapsed_ms(
                        transfer_started
                    )

                    for chunk in response.iter_content():

                        if not chunk:

                            continue

                        written = content_callback(
                            chunk
                        )

                        if written != len(chunk):

                            if STOP_EVENT.is_set():

                                raise StopRequested(
                                    "Request interrupted"
                                )

                            if limit_exceeded:

                                raise ResponseTooLarge(
                                    "Response exceeded configured limit"
                                )

                            raise ResponseError(
                                "Response write failed"
                            )

                # Streaming response timing snapshots
                # may not represent the final transfer.
                # Do not report unavailable CurlInfo values.
                metrics = {
                    "transfer_ms": elapsed_ms(
                        transfer_started
                    ),
                }

            else:

                raise EngineError(
                    "Unknown HTTP engine"
                )

            transfer_wall_ms = elapsed_ms(
                transfer_started
            )

            flush_started = time.perf_counter()

            file.flush()

            file_write_ms += elapsed_ms(
                flush_started
            )

        if STOP_EVENT.is_set():

            raise StopRequested(
                "Request interrupted"
            )

        if limit_exceeded:

            raise ResponseTooLarge(
                "Response exceeded configured limit"
            )

        if callback_error is not None:

            raise ResponseError(
                "Unable to write XML response"
            )

        status = int(
            response.status_code
        )

        content_type = response.headers.get(
            "Content-Type",
            "",
        )

        content_encoding = response.headers.get(
            "Content-Encoding",
            "",
        )

        retry_after = response.headers.get(
            "Retry-After"
        )

        version = getattr(
            response,
            "http_version",
            None,
        )

        if status == 200:

            finalize_started = time.perf_counter()

            temporary.replace(destination)

            finalize_ms = elapsed_ms(
                finalize_started
            )

        # Callback API does not expose the exact
        # headers-complete timestamp.
        # Keep header_wait_ms unavailable in callback mode.
        # TTFB and first_body_ms are measured separately.

        receive_ms = None

        if (
            first_body_at is not None
            and last_body_at is not None
        ):

            receive_ms = round(
                (
                    last_body_at
                    - first_body_at
                ) * 1000,
                3,
            )

        download_bytes = metrics.get(
            "download_bytes"
        )

        download_kib_per_sec = None

        if (
            download_bytes is not None
            and metrics.get("transfer_ms")
            and metrics.get("ttfb_ms") is not None
        ):

            download_duration = max(
                0.0,
                metrics["transfer_ms"]
                - metrics["ttfb_ms"],
            )

            if download_duration > 0:

                download_kib_per_sec = round(
                    download_bytes / 1024
                    / (download_duration / 1000),
                    3,
                )

        return {
            **metrics,

            "http_status": status,

            "http_version": http_version_name(
                version
            ),

            "content_type": content_type,
            "content_encoding": content_encoding,

            "retry_after": retry_after,

            "header_wait_ms": header_wait_ms,

            "first_body_ms": first_body_ms,

            # Includes callback/file handling time.
            "receive_ms": receive_ms,

            "file_write_ms": round(
                file_write_ms,
                3,
            ),

            "finalize_ms": finalize_ms,

            "transfer_wall_ms": transfer_wall_ms,

            "execute_ms": elapsed_ms(started),

            "response_bytes": response_bytes,

            "download_kib_per_sec": (
                download_kib_per_sec
            ),

            "response_file": (
                str(destination)
                if status == 200
                else None
            ),
        }

    except RequestException:

        if STOP_EVENT.is_set():

            raise StopRequested(
                "Request interrupted"
            )

        if limit_exceeded:

            raise ResponseTooLarge(
                "Response exceeded configured limit"
            )

        if callback_error is not None:

            raise ResponseError(
                "Unable to write XML response"
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
# XMLA RESPONSE PROCESSING
# ============================================================

def local_name(tag: str) -> str:

    return tag.rsplit("}", 1)[-1]


def load_xml_dataset(
    path: Path,
):

    try:

        root = ET.parse(path).getroot()

    except (OSError, ET.ParseError) as exc:

        raise ResponseError(
            "Invalid XML response"
        ) from exc

    for element in root.iter():

        if local_name(element.tag) == "Fault":

            raise ResponseError(
                "XMLA SOAP Fault"
            )

    dataset = root.find(
        f".//{{{MD_NS}}}root"
    )

    if dataset is None:

        raise ResponseError(
            "XMLA multidimensional dataset missing"
        )

    for element in dataset.iter():

        if local_name(element.tag) == "Error":

            raise ResponseError(
                "XMLA dataset contains an error"
            )

    return root, dataset


def read_axis(
    dataset,
    axis_name: str,
) -> list[list[dict]]:

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
                    "XMLA member has no unique identifier"
                )

            members.append({
                "hierarchy": member.get(
                    "Hierarchy",
                    "",
                ),

                "unique_name": unique_name,

                "caption": member.findtext(
                    f"{{{MD_NS}}}Caption",
                    default="",
                ),
            })

        tuples.append(members)

    return tuples


def read_cells(dataset) -> dict[int, str | None]:

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
                "Invalid XMLA CellOrdinal"
            ) from exc

        if ordinal < 0:

            raise ResponseError(
                "Negative XMLA CellOrdinal"
            )

        if ordinal in cells:

            raise ResponseError(
                "Duplicate XMLA CellOrdinal"
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


def is_measure(member: dict) -> bool:

    return (
        member["hierarchy"].casefold()
        == "[measures]"
        or member["unique_name"].casefold().startswith(
            "[measures]."
        )
    )


def canonical_value(value):

    if value is None:

        return None

    value = str(value).strip()

    if not value:

        return None

    try:

        number = Decimal(value)

        if number.is_finite():

            if number == 0:

                return "0"

            return str(number.normalize())

    except InvalidOperation:

        pass

    return value


def parse_response(
    path: Path,
    query: QuerySpec,
    task: RequestTask,
    *,
    capture_values: bool,
) -> dict:

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

    x_members = set()

    row_keys = []

    allowed_measures = {
        identifier.casefold()
        for identifier in query.measures
    }

    measures_by_row = {}

    for row in axis1:

        dimensions = tuple(
            member["unique_name"]
            for member in row
            if not is_measure(member)
        )

        measures = [
            member["unique_name"]
            for member in row
            if is_measure(member)
        ]

        if len(dimensions) != len(query.x):

            raise ResponseError(
                "Unexpected X dimension count"
            )

        if len(measures) != 1:

            raise ResponseError(
                "Expected one measure per XMLA row tuple"
            )

        measure = measures[0]

        if measure.casefold() not in allowed_measures:

            raise ResponseError(
                "Unexpected XMLA measure"
            )

        x_members.add(dimensions)

        measures_by_row.setdefault(
            dimensions,
            set(),
        ).add(measure.casefold())

        row_keys.append(
            dimensions + (measure,)
        )

    column_keys = []

    for column in axis0:

        if len(column) != len(query.y):

            raise ResponseError(
                "Unexpected Y dimension count"
            )

        column_keys.append(
            tuple(
                member["unique_name"]
                for member in column
            )
        )

    if len(set(row_keys)) != len(row_keys):

        raise ResponseError(
            "Duplicate XMLA row tuple"
        )

    if len(set(column_keys)) != len(column_keys):

        raise ResponseError(
            "Duplicate XMLA column tuple"
        )

    for returned in measures_by_row.values():

        if returned != allowed_measures:

            raise ResponseError(
                "Incomplete measure set"
            )

    x_count = len(x_members)

    y_count = len(column_keys)

    max_cell_ordinal = (
        len(row_keys)
        * len(column_keys)
    )

    if any(
        ordinal >= max_cell_ordinal
        for ordinal in cells
    ):

        raise ResponseError(
            "CellOrdinal outside XMLA axes"
        )

    if x_count != task.x_chunk:

        raise ResponseError(
            f"X coverage mismatch: "
            f"expected {task.x_chunk}, "
            f"received {x_count}"
        )

    if y_count != task.y_chunk:

        raise ResponseError(
            f"Y coverage mismatch: "
            f"expected {task.y_chunk}, "
            f"received {y_count}"
        )

    non_null_cells = sum(
        value is not None
        and bool(str(value).strip())
        for value in cells.values()
    )

    if x_count == 0 or y_count == 0:

        chunk_kind = "EMPTY_PAGE"

    elif non_null_cells == 0:

        chunk_kind = "NO_CELLS"

    else:

        chunk_kind = "DATA"

    comparison_values = {}

    if capture_values:

        column_count = len(column_keys)

        for row_index, row_key in enumerate(
            row_keys
        ):

            for col_index, col_key in enumerate(
                column_keys
            ):

                ordinal = (
                    row_index * column_count
                    + col_index
                )

                key = row_key + col_key

                comparison_values[key] = (
                    canonical_value(
                        cells.get(ordinal)
                    )
                )

    last_data_update = next(
        (
            element.text
            for element in root.iter()
            if local_name(element.tag) == "LastDataUpdate"
        ),
        None,
    )

    return {
        "x_members": x_count,
        "y_members": y_count,

        "axis0_tuples": len(axis0),
        "axis1_tuples": len(axis1),

        "cells": len(cells),

        "non_null_cells": non_null_cells,

        "chunk_kind": chunk_kind,

        "parse_ms": elapsed_ms(started),

        "last_data_update": last_data_update,

        "comparison_values": comparison_values,
    }


# ============================================================
# INITIAL CUBE VALIDATION AND MEMBER COUNTS
# ============================================================

def discover_axis_counts(
    query: QuerySpec,
    auth: AuthenticationManager,
    control: RequestControl,
    run_dir: Path,
) -> tuple[int, int, dict]:

    session = create_session(
        auth.cookie
    )

    destination = (
        run_dir
        / "responses"
        / "axis_counts.xml"
    )

    rejected_token = None

    try:

        payload = build_soap(
            build_count_mdx(query)
        )

        for attempt in range(MAX_ATTEMPTS):

            check_stop()

            token, lock_wait, refresh_ms = (
                auth.get_token(
                    session,
                    rejected_token=rejected_token,
                )
            )

            rate_wait = control.wait_slot()

            result = execute_request(
                session,
                token,
                payload,
                destination,
                "callback",
            )

            result.update({
                "token_lock_wait_ms": lock_wait,
                "token_refresh_ms": refresh_ms,
                "rate_wait_ms": rate_wait,
            })

            status = result["http_status"]

            if status == 200:

                break

            if status in (401, 403):

                if attempt + 1 >= MAX_ATTEMPTS:

                    raise AuthenticationError(
                        "XMLA authentication rejected"
                    )

                rejected_token = token

                continue

            if status in (429, 503):

                seconds = retry_after_seconds(
                    result.get("retry_after"),
                    default=15,
                )

                control.set_cooldown(
                    seconds,
                    f"HTTP {status}",
                )

                if attempt + 1 < MAX_ATTEMPTS:

                    continue

            raise ResponseError(
                f"Axis count HTTP {status}"
            )

        started = time.perf_counter()

        _, dataset = load_xml_dataset(
            destination
        )

        columns = read_axis(
            dataset,
            "Axis0",
        )

        cells = read_cells(dataset)

        if len(columns) != 2:

            raise ResponseError(
                "Unexpected axis-count result"
            )

        counts = {}

        for ordinal, column in enumerate(
            columns
        ):

            if len(column) != 1:

                raise ResponseError(
                    "Invalid axis-count tuple"
                )

            identifier = column[0][
                "unique_name"
            ].casefold()

            raw = cells.get(ordinal)

            if raw is None:

                raise ResponseError(
                    "Axis-count value is null"
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

            counts[identifier] = int(number)

        x_count = counts.get(
            "[measures].[__loadenginexcount]"
        )

        y_count = counts.get(
            "[measures].[__loadengineycount]"
        )

        if x_count is None or y_count is None:

            raise ResponseError(
                "Unable to determine X/Y cardinality"
            )

        result.update({
            "x_members": x_count,
            "y_members": y_count,
            "parse_ms": elapsed_ms(started),
        })

        write_json(
            run_dir / "axis_discovery.json",
            result,
        )

        return x_count, y_count, result

    finally:

        session.close()


# ============================================================
# WORKER STATE AND PERFORMANCE METRICS
# ============================================================

class EngineState:

    def __init__(
        self,
        profile: Profile,
        total_tasks: int,
    ):

        self.profile = profile

        self.total_tasks = total_tasks

        self.lock = threading.Lock()

        self.started = time.monotonic()

        self.workers = {
            wid: {
                "worker_id": wid,
                "status": "STARTING",
                "phase": "INITIALIZING",
                "x_offset": None,
                "y_offset": None,
                "requests": 0,
                "completed": 0,
                "failed": 0,
                "cells": 0,
                "bytes": 0,
                "last_http": None,
                "last_request_ms": None,
                "last_ttfb_ms": None,
                "request_started": None,
                "last_event": "Initializing",
                "event_time": local_time(),
            }
            for wid in range(
                1,
                profile.workers + 1,
            )
        }

        self.requests = []

        self.recent_requests = []

        self.events = []

        self.completed_tasks = 0
        self.failed_tasks = 0

        self.data_chunks = 0
        self.empty_pages = 0
        self.no_cell_chunks = 0

        self.cells = 0
        self.non_null_cells = 0

        self.response_bytes = 0

        self.fatal_error = None

    def update_worker(
        self,
        worker_id: int,
        **values,
    ):

        with self.lock:

            worker = self.workers[worker_id]

            worker.update(values)

            if "last_event" in values:

                worker["event_time"] = (
                    local_time()
                )

                self.events.append({
                    "time": worker["event_time"],
                    "worker": worker_id,
                    "message": worker["last_event"],
                })

                self.events = self.events[-20:]

    def add_request(
        self,
        record: dict,
    ):

        with self.lock:

            sample = dict(record)

            self.requests.append(sample)

            self.recent_requests.append(sample)

            self.recent_requests = (
                self.recent_requests[
                    -RECENT_HISTORY_LIMIT
                    * self.profile.workers:
                ]
            )

            worker = self.workers[
                record["worker_id"]
            ]

            worker["requests"] += 1

            worker["last_request_ms"] = (
                record.get("execute_ms")
            )

            worker["last_ttfb_ms"] = (
                record.get("ttfb_ms")
            )

            if record.get("http_status") is not None:

                worker["last_http"] = (
                    record["http_status"]
                )

    def finish_task(
        self,
        worker_id: int,
        result: dict | None,
    ):

        with self.lock:

            worker = self.workers[
                worker_id
            ]

            if result is None:

                self.failed_tasks += 1

                worker["failed"] += 1

                return

            self.completed_tasks += 1

            worker["completed"] += 1

            cells = result.get(
                "cells",
                0,
            )

            non_null = result.get(
                "non_null_cells",
                0,
            )

            received = result.get(
                "response_bytes",
                0,
            )

            self.cells += cells

            self.non_null_cells += non_null

            self.response_bytes += received

            worker["cells"] += cells

            worker["bytes"] += received

            kind = result.get("chunk_kind")

            if kind == "DATA":

                self.data_chunks += 1

            elif kind == "EMPTY_PAGE":

                self.empty_pages += 1

            elif kind == "NO_CELLS":

                self.no_cell_chunks += 1

    def set_fatal(self, message: str):

        with self.lock:

            if self.fatal_error is None:

                self.fatal_error = message

    def snapshot(self) -> dict:

        with self.lock:

            return {
                "profile": self.profile.name,

                "elapsed": max(
                    0.0,
                    time.monotonic() - self.started,
                ),

                "total_tasks": self.total_tasks,

                "completed_tasks": self.completed_tasks,

                "failed_tasks": self.failed_tasks,

                "data_chunks": self.data_chunks,

                "empty_pages": self.empty_pages,

                "no_cell_chunks": self.no_cell_chunks,

                "cells": self.cells,

                "non_null_cells": self.non_null_cells,

                "response_bytes": self.response_bytes,

                "requests": [
                    dict(item)
                    for item in self.requests
                ],

                "recent_requests": [
                    dict(item)
                    for item in self.recent_requests
                ],

                "workers": [
                    dict(worker)
                    for worker in self.workers.values()
                ],

                "events": [
                    dict(event)
                    for event in self.events
                ],

                "fatal_error": self.fatal_error,
            }


# ============================================================
# METRICS AGGREGATION
# ============================================================

def aggregate_metrics(snapshot: dict) -> dict:

    requests = snapshot["requests"]

    elapsed = max(
        0.001,
        snapshot["elapsed"],
    )

    durations = [
        float(record["execute_ms"])
        for record in requests
        if record.get("execute_ms") is not None
    ]

    now = datetime.now(
        timezone.utc
    )

    recent = []

    for record in snapshot["recent_requests"]:

        completed = parse_timestamp(
            record.get("completed_at")
        )

        if (
            completed is not None
            and (
                now - completed
            ).total_seconds()
            <= RECENT_WINDOW_SECONDS
        ):

            recent.append(record)

    recent_durations = [
        float(record["execute_ms"])
        for record in recent
        if record.get("execute_ms") is not None
    ]

    def average_metric(name):

        values = [
            float(record[name])
            for record in requests
            if record.get(name) is not None
        ]

        return mean_or_none(values)

    known_connections = [
        int(record["new_connections"])
        for record in requests
        if record.get("new_connections") is not None
    ]

    reused_transfers = sum(
        count == 0
        for count in known_connections
    )

    connections_total = sum(
        known_connections
    )

    reuse_percent = (
        reused_transfers
        / len(known_connections)
        * 100
        if known_connections
        else None
    )

    workers = snapshot["workers"]

    max_completed = max(
        (
            worker["completed"]
            for worker in workers
        ),
        default=0,
    )

    for worker in workers:

        worker["relative_progress"] = (
            worker["completed"] / max_completed
            if max_completed
            else 0.0
        )

        request_started = worker.get(
            "request_started"
        )

        if (
            worker.get("phase") == "REQUESTING"
            and request_started is not None
        ):

            worker["current_request_ms"] = (
                round(
                    (
                        time.monotonic()
                        - request_started
                    ) * 1000,
                    3,
                )
            )

        else:

            worker["current_request_ms"] = None

    ok_requests = sum(
        record.get("status") == "OK"
        for record in requests
    )

    failed_requests = sum(
        record.get("status") == "ERROR"
        for record in requests
    )

    finished_tasks = (
        snapshot["completed_tasks"]
        + snapshot["failed_tasks"]
    )

    slow_requests = sum(
        value >= SLOW_REQUEST_MS
        for value in durations
    )

    slow_inflight = sum(
        worker.get("current_request_ms") is not None
        and worker["current_request_ms"] >= SLOW_REQUEST_MS
        for worker in workers
    )

    recent_window = min(
        elapsed,
        RECENT_WINDOW_SECONDS,
    )

    phases = {
        phase: sum(
            worker.get("phase") == phase
            for worker in workers
        )
        for phase in (
            "REQUESTING",
            "AUTHENTICATING",
            "PARSING",
            "QUEUED",
            "COOLDOWN",
            "RETRYING",
            "DELAY",
            "FINISHED",
        )
    }

    timing = {
        name: average_metric(name)
        for name in (
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
    }

    return {
        "elapsed": elapsed,

        "total_tasks": snapshot["total_tasks"],

        "finished_tasks": finished_tasks,

        "completed_tasks": snapshot["completed_tasks"],

        "failed_tasks": snapshot["failed_tasks"],

        "data_chunks": snapshot["data_chunks"],

        "empty_pages": snapshot["empty_pages"],

        "no_cell_chunks": snapshot["no_cell_chunks"],

        "requests": len(requests),

        "successful_requests": ok_requests,

        "failed_requests": failed_requests,

        "request_rate": round(
            len(requests) / elapsed,
            3,
        ),

        "recent_request_rate": round(
            len(recent) / max(
                recent_window,
                0.001,
            ),
            3,
        ),

        "recent_samples": len(recent),

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
            durations,
            95,
        ),

        "max_ms": (
            round(max(durations), 3)
            if durations
            else None
        ),

        "recent_average_ms": mean_or_none(
            recent_durations
        ),

        "recent_p95_ms": percentile(
            recent_durations,
            95,
        ),

        "slow_requests": slow_requests,

        "slow_inflight_workers": slow_inflight,

        "cells": snapshot["cells"],

        "non_null_cells": snapshot["non_null_cells"],

        "cells_per_second": round(
            snapshot["non_null_cells"]
            / elapsed,
            3,
        ),

        "response_bytes": snapshot["response_bytes"],

        "bytes_per_second": round(
            snapshot["response_bytes"]
            / elapsed,
            3,
        ),

        "new_connections": (
            connections_total
            if known_connections
            else None
        ),

        "connection_reuse_percent": (
            round(reuse_percent, 2)
            if reuse_percent is not None
            else None
        ),

        "connection_samples": len(
            known_connections
        ),

        "max_worker_chunks": max_completed,

        "phases": phases,

        "timing": timing,

        "workers": workers,

        "events": snapshot["events"],

        "fatal_error": snapshot["fatal_error"],
    }


# ============================================================
# REQUEST WORKER
# ============================================================

def request_worker(
    worker_id: int,
    query: QuerySpec,
    profile: Profile,
    tasks: queue.Queue,
    results: queue.Queue,
    no_more_tasks: threading.Event,
    profile_stop: threading.Event,
    state: EngineState,
    auth: AuthenticationManager,
    control: RequestControl,
    profile_dir: Path,
    *,
    capture_values: bool,
):

    session = None

    last_request_at = 0.0

    try:

        session = create_session(
            auth.cookie
        )

        state.update_worker(
            worker_id,
            status="RUNNING",
            phase="QUEUED",
            last_event="Persistent HTTP session ready",
        )

        while not (
            STOP_EVENT.is_set()
            or profile_stop.is_set()
        ):

            try:

                task = tasks.get(
                    timeout=0.25
                )

            except queue.Empty:

                if no_more_tasks.is_set():

                    break

                continue

            task_started = time.perf_counter()

            queue_wait_ms = round(
                (
                    task_started
                    - task.queued_at
                ) * 1000,
                3,
            )

            final_result = None
            final_error = None

            try:

                request_body = build_soap(
                    build_mdx(
                        query,
                        task,
                    )
                )

                rejected_token = None

                destination = (
                    profile_dir
                    / "responses"
                    / f"{task.request_id}.xml"
                )

                for attempt in range(
                    1,
                    MAX_ATTEMPTS + 1,
                ):

                    check_stop()

                    if profile_stop.is_set():

                        raise StopRequested(
                            "Profile stopped"
                        )

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

                    token, lock_wait, refresh_ms = (
                        auth.get_token(
                            session,
                            rejected_token=rejected_token,
                        )
                    )

                    rejected_token = None

                    local_wait = max(
                        0.0,
                        last_request_at
                        + WORKER_REQUEST_DELAY
                        - time.monotonic(),
                    )

                    if local_wait:

                        state.update_worker(
                            worker_id,
                            phase="DELAY",
                        )

                        interruptible_sleep(
                            local_wait
                        )

                    if control.remaining() > 0:

                        state.update_worker(
                            worker_id,
                            status="PAUSED",
                            phase="COOLDOWN",
                            last_event="Shared server cooldown",
                        )

                    rate_wait = control.wait_slot()

                    if profile_stop.is_set():

                        raise StopRequested(
                            "Profile stopped"
                        )

                    request_started = (
                        time.perf_counter()
                    )

                    last_request_at = (
                        time.monotonic()
                    )

                    record = {
                        "profile": profile.name,

                        "engine": profile.engine,

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
                            else 0
                        ),

                        "rate_wait_ms": (
                            rate_wait
                            + local_wait * 1000
                        ),

                        "token_lock_wait_ms": lock_wait,

                        "token_refresh_ms": refresh_ms,
                    }

                    state.update_worker(
                        worker_id,
                        status="RUNNING",
                        phase="REQUESTING",
                        request_started=(
                            time.monotonic()
                        ),
                        last_event=(
                            f"REQUEST X={task.x_offset} "
                            f"Y={task.y_offset}"
                        ),
                    )

                    retry = False
                    fatal = False

                    try:

                        response = execute_request(
                            session,
                            token,
                            request_body,
                            destination,
                            profile.engine,
                        )

                        record.update(
                            response
                        )

                        status = record[
                            "http_status"
                        ]

                        state.update_worker(
                            worker_id,
                            request_started=None,
                            last_http=status,
                            last_request_ms=record.get(
                                "execute_ms"
                            ),
                            last_ttfb_ms=record.get(
                                "ttfb_ms"
                            ),
                        )

                        if status == 200:

                            state.update_worker(
                                worker_id,
                                phase="PARSING",
                            )

                            parsed = parse_response(
                                destination,
                                query,
                                task,
                                capture_values=capture_values,
                            )

                            comparison_values = parsed.pop(
                                "comparison_values"
                            )

                            record.update(parsed)

                            record["status"] = "OK"

                            final_result = {
                                **record,
                                "comparison_values": (
                                    comparison_values
                                ),
                            }

                        elif status in (401, 403):

                            record["status"] = "ERROR"

                            record["error"] = (
                                f"Authentication HTTP {status}"
                            )

                            if attempt < MAX_ATTEMPTS:

                                rejected_token = token

                                retry = True

                            else:

                                fatal = True

                        elif status in (429, 503):

                            record["status"] = "ERROR"

                            record["error"] = (
                                f"Server HTTP {status}"
                            )

                            cooldown = retry_after_seconds(
                                response.get(
                                    "retry_after"
                                ),
                                default=15.0 * attempt,
                            )

                            control.set_cooldown(
                                cooldown,
                                f"HTTP {status}",
                            )

                            retry = (
                                attempt < MAX_ATTEMPTS
                            )

                            if not retry:

                                fatal = True

                        elif status in (
                            408,
                            425,
                            500,
                            502,
                            504,
                        ):

                            record["status"] = "ERROR"

                            record["error"] = (
                                f"Transient HTTP {status}"
                            )

                            retry = (
                                attempt < MAX_ATTEMPTS
                            )

                        else:

                            record["status"] = "ERROR"

                            record["error"] = (
                                f"HTTP {status}"
                            )

                    except StopRequested:

                        record["status"] = "STOPPED"

                        record["error"] = "Stopped"

                        raise

                    except RequestException as exc:

                        record["status"] = "ERROR"

                        record["error"] = safe_error(
                            exc
                        )

                        retry = (
                            attempt < MAX_ATTEMPTS
                        )

                    except (
                        ResponseError,
                        ResponseTooLarge,
                    ) as exc:

                        record["status"] = "ERROR"

                        record["error"] = safe_error(
                            exc
                        )

                        # Invalid datasets and oversize
                        # responses are not blindly retried.
                        retry = False

                    finally:

                        record.setdefault(
                            "execute_ms",
                            elapsed_ms(
                                request_started
                            ),
                        )

                        record["completed_at"] = (
                            utc_now()
                        )

                        record["task_total_ms"] = round(
                            (
                                time.perf_counter()
                                - task.queued_at
                            ) * 1000,
                            3,
                        )

                        record.setdefault(
                            "status",
                            "ERROR",
                        )

                        state.add_request(
                            record
                        )

                        results.put((
                            "REQUEST",
                            dict(record),
                        ))

                        state.update_worker(
                            worker_id,
                            request_started=None,
                        )

                    if final_result is not None:

                        duration = final_result.get(
                            "execute_ms",
                            0,
                        )

                        kind = final_result.get(
                            "chunk_kind",
                            "DATA",
                        )

                        prefix = (
                            "SLOW"
                            if duration >= SLOW_REQUEST_MS
                            else "OK"
                        )

                        state.update_worker(
                            worker_id,
                            status="RUNNING",
                            phase="QUEUED",
                            last_event=(
                                f"{prefix} HTTP 200 | "
                                f"{format_ms(duration)} | "
                                f"{kind} | "
                                f"X={final_result['x_members']} "
                                f"Y={final_result['y_members']} "
                                f"cells={final_result['cells']}"
                            ),
                        )

                        break

                    final_error = record.get(
                        "error",
                        "Request failed",
                    )

                    if fatal:

                        state.set_fatal(
                            final_error
                        )

                        profile_stop.set()

                        break

                    if not retry:

                        break

                    state.update_worker(
                        worker_id,
                        status="RETRYING",
                        phase="RETRYING",
                        last_event=(
                            f"Retry {attempt + 1}/"
                            f"{MAX_ATTEMPTS} | "
                            f"{final_error}"
                        ),
                    )

                    interruptible_sleep(
                        min(
                            10,
                            2 ** attempt,
                        )
                    )

                state.finish_task(
                    worker_id,
                    final_result,
                )

                if final_result is not None:

                    results.put((
                        "RESULT",
                        final_result,
                    ))

                else:

                    state.update_worker(
                        worker_id,
                        last_event=(
                            "FAILED | "
                            + (
                                final_error
                                or "No result"
                            )
                        ),
                    )

            except StopRequested:

                break

            except AuthenticationError as exc:

                state.set_fatal(
                    safe_error(exc)
                )

                state.finish_task(
                    worker_id,
                    None,
                )

                state.update_worker(
                    worker_id,
                    status="FAILED",
                    phase="FINISHED",
                    last_event=safe_error(exc),
                )

                profile_stop.set()

                break

            except Exception as exc:

                state.finish_task(
                    worker_id,
                    None,
                )

                state.set_fatal(
                    safe_error(exc)
                )

                state.update_worker(
                    worker_id,
                    status="FAILED",
                    phase="FINISHED",
                    last_event=safe_error(exc),
                )

                profile_stop.set()

                break

            finally:

                tasks.task_done()

        current = state.snapshot()

        worker = next(
            item
            for item in current["workers"]
            if item["worker_id"] == worker_id
        )

        if STOP_EVENT.is_set():

            status = "STOPPED"

        elif worker["failed"]:

            status = (
                "PARTIAL"
                if worker["completed"]
                else "FAILED"
            )

        elif profile_stop.is_set():

            status = "STOPPED"

        else:

            status = "COMPLETED"

        state.update_worker(
            worker_id,
            status=status,
            phase="FINISHED",
            request_started=None,
            last_event=status,
        )

    except Exception as exc:

        state.set_fatal(
            safe_error(exc)
        )

        state.update_worker(
            worker_id,
            status="FAILED",
            phase="FINISHED",
            request_started=None,
            last_event=safe_error(exc),
        )

        profile_stop.set()

    finally:

        if session is not None:

            session.close()


# ============================================================
# RICH LIVE MONITOR
# ============================================================

def status_color(status: str) -> str:

    return {
        "STARTING": "dim",
        "RUNNING": "bright_cyan",
        "RETRYING": "yellow",
        "PAUSED": "yellow",
        "COMPLETED": "green",
        "PARTIAL": "yellow",
        "FAILED": "red",
        "STOPPED": "dim",
    }.get(
        status,
        "white",
    )


def latency_color(value) -> str:

    if value is None:

        return "dim"

    if value >= 15000:

        return "bright_red"

    if value >= SLOW_REQUEST_MS:

        return "yellow"

    if value < 1000:

        return "green"

    return "white"


def make_panel(
    content,
    title: str,
):

    return Panel(
        content,
        title=title,
        border_style="cyan",
        padding=(0, 1),
    )


def kv_table(
    values: list[tuple],
):

    table = Table.grid(
        expand=True
    )

    table.add_column(
        style="dim",
        ratio=1,
    )

    table.add_column(
        justify="right",
        ratio=2,
    )

    for key, value in values:

        table.add_row(
            Text(str(key)),
            value if isinstance(value, Text)
            else Text(str(value)),
        )

    return table


def render_monitor(
    query: QuerySpec,
    profile: Profile,
    metrics: dict,
    control: RequestControl,
    auth: AuthenticationManager,
    *,
    mode: str,
    x_total: int,
    y_total: int,
    benchmark_results: list[dict],
):

    width = console.size.width
    height = console.size.height

    workers = metrics["workers"]

    completed = metrics["completed_tasks"]

    failed = metrics["failed_tasks"]

    task_total = metrics["total_tasks"]

    active = sum(
        worker["status"] in (
            "RUNNING",
            "RETRYING",
            "PAUSED",
        )
        for worker in workers
    )

    timing = metrics["timing"]

    header = Panel(
        Text(
            APP_NAME
            + "\n"
            + "MV / FY | XMLA / MDX | "
            + "CURL_CFFI PERFORMANCE LAB",
            justify="center",
            style="bold bright_cyan",
        ),
        border_style="cyan",
    )

    overview = kv_table([
        ("Mode", mode.upper()),

        ("Profile", profile.name),

        ("Engine", profile.engine.upper()),

        ("Cookie", Text(
            "VALID",
            style="green",
        )),

        ("JWT", Text(
            "VALID"
            if jwt_exp(auth.token) > time.time()
            else "REFRESH REQUIRED",
            style=(
                "green"
                if jwt_exp(auth.token) > time.time()
                else "yellow"
            ),
        )),

        ("Workers", f"{active}/{profile.workers}"),

        ("Elapsed", format_duration(
            metrics["elapsed"]
        )),
    ])

    configuration = kv_table([
        ("Category", "MV"),

        ("Cube", "FY"),

        ("X1", query.x_captions[0]),

        ("X2", query.x_captions[1]),

        ("Y1", query.y_captions[0]),

        ("Measures", len(query.measures)),

        (
            "Chunk X/Y",
            f"{profile.x_chunk}/{profile.y_chunk}",
        ),

        (
            "Scope X/Y",
            f"{x_total:,}/{y_total:,}",
        ),
    ])

    top = Table.grid(
        expand=True
    )

    top.add_column(
        ratio=1
    )

    top.add_column(
        ratio=1
    )

    top.add_row(
        make_panel(
            overview,
            "01 / ENGINE OVERVIEW",
        ),

        make_panel(
            configuration,
            "02 / FIXED QUERY",
        ),
    )

    progress_lines = []

    progress_lines.append(
        Text.assemble(
            (
                f"{completed:,}/{task_total:,} "
                "COMPLETED",
                "bold green",
            ),

            " | ",

            (
                f"{failed:,} FAILED",
                "red" if failed else "dim",
            ),
        )
    )

    progress_lines.append(
        Bar(
            size=max(
                1,
                task_total,
            ),
            begin=0,
            end=completed,
            color="cyan",
            bgcolor="grey23",
        )
    )

    progress_percent = (
        completed / task_total * 100
        if task_total
        else 100
    )

    progress_lines.append(
        Text(
            f"Completion: {progress_percent:.2f}%"
            f" | Data: {metrics['data_chunks']:,}"
            f" | Empty: {metrics['empty_pages']:,}"
            f" | No cells: {metrics['no_cell_chunks']:,}",
            style="dim",
        )
    )

    progress_panel = make_panel(
        Group(*progress_lines),
        "03 / EXTRACTION PROGRESS",
    )

    latency = metrics[
        "average_ms"
    ]

    performance = kv_table([
        ("Requests", f"{metrics['requests']:,}"),

        ("Requests/sec", (
            f"{metrics['request_rate']:.2f}"
        )),

        ("Cells/sec", Text(
            f"{metrics['cells_per_second']:.2f}",
            style="bold green",
        )),

        ("Average latency", Text(
            format_ms(latency),
            style=latency_color(latency),
        )),

        ("Median latency", format_ms(
            metrics["median_ms"]
        )),

        ("P95 latency", Text(
            format_ms(metrics["p95_ms"]),
            style=latency_color(
                metrics["p95_ms"]
            ),
        )),

        ("Recent average", format_ms(
            metrics["recent_average_ms"]
        )),

        ("Recent P95", format_ms(
            metrics["recent_p95_ms"]
        )),

        ("Slow requests", metrics[
            "slow_requests"
        ]),

        ("Non-null cells", f"{metrics['non_null_cells']:,}"),

        ("Saved XML", format_bytes(
            metrics["response_bytes"]
        )),

        ("Cooldown", (
            f"{control.remaining():.1f}s"
            if control.remaining()
            else "NONE"
        )),
    ])

    timing_table = Table.grid(
        expand=True
    )

    timing_table.add_column(
        style="dim",
        ratio=2,
    )

    timing_table.add_column(
        justify="right",
        ratio=1,
    )

    timing_rows = [
        ("Queue wait", "queue_wait_ms"),

        ("Rate wait", "rate_wait_ms"),

        ("Token refresh", "token_refresh_ms"),

        ("DNS lookup", "dns_ms"),

        ("TCP connect", "connect_ms"),

        ("TLS handshake", "tls_ms"),

        ("TTFB", "ttfb_ms"),

        (
            "Response wait estimate",
            "response_wait_estimate_ms",
        ),

        ("Receive", "receive_ms"),

        ("File write", "file_write_ms"),

        ("Finalize XML", "finalize_ms"),

        ("Parse XML", "parse_ms"),
    ]

    for label, key in timing_rows:

        value = timing.get(key)

        timing_table.add_row(
            Text(label),

            Text(
                format_ms(value),
                style=latency_color(value),
            ),
        )

    metric_grid = Table.grid(
        expand=True
    )

    metric_grid.add_column(
        ratio=1
    )

    metric_grid.add_column(
        ratio=1
    )

    metric_grid.add_row(
        make_panel(
            performance,
            "04 / PERFORMANCE",
        ),

        make_panel(
            timing_table,
            "05 / TIMING BREAKDOWN",
        ),
    )

    worker_blocks = []

    normal_mode = (
        height >= (
            42 + len(workers) * 4
        )
        and width >= 90
    )

    for worker in workers:

        worker_id = worker["worker_id"]

        status = worker["status"]

        color = status_color(status)

        last_ms = worker.get(
            "last_request_ms"
        )

        current_ms = worker.get(
            "current_request_ms"
        )

        relative = worker.get(
            "relative_progress",
            0,
        )

        x_offset = worker.get(
            "x_offset"
        )

        y_offset = worker.get(
            "y_offset"
        )

        offset = (
            f"{x_offset}/{y_offset}"
            if x_offset is not None
            else "—"
        )

        last_http = worker.get(
            "last_http"
        )

        left = Text.assemble(
            (
                f"W{worker_id:02d}  ",
                "bold white",
            ),
            (
                status,
                color,
            ),
        )

        right = Text(
            f"Last {format_ms(last_ms)}",
            style=latency_color(last_ms),
        )

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
            left,
            right,
        )

        bar = Bar(
            size=1000,
            begin=0,
            end=int(
                relative * 1000
            ),
            color=color,
            bgcolor="grey23",
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
                f"Offset {offset}"
                f" | {worker['phase']}",
                style="dim",
            ),

            Text(
                f"{worker['completed']:,} chunks"
                f" | {worker['failed']:,} failed",
                style=(
                    "yellow"
                    if worker["failed"]
                    else "white"
                ),
            ),
        )

        extra = Text.assemble(
            (
                f"HTTP {last_http or '—'}",
                "green"
                if last_http == 200
                else "dim",
            ),

            f" | TTFB "
            f"{format_ms(worker.get('last_ttfb_ms'))}",

            (
                f" | Current {format_ms(current_ms)}",
                latency_color(current_ms),
            ),

            f" | Cells {worker['cells']:,}",
        )

        if normal_mode:

            worker_blocks.extend([
                row,
                bar,
                details,
                extra,
                Text(""),
            ])

        else:

            compact = Table.grid(
                expand=True
            )

            compact.add_column(
                width=5
            )

            compact.add_column(
                width=10
            )

            compact.add_column(
                ratio=1
            )

            compact.add_column(
                justify="right"
            )

            compact.add_row(
                Text(
                    f"W{worker_id:02d}",
                    style="bold",
                ),

                Text(
                    status,
                    style=color,
                ),

                bar,

                Text(
                    f"{worker['completed']:,}"
                    f"/{worker['failed']:,}"
                    f" | {format_ms(last_ms)}"
                    f" | Now {format_ms(current_ms)}",
                    style=latency_color(last_ms),
                ),
            )

            worker_blocks.append(
                compact
            )

    worker_blocks.append(
        Text(
            "Relative successful chunks versus leading worker.",
            style="dim",
        )
    )

    worker_panel = make_panel(
        Group(*worker_blocks),
        "06 / PARALLEL WORKERS",
    )

    alerts = []

    if metrics["slow_inflight_workers"]:

        alerts.append(
            Text(
                "SLOW IN-FLIGHT REQUESTS: "
                f"{metrics['slow_inflight_workers']}",
                style="yellow",
            )
        )

    recent_p95 = metrics[
        "recent_p95_ms"
    ]

    if (
        recent_p95 is not None
        and recent_p95 >= SLOW_REQUEST_MS
    ):

        alerts.append(
            Text(
                "HIGH TAIL LATENCY | "
                f"P95 {format_ms(recent_p95)}",
                style="yellow",
            )
        )

    if control.remaining() > 0:

        alerts.append(
            Text(
                "SERVER COOLDOWN | "
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
        "07 / PERFORMANCE ALERTS",
    )

    benchmark_table = Table(
        box=box.SIMPLE,
        expand=True,
    )

    for name in (
        "Profile",
        "W",
        "Chunk",
        "Elapsed",
        "P95",
        "Cells/s",
        "Match",
    ):

        benchmark_table.add_column(
            name,
            no_wrap=True,
        )

    for result in benchmark_results:

        equivalent = result.get(
            "equivalent"
        )

        benchmark_table.add_row(
            result["profile"],

            str(result["workers_count"]),

            (
                f"{result['x_chunk']}/"
                f"{result['y_chunk']}"
            ),

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

    activity = []

    for event in metrics["events"][-4:]:

        activity.append(
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
        )

    if not activity:

        activity.append(
            Text(
                "Waiting for workers...",
                style="dim",
            )
        )

    sections = [
        header,
        top,
        progress_panel,
        metric_grid,
        worker_panel,
        alert_panel,
    ]

    if benchmark_results:

        sections.append(
            make_panel(
                benchmark_table,
                "08 / BENCHMARK RESULTS",
            )
        )

    if height >= 42:

        sections.append(
            make_panel(
                Group(*activity),
                "LATEST ACTIVITY",
            )
        )

    sections.append(
        Text(
            "LOCAL TIME | "
            "REFRESH 1s | "
            "Ctrl+C TO STOP",
            style="dim",
        )
    )

    return Group(*sections)


# ============================================================
# DYNAMIC REQUEST SCHEDULER
# ============================================================

def generate_tasks(
    query: QuerySpec,
    profile: Profile,
    x_total: int,
    y_total: int,
):

    for x_offset in range(
        0,
        x_total,
        profile.x_chunk,
    ):

        x_size = min(
            profile.x_chunk,
            x_total - x_offset,
        )

        for y_offset in range(
            0,
            y_total,
            profile.y_chunk,
        ):

            y_size = min(
                profile.y_chunk,
                y_total - y_offset,
            )

            yield (
                x_offset,
                y_offset,
                x_size,
                y_size,
            )


def run_scheduler(
    query: QuerySpec,
    profile: Profile,
    x_total: int,
    y_total: int,
    auth: AuthenticationManager,
    control: RequestControl,
    run_dir: Path,
    csv_writer,
    csv_handle,
    *,
    mode: str,
    capture_values: bool,
    benchmark_results: list[dict],
) -> tuple[dict, dict]:

    task_total = (
        math.ceil(
            x_total / profile.x_chunk
        )
        * math.ceil(
            y_total / profile.y_chunk
        )
    )

    if task_total > MAX_TASKS_PER_RUN:

        raise EngineError(
            f"Planned tasks exceed local limit: "
            f"{task_total:,}"
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

    tasks = queue.Queue(
        maxsize=max(
            2,
            profile.workers * 2,
        )
    )

    results = queue.Queue()

    no_more_tasks = threading.Event()

    profile_stop = threading.Event()

    generator = iter(
        generate_tasks(
            query,
            profile,
            x_total,
            y_total,
        )
    )

    comparison_values = {}

    duplicate_keys = 0

    data_updates = set()

    threads = []

    def drain_results():

        nonlocal duplicate_keys

        while True:

            try:

                kind, record = (
                    results.get_nowait()
                )

            except queue.Empty:

                break

            if kind == "REQUEST":

                csv_writer.writerow({
                    name: record.get(name)
                    for name in CSV_FIELDS
                })

            elif kind == "RESULT":

                update = record.get(
                    "last_data_update"
                )

                if update:

                    data_updates.add(update)

                if capture_values:

                    values = record.get(
                        "comparison_values",
                        {},
                    )

                    for key, value in values.items():

                        if key in comparison_values:

                            duplicate_keys += 1

                        comparison_values[key] = value

        csv_handle.flush()

    for worker_id in range(
        1,
        profile.workers + 1,
    ):

        thread = threading.Thread(
            target=request_worker,

            name=f"W{worker_id:02d}",

            args=(
                worker_id,
                query,
                profile,
                tasks,
                results,
                no_more_tasks,
                profile_stop,
                state,
                auth,
                control,
                profile_dir,
            ),

            kwargs={
                "capture_values": capture_values,
            },

            daemon=False,
        )

        threads.append(thread)

        thread.start()

    try:

        with Live(
            console=console,
            refresh_per_second=1,
            auto_refresh=False,
            screen=True,
            vertical_overflow="crop",
        ) as live:

            next_refresh = 0.0

            while True:

                if (
                    STOP_EVENT.is_set()
                    or profile_stop.is_set()
                ):

                    no_more_tasks.set()

                if not no_more_tasks.is_set():

                    while not tasks.full():

                        try:

                            (
                                x_offset,
                                y_offset,
                                x_size,
                                y_size,
                            ) = next(generator)

                        except StopIteration:

                            no_more_tasks.set()

                            break

                        request_id = (
                            f"{query.query_hash}"
                            f"_x{x_offset:06d}"
                            f"_y{y_offset:06d}"
                            f"_xs{x_size:04d}"
                            f"_ys{y_size:04d}"
                        )

                        task = RequestTask(
                            request_id=request_id,

                            x_offset=x_offset,

                            y_offset=y_offset,

                            x_chunk=x_size,

                            y_chunk=y_size,

                            queued_at=time.perf_counter(),
                        )

                        tasks.put_nowait(
                            task
                        )

                drain_results()

                now = time.monotonic()

                if now >= next_refresh:

                    metrics = aggregate_metrics(
                        state.snapshot()
                    )

                    live.update(
                        render_monitor(
                            query,
                            profile,
                            metrics,
                            control,
                            auth,
                            mode=mode,
                            x_total=x_total,
                            y_total=y_total,
                            benchmark_results=benchmark_results,
                        ),
                        refresh=True,
                    )

                    next_refresh = (
                        now
                        + REFRESH_INTERVAL
                    )

                if not any(
                    thread.is_alive()
                    for thread in threads
                ):

                    break

                time.sleep(0.1)

            drain_results()

            live.update(
                render_monitor(
                    query,
                    profile,
                    aggregate_metrics(
                        state.snapshot()
                    ),
                    control,
                    auth,
                    mode=mode,
                    x_total=x_total,
                    y_total=y_total,
                    benchmark_results=benchmark_results,
                ),
                refresh=True,
            )

    except KeyboardInterrupt:

        STOP_EVENT.set()

        profile_stop.set()

        no_more_tasks.set()

    finally:

        if (
            STOP_EVENT.is_set()
            or profile_stop.is_set()
        ):

            no_more_tasks.set()

        for thread in threads:

            thread.join()

        drain_results()

    metrics = aggregate_metrics(
        state.snapshot()
    )

    status = "COMPLETED"

    if STOP_EVENT.is_set():

        status = "STOPPED"

    elif (
        metrics["failed_tasks"]
        or metrics["fatal_error"]
        or metrics["completed_tasks"] != task_total
        or duplicate_keys
        or len(data_updates) > 1
    ):

        status = (
            "PARTIAL"
            if metrics["completed_tasks"]
            else "FAILED"
        )

    report = {
        **metrics,

        "status": status,

        "profile": profile.name,

        "engine": profile.engine,

        "workers_count": profile.workers,

        "x_chunk": profile.x_chunk,
        "y_chunk": profile.y_chunk,

        "x_total": x_total,
        "y_total": y_total,

        "duplicate_keys": duplicate_keys,

        "last_data_updates": sorted(
            data_updates
        ),

        "comparison_coordinates": len(
            comparison_values
        ),

        "completed_at": utc_now(),
    }

    write_json(
        profile_dir / "summary.json",
        report,
    )

    return report, comparison_values


# ============================================================
# BOTTLENECK ANALYSIS
# ============================================================

def identify_bottleneck(
    result: dict,
) -> dict:

    timing = result.get(
        "timing",
        {},
    )

    execution = timing.get(
        "execute_ms"
    )

    if not execution:

        return {
            "name": "UNKNOWN",
            "message": (
                "Insufficient completed request measurements"
            ),
        }

    measured_wait = timing.get(
        "response_wait_estimate_ms"
    )

    ttfb = timing.get(
        "ttfb_ms"
    )

    receive = timing.get(
        "receive_ms"
    )

    write = timing.get(
        "file_write_ms"
    )

    parse = timing.get(
        "parse_ms"
    )

    token = timing.get(
        "token_refresh_ms"
    )

    candidates = {
        name: value
        for name, value in {
            "RESPONSE_WAIT": (
                measured_wait
                if measured_wait is not None
                else ttfb
            ),

            "RESPONSE_RECEIVE": receive,

            "FILE_WRITE": write,

            "XML_PARSE": parse,

            "TOKEN_REFRESH": token,
        }.items()
        if value is not None
    }

    if not candidates:

        return {
            "name": "UNKNOWN",
            "message": "Detailed timings unavailable",
        }

    name, duration = max(
        candidates.items(),
        key=lambda item: item[1],
    )

    descriptions = {
        "RESPONSE_WAIT": (
            "Most measured time is spent waiting for "
            "response data. Upstream processing, network "
            "transit and server queuing remain possible. "
            "This does not prove server throttling."
        ),

        "RESPONSE_RECEIVE": (
            "Response body handling is a major measured cost. "
            "Check transfer speed, response size and writes."
        ),

        "FILE_WRITE": (
            "Local response file writing is a major measured cost."
        ),

        "XML_PARSE": (
            "XML parsing is a major measured local cost."
        ),

        "TOKEN_REFRESH": (
            "Authentication refresh is a major measured cost."
        ),
    }

    return {
        "name": name,
        "average_ms": round(duration, 3),
        "message": descriptions[name],
        "note": (
            "Timing stages may overlap. "
            "Average timings are not an exact additive breakdown."
        ),
    }


# ============================================================
# BENCHMARK PROFILES
# ============================================================

def build_benchmark_profiles(
    max_workers: int,
    x_total: int,
    y_total: int,
) -> list[Profile]:

    profiles = [
        Profile(
            name="01_BASELINE_ITERATOR",
            engine="iterator",
            workers=1,
            x_chunk=1,
            y_chunk=1,
        ),

        Profile(
            name="02_NATIVE_CALLBACK",
            engine="callback",
            workers=1,
            x_chunk=1,
            y_chunk=1,
        ),
    ]

    if max_workers >= 2:

        profiles.append(
            Profile(
                name="03_PARALLEL_2",
                engine="callback",
                workers=2,
                x_chunk=1,
                y_chunk=1,
            )
        )

    if max_workers > 2:

        profiles.append(
            Profile(
                name="04_PARALLEL_MAX",
                engine="callback",
                workers=max_workers,
                x_chunk=1,
                y_chunk=1,
            )
        )

    profiles.extend([
        Profile(
            name="05_CHUNK_SMALL",
            engine="callback",
            workers=max_workers,
            x_chunk=min(5, x_total),
            y_chunk=min(5, y_total),
        ),

        Profile(
            name="06_CHUNK_BALANCED",
            engine="callback",
            workers=max_workers,
            x_chunk=min(10, x_total),
            y_chunk=min(10, y_total),
        ),
    ])

    unique = []

    seen = set()

    for profile in profiles:

        key = (
            profile.engine,
            profile.workers,
            profile.x_chunk,
            profile.y_chunk,
        )

        if key in seen:

            continue

        seen.add(key)

        unique.append(profile)

    return unique


def result_digest(
    values: dict,
) -> str:

    digest = hashlib.sha256()

    for key in sorted(values):

        encoded = json.dumps(
            [
                key,
                values[key],
            ],
            ensure_ascii=False,
            separators=(",", ":"),
        ).encode("utf-8")

        digest.update(encoded)

        digest.update(b"\n")

    return digest.hexdigest()


def run_benchmark(
    query: QuerySpec,
    x_total: int,
    y_total: int,
    max_workers: int,
    auth: AuthenticationManager,
    control: RequestControl,
    run_dir: Path,
    csv_writer,
    csv_handle,
) -> dict:

    sample_x = min(
        BENCHMARK_X_MEMBERS,
        x_total,
    )

    sample_y = min(
        BENCHMARK_Y_MEMBERS,
        y_total,
    )

    profiles = build_benchmark_profiles(
        max_workers,
        sample_x,
        sample_y,
    )

    reports = []

    baseline_values = None

    baseline_report = None

    expected_coordinates = (
        sample_x
        * sample_y
        * len(query.measures)
    )

    for profile in profiles:

        if STOP_EVENT.is_set():

            break

        report, values = run_scheduler(
            query,
            profile,
            sample_x,
            sample_y,
            auth,
            control,
            run_dir,
            csv_writer,
            csv_handle,
            mode="benchmark",
            capture_values=True,
            benchmark_results=reports,
        )

        complete = (
            report["status"] == "COMPLETED"
            and len(values) == expected_coordinates
            and not report["duplicate_keys"]
        )

        if baseline_values is None:

            if not complete:

                report["equivalent"] = None

                reports.append(report)

                break

            baseline_values = dict(values)

            baseline_report = dict(report)

            report["equivalent"] = True

        else:

            baseline_updates = set(
                baseline_report.get(
                    "last_data_updates",
                    [],
                )
            )

            current_updates = set(
                report.get(
                    "last_data_updates",
                    [],
                )
            )

            same_snapshot = (
                not baseline_updates
                or not current_updates
                or baseline_updates == current_updates
            )

            report["equivalent"] = (
                complete
                and same_snapshot
                and values == baseline_values
            )

        report["dataset_digest"] = (
            result_digest(values)
            if complete
            else None
        )

        report["expected_coordinates"] = (
            expected_coordinates
        )

        if baseline_report:

            report["speedup"] = round(
                baseline_report["elapsed"]
                / max(
                    0.001,
                    report["elapsed"],
                ),
                3,
            )

        report["bottleneck"] = (
            identify_bottleneck(report)
        )

        reports.append(report)

        write_json(
            run_dir
            / profile.name
            / "summary.json",
            report,
        )

        write_json(
            run_dir / "benchmark_progress.json",
            {
                "query_hash": query.query_hash,
                "profiles": reports,
            },
        )

        # Do not continue increasing load after
        # authentication errors, repeated server errors
        # or invalid/incomplete results.
        if (
            report["status"] != "COMPLETED"
            or report.get("fatal_error")
            or report["equivalent"] is False
        ):

            break

        interruptible_sleep(2)

    eligible = [
        report
        for report in reports
        if (
            report["status"] == "COMPLETED"
            and report.get("equivalent") is True
            and report["failed_tasks"] == 0
            and report["failed_requests"] == 0
        )
    ]

    recommendation = None

    if eligible:

        best = min(
            eligible,
            key=lambda report: (
                report["elapsed"],
                report.get("p95_ms")
                or float("inf"),
                report["workers_count"],
            ),
        )

        recommendation = {
            "profile": best["profile"],

            "engine": best["engine"],

            "workers": best["workers_count"],

            "x_chunk": best["x_chunk"],

            "y_chunk": best["y_chunk"],

            "sample_speedup": best.get(
                "speedup",
                1.0,
            ),

            "sample_cells_per_second": best[
                "cells_per_second"
            ],

            "sample_p95_ms": best.get(
                "p95_ms"
            ),

            "sample_elapsed": best[
                "elapsed"
            ],

            "bottleneck": best.get(
                "bottleneck"
            ),
        }

    return {
        "mode": "benchmark",

        "query_hash": query.query_hash,

        "status": (
            "STOPPED"
            if STOP_EVENT.is_set()
            else "COMPLETED"
            if len(reports) == len(profiles)
            and all(
                report.get("equivalent") is True
                for report in reports
            )
            else "PARTIAL"
        ),

        "benchmark_scope": {
            "x_members": sample_x,

            "y_members": sample_y,

            "measures": len(
                query.measures
            ),

            "expected_coordinates": (
                expected_coordinates
            ),
        },

        "profiles": reports,

        "recommended_profile": recommendation,

        "limitations": [
            "Profiles use the same bounded dataset.",
            "Results must match before performance comparison.",
            "Sequential profiles may benefit from server cache warming.",
            "Small benchmark results do not guarantee full-run speedup.",
            "No request-rate limits are bypassed.",
            "Raw measure values are not converted to millions.",
        ],
    }


# ============================================================
# SELECT VERIFIED BENCHMARK CONFIGURATION
# ============================================================

def load_recommended_profile(
    query: QuerySpec,
    max_workers: int,
) -> Profile:

    summaries = sorted(
        OUTPUT_ROOT.glob(
            "*/summary.json"
        ),
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )

    for path in summaries:

        report = read_json(path)

        if report.get("mode") != "benchmark":

            continue

        if report.get("status") != "COMPLETED":

            continue

        if report.get("query_hash") != query.query_hash:

            continue

        recommendation = report.get(
            "recommended_profile"
        ) or {}

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

            engine = recommendation[
                "engine"
            ]

        except (
            KeyError,
            TypeError,
            ValueError,
        ):

            continue

        if not 1 <= workers <= max_workers:

            continue

        if not 1 <= x_chunk <= 50:

            continue

        if not 1 <= y_chunk <= 20:

            continue

        if engine not in (
            "iterator",
            "callback",
        ):

            continue

        return Profile(
            name="FULL_EXTRACTION",
            engine=engine,
            workers=workers,
            x_chunk=x_chunk,
            y_chunk=y_chunk,
        )

    return Profile(
        name="FULL_EXTRACTION",
        engine="callback",
        workers=max_workers,
        x_chunk=DEFAULT_X_CHUNK,
        y_chunk=DEFAULT_Y_CHUNK,
    )


# ============================================================
# FULL EXTRACTION
# ============================================================

def run_full_extraction(
    query: QuerySpec,
    x_total: int,
    y_total: int,
    max_workers: int,
    auth: AuthenticationManager,
    control: RequestControl,
    run_dir: Path,
    csv_writer,
    csv_handle,
) -> dict:

    profile = load_recommended_profile(
        query,
        max_workers,
    )

    report, _ = run_scheduler(
        query,
        profile,
        x_total,
        y_total,
        auth,
        control,
        run_dir,
        csv_writer,
        csv_handle,
        mode="run",
        capture_values=False,
        benchmark_results=[],
    )

    report["bottleneck"] = identify_bottleneck(
        report
    )

    return {
        "mode": "run",

        "query_hash": query.query_hash,

        "status": report["status"],

        "profile": report,

        "expected_coordinates": (
            x_total
            * y_total
            * len(query.measures)
        ),

        "raw_response_directory": str(
            run_dir
            / profile.name
            / "responses"
        ),
    }


# ============================================================
# INITIAL INFORMATION
# ============================================================

def render_startup(
    query: QuerySpec,
    args,
    run_dir: Path,
):

    console.clear()

    console.print(
        Panel(
            Text(
                APP_NAME
                + f"\nVERSION {APP_VERSION}"
                + "\nXMLA / MDX / CURL_CFFI",
                justify="center",
                style="bold cyan",
            ),
            border_style="cyan",
        )
    )

    configuration = kv_table([
        ("Mode", args.mode.upper()),

        ("Category", "MV — รถยนต์"),

        ("Cube", "FY — ปีบัญชี"),

        ("Catalog", CATALOG),

        ("Cube name", CUBE),

        ("Max workers", args.workers),

        ("HTTP engine", "CURL_CFFI"),

        ("Query hash", query.query_hash),

        ("Output", str(run_dir)),
    ])

    console.print(
        make_panel(
            configuration,
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
        no_wrap=True,
    )

    fields.add_column(
        "OIC Caption"
    )

    fields.add_column(
        "Verified MDX"
    )

    for index, (caption, mdx) in enumerate(
        zip(
            query.x_captions,
            query.x,
        ),
        start=1,
    ):

        fields.add_row(
            f"X{index}",
            caption,
            Text(mdx),
        )

    for index, (caption, mdx) in enumerate(
        zip(
            query.y_captions,
            query.y,
        ),
        start=1,
    ):

        fields.add_row(
            f"Y{index}",
            caption,
            Text(mdx),
        )

    for index, (caption, mdx) in enumerate(
        zip(
            query.measure_captions,
            query.measures,
        ),
        start=1,
    ):

        fields.add_row(
            f"M{index}",
            caption,
            Text(mdx),
        )

    console.print(
        make_panel(
            fields,
            "VERIFIED FIXED QUERY",
        )
    )

    console.print(
        Text(
            "Validating authenticated XMLA access "
            "and discovering exact axis cardinalities...",
            style="cyan",
        )
    )


# ============================================================
# FINAL REPORT
# ============================================================

def render_final_report(
    summary: dict,
    run_dir: Path,
):

    status = summary.get(
        "status",
        "UNKNOWN",
    )

    console.print()

    console.print(
        Panel(
            Text(
                "OIC LOAD ENGINE\n"
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

    if summary.get("mode") == "benchmark":

        table = Table(
            title="BENCHMARK RESULTS",
            box=box.SIMPLE,
            expand=True,
        )

        for column in (
            "Profile",
            "Engine",
            "Workers",
            "Chunk",
            "Requests",
            "Elapsed",
            "P95",
            "Cells/s",
            "Speedup",
            "Match",
        ):

            table.add_column(
                column
            )

        for result in summary.get(
            "profiles",
            [],
        ):

            equivalent = result.get(
                "equivalent"
            )

            table.add_row(
                result["profile"],

                result["engine"],

                str(
                    result["workers_count"]
                ),

                (
                    f"{result['x_chunk']}/"
                    f"{result['y_chunk']}"
                ),

                str(result["requests"]),

                format_duration(
                    result["elapsed"]
                ),

                format_ms(
                    result.get("p95_ms")
                ),

                (
                    f"{result['cells_per_second']:.2f}"
                ),

                (
                    f"{result.get('speedup', 0):.2f}x"
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
                        else "red"
                        if equivalent is False
                        else "dim"
                    ),
                ),
            )

        console.print(table)

        recommended = summary.get(
            "recommended_profile"
        )

        if recommended:

            recommendation_table = kv_table([
                (
                    "Profile",
                    recommended["profile"],
                ),

                (
                    "Engine",
                    recommended["engine"],
                ),

                (
                    "Workers",
                    recommended["workers"],
                ),

                (
                    "X Chunk",
                    recommended["x_chunk"],
                ),

                (
                    "Y Chunk",
                    recommended["y_chunk"],
                ),

                (
                    "Sample speedup",
                    Text(
                        f"{recommended['sample_speedup']:.2f}x",
                        style="bold green",
                    ),
                ),

                (
                    "Sample cells/sec",
                    recommended[
                        "sample_cells_per_second"
                    ],
                ),

                (
                    "Sample P95",
                    format_ms(
                        recommended.get(
                            "sample_p95_ms"
                        )
                    ),
                ),
            ])

            console.print(
                make_panel(
                    recommendation_table,
                    "RECOMMENDED CONFIGURATION",
                )
            )

            bottleneck = recommended.get(
                "bottleneck"
            ) or {}

            if bottleneck:

                console.print(
                    make_panel(
                        Group(
                            Text(
                                bottleneck.get(
                                    "name",
                                    "UNKNOWN",
                                ),
                                style="bold yellow",
                            ),

                            Text(
                                bottleneck.get(
                                    "message",
                                    "",
                                )
                            ),
                        ),
                        "PERFORMANCE DIAGNOSIS",
                    )
                )

    elif summary.get("mode") == "run":

        report = summary.get(
            "profile",
            {},
        )

        if report:

            console.print(
                make_panel(
                    kv_table([
                        (
                            "Profile",
                            report["profile"],
                        ),

                        (
                            "Completed chunks",
                            report["completed_tasks"],
                        ),

                        (
                            "Failed chunks",
                            report["failed_tasks"],
                        ),

                        (
                            "Requests",
                            report["requests"],
                        ),

                        (
                            "Non-null cells",
                            f"{report['non_null_cells']:,}",
                        ),

                        (
                            "Saved XML",
                            format_bytes(
                                report["response_bytes"]
                            ),
                        ),

                        (
                            "Cells/sec",
                            report["cells_per_second"],
                        ),

                        (
                            "Average latency",
                            format_ms(
                                report.get("average_ms")
                            ),
                        ),

                        (
                            "P95 latency",
                            format_ms(
                                report.get("p95_ms")
                            ),
                        ),

                        (
                            "Elapsed",
                            format_duration(
                                report["elapsed"]
                            ),
                        ),
                    ]),
                    "EXTRACTION RESULT",
                )
            )

    console.print(
        make_panel(
            kv_table([
                (
                    "Summary",
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
                    "Raw XML",
                    str(run_dir),
                ),
            ]),
            "OUTPUT FILES",
        )
    )

    if summary.get("error"):

        console.print(
            Text(
                summary["error"],
                style="bold red",
            )
        )


# ============================================================
# MAIN
# ============================================================

def main() -> int:

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
                query.x_captions
            ),

            "y_captions": list(
                query.y_captions
            ),

            "measure_captions": list(
                query.measure_captions
            ),

            "mapping_source": query.mapping_source,
        }

        render_startup(
            query,
            args,
            run_dir,
        )

        with console.status(
            "Connecting to OIC XMLA...",
            spinner="dots",
        ):

            x_total, y_total, discovery = (
                discover_axis_counts(
                    query,
                    auth,
                    control,
                    run_dir,
                )
            )

        summary["axis_discovery"] = discovery

        summary["axis_counts"] = {
            "x_members": x_total,

            "y_members": y_total,

            "measures": len(
                query.measures
            ),

            "potential_coordinates": (
                x_total
                * y_total
                * len(query.measures)
            ),
        }

        console.print(
            Text(
                "AUTHENTICATION VALID | "
                f"X={x_total:,} | "
                f"Y={y_total:,} | "
                f"M={len(query.measures)}",
                style="bold green",
            )
        )

        if x_total == 0 or y_total == 0:

            raise EngineError(
                "Fixed query returned an empty axis"
            )

        with (
            run_dir
            / "requests.csv"
        ).open(
            "w",
            encoding="utf-8-sig",
            newline="",
        ) as csv_handle:

            csv_writer = csv.DictWriter(
                csv_handle,
                fieldnames=CSV_FIELDS,
                extrasaction="ignore",
            )

            csv_writer.writeheader()

            if args.mode == "benchmark":

                result = run_benchmark(
                    query,
                    x_total,
                    y_total,
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
                    x_total,
                    y_total,
                    args.workers,
                    auth,
                    control,
                    run_dir,
                    csv_writer,
                    csv_handle,
                )

            summary.update(result)

    except KeyboardInterrupt:

        STOP_EVENT.set()

        summary["status"] = "STOPPED"

    except StopRequested:

        STOP_EVENT.set()

        summary["status"] = "STOPPED"

    except AuthenticationError as exc:

        summary["status"] = "FAILED"

        summary["error"] = safe_error(exc)

    except EngineError as exc:

        summary["status"] = "FAILED"

        summary["error"] = str(exc)[:500]

    except Exception as exc:

        summary["status"] = "FAILED"

        summary["error"] = safe_error(exc)

    finally:

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

        write_json(
            run_dir / "summary.json",
            summary,
        )

    render_final_report(
        summary,
        run_dir,
    )

    return (
        0
        if summary["status"] == "COMPLETED"
        else 1
    )


if __name__ == "__main__":

    raise SystemExit(main())