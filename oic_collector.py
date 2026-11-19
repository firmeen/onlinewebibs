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
                       validate_query, build_mdx as selected_mdx, query_hash)

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



# Applied queries are local-only configuration.  Workers never write this file.
def read_run_status(batch_dir: Path) -> str:
    try:
        return json.loads((batch_dir / "batch_manifest.json").read_text(encoding="utf-8")).get("status", "")
    except (OSError, ValueError, TypeError):
        return ""


def applied_list_path(output: Path) -> Path:
    return output / "applied_queries.json"


def applied_empty() -> dict:
    return {"storage_version": 1, "revision": uuid.uuid4().hex,
            "next_query_sequence": 1, "queries": []}


def applied_validate(document: dict) -> dict:
    if not isinstance(document, dict) or document.get("storage_version") != 1:
        raise ValueError("Invalid applied list storage version")
    if not isinstance(document.get("revision"), str) or not document["revision"]:
        raise ValueError("Invalid applied list revision")
    if not isinstance(document.get("queries"), list):
        raise ValueError("Invalid applied queries array")
    sequence = document.get("next_query_sequence")
    if type(sequence) is not int or sequence < 1:
        raise ValueError("Invalid next query sequence")
    seen = set()
    for entry in document["queries"]:
        if not isinstance(entry, dict) or not re.fullmatch(r"Q\d{3,}", str(entry.get("query_id", ""))):
            raise ValueError("Invalid query identity")
        if entry["query_id"] in seen:
            raise ValueError("Duplicate query identity")
        seen.add(entry["query_id"])
        if not isinstance(entry.get("applied_at"), str):
            raise ValueError("Missing applied time")
        validate_query(entry["config"])
        if entry.get("code") != entry["config"]["code"] or entry.get("cube") != entry["config"]["cube"]:
            raise ValueError("Applied query category/cube mismatch")
        if entry.get("fingerprint") != applied_fingerprint(entry["config"]):
            raise ValueError("Applied query fingerprint mismatch")
        verify_apply_plan(entry["config"], entry["query_id"], entry.get("task_plan"))
    if any(int(q[1:]) >= sequence for q in seen):
        raise ValueError("Next query sequence would reuse an identity")
    return document


def applied_load(path: Path) -> dict:
    try:
        return applied_validate(json.loads(path.read_text(encoding="utf-8")))
    except FileNotFoundError:
        return applied_empty()
    except (OSError, UnicodeError, ValueError, TypeError, KeyError) as exc:
        raise ValueError(f"Applied list cannot be loaded: {type(exc).__name__}: {exc}") from exc


def applied_fingerprint(config: dict) -> str:
    values = {name: config.get(name, [] if name in ("x", "y", "measures", "filters") else None)
              for name in ("code", "cube", "catalog", "x", "y", "measures", "filters")}
    return json.dumps(values, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def make_apply_plan(config: dict, query_id: str) -> dict:
    """Precompute worker ownership of unbounded X pages without doing I/O."""
    validate_query(config)
    count = int(config["workers"])
    size = int(config["x_chunk"])
    plan = {
        "version": 1, "query_id": query_id, "query_hash": query_hash(config),
        "created_at": datetime.now(timezone.utc).isoformat(),
        "status": "PLANNED", "partition": "x_page_stride",
        "workers": count, "x_chunk": size, "y_chunk": int(config["y_chunk"]),
        "total_chunks": None,
        "assignments": [
            {"worker_id": wid, "first_x_offset": (wid - 1) * size,
             "x_stride": count * size, "first_y_offset": 0,
             "y_stride": int(config["y_chunk"])}
            for wid in range(1, count + 1)
        ],
    }
    return plan


def verify_apply_plan(config: dict, query_id: str, plan: dict) -> None:
    if not isinstance(plan, dict) or plan.get("version") != 1:
        raise ValueError(f"{query_id}: missing or invalid task plan")
    expected = make_apply_plan(config, query_id)
    for key in ("query_id", "query_hash", "status", "partition", "workers",
                "x_chunk", "y_chunk", "total_chunks", "assignments"):
        if plan.get(key) != expected[key]:
            raise ValueError(f"{query_id}: task plan mismatch: {key}")


class applied_write_lock:
    """Cross-process lock held across read, revision, and atomic replacement."""
    def __init__(self, path: Path):
        self.path = path.with_name(path.name + ".write.lock")
        self.stream = None

    def __enter__(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.stream = self.path.open("a+b")
        self.stream.seek(0)
        if not self.stream.read(1):
            self.stream.write(b"\0")
            self.stream.flush()
        deadline = time.monotonic() + 8
        while True:
            try:
                self.stream.seek(0)
                if os.name == "nt":
                    import msvcrt
                    msvcrt.locking(self.stream.fileno(), msvcrt.LK_NBLCK, 1)
                else:
                    import fcntl
                    fcntl.flock(self.stream.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
                return self
            except OSError:
                if time.monotonic() >= deadline:
                    self.stream.close()
                    raise TimeoutError("Applied list is locked by another process")
                time.sleep(0.05)

    def __exit__(self, kind, error, traceback):
        try:
            self.stream.seek(0)
            if os.name == "nt":
                import msvcrt
                msvcrt.locking(self.stream.fileno(), msvcrt.LK_UNLCK, 1)
            else:
                import fcntl
                fcntl.flock(self.stream.fileno(), fcntl.LOCK_UN)
        finally:
            self.stream.close()


def local_json_atomic(path: Path, value: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + "." + uuid.uuid4().hex + ".tmp")
    try:
        with temp.open("w", encoding="utf-8") as stream:
            json.dump(value, stream, ensure_ascii=False, indent=2)
            stream.flush()
            os.fsync(stream.fileno())
        os.replace(temp, path)
    finally:
        temp.unlink(missing_ok=True)


def applied_update(path: Path, operation: str, *, config=None,
                   query_id=None, direction=None, allow_duplicate=False):
    with applied_write_lock(path):
        data = applied_load(path)
        entries = data["queries"]
        if operation == "append":
            config = validate_query(dict(config))
            fingerprint = applied_fingerprint(config)
            duplicate = next((entry["query_id"] for entry in entries
                              if applied_fingerprint(entry["config"]) == fingerprint), None)
            if duplicate and not allow_duplicate:
                return duplicate, False
            new_id = f"Q{data['next_query_sequence']:03d}"
            entries.append({"query_id": new_id,
                            "applied_at": datetime.now(timezone.utc).isoformat(),
                            "code": config["code"], "cube": config["cube"],
                            "catalog": config["catalog"],
                            "fingerprint": fingerprint,
                            "task_plan": make_apply_plan(config, new_id),
                            "config": config})
            data["next_query_sequence"] += 1
            result = new_id
        else:
            position = next((index for index, entry in enumerate(entries)
                             if entry["query_id"] == query_id), None)
            if position is None:
                raise ValueError(f"Query {query_id} no longer exists")
            if operation == "remove":
                entries.pop(position)
            elif operation == "move":
                neighbor = position + direction
                if not 0 <= neighbor < len(entries):
                    return query_id, False
                entries[position], entries[neighbor] = entries[neighbor], entries[position]
            else:
                raise ValueError("Invalid applied list operation")
            result = query_id
        data["revision"] = uuid.uuid4().hex
        applied_validate(data)
        local_json_atomic(path, data)
        saved = applied_load(path)
        if saved["revision"] != data["revision"]:
            raise RuntimeError("Applied list write verification failed")
        return result, True

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
        from textual.containers import Horizontal, Vertical, VerticalScroll
        from textual.screen import ModalScreen
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
        )

    except ImportError as exc:
        raise RuntimeError(
            "Install dependencies: "
            "python -m pip install -U requests textual rich"
        ) from exc

    class ForceStopConfirmation(ModalScreen[bool]):
        CSS = """
        ForceStopConfirmation { align: center middle; }
        #force_dialog { width: 74; height: 11; padding: 1 2;
            background: #16232e; border: solid #bd5555; }
        #force_message { height: 1fr; }
        #force_buttons { height: 3; align: center middle; }
        #force_buttons Button { margin-right: 2; }
        """

        def compose(self) -> ComposeResult:
            with Vertical(id="force_dialog"):
                yield Static("FORCE STOP ACTIVE BATCH", classes="designer_section")
                yield Static("Terminate only verified worker/coordinator processes "
                             "from this batch. Active requests will be interrupted; "
                             "verified XML is preserved.", id="force_message")
                with Horizontal(id="force_buttons"):
                    yield Button("FORCE STOP", id="force_accept", variant="error")
                    yield Button("CANCEL", id="force_cancel")

        def on_button_pressed(self, event: Button.Pressed) -> None:
            self.dismiss(event.button.id == "force_accept")

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
            height: 3;
            margin-top: 1;
            align: left middle;
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
            height: 16; min-height: 9; margin-top: 1;
            padding: 0 1; border: solid #304352; background: #16232e;
        }
        #applied_heading { height: 1; color: #c9b996; text-style: bold; }
        #applied_queries { height: 1fr; }
        #applied_actions { height: 3; }
        #applied_actions Button { margin-right: 1; }

        #catalog_status { height: 3; margin-top: 1; padding: 0 2;
                           border: solid #293c49; color: #9fb0ba; }
        #activity_panel { height: 8; min-height: 6; }
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
            self.field_catalog = load_catalog("MV")
            self.selected_field = None
            self.field_rows = {}
            self.axis_x = []
            self.axis_y = []
            self.selected_measures = []
            self.active_run = None
            self.last_run_path = None
            self.applied_config = None  # legacy call compatibility only
            self.selection_revision = None
            self.applied_queries = []
            self.applied_signature = object()
            self.applied_error_signature = None
            self.selected_query_id = None
            self.catalog_running = False
            self.catalog_started = False
            self.catalog_statuses = {code: "CACHE" if load_catalog(code)["source"] == "CACHE" else "PROFILE" for code in QUERY_PROFILES}

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

            with VerticalScroll(id="workspace"):

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
                    yield Button(
                        "START COLLECTION",
                        id="start",
                    )

                    yield Button(
                        "STOP",
                        id="stop",
                    )

                    yield Button(
                        "OPEN OUTPUT",
                        id="open",
                    )
                    yield Button("RESUME", id="resume_run")
                    yield Button("FORCE STOP", id="force_stop", variant="error", disabled=True)

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
                    yield Static("APPLIED QUERIES / 0 QUERIES", id="applied_heading")
                    yield DataTable(id="applied_queries", cursor_type="row", zebra_stripes=True)
                    with Horizontal(id="applied_actions"):
                        yield Button("REMOVE", id="applied_remove")
                        yield Button("MOVE UP", id="applied_up")
                        yield Button("MOVE DOWN", id="applied_down")
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

            self.query_one(
                "#stop",
                Button,
            ).disabled = True
            self.query_one("#force_stop", Button).disabled = True
            applied_table = self.query_one("#applied_queries", DataTable)
            for heading in ("ID", "CATEGORY / CUBE", "X", "Y", "VALUES", "WORKERS", "CHUNKS", "PLAN"):
                applied_table.add_column(heading)
            self.select_category("MV")
            self.set_interval(1.0, self.poll_selected_run)
            self.set_interval(1.0, self.poll_applied_selection)
            self.poll_applied_selection()
            self.update_catalog_summary()
            # Recover a previously launched batch after reopening Main App.
            try:
                from oic_parallel_collector import batch_controller_active
                runs = resolve_output_path(self.output_value()) / "Runs"
                if runs.is_dir():
                    for path in sorted(runs.iterdir(), reverse=True):
                        if (path / "batch_manifest.json").is_file():
                            self.last_run_path = self.last_run_path or path
                            if batch_controller_active(path):
                                self.active_run = path
                                self.query_one("#start", Button).disabled = True
                                self.query_one("#stop", Button).disabled = (path / "stop.flag").exists()
                                self.query_one("#force_stop", Button).disabled = not (path / "stop.flag").exists()
                                self.update_query_summary()
                                break
            except Exception as exc:
                self.info(f"Batch recovery status unavailable: {exc}", "warning")

            # Load local configuration without making
            # any network requests.

            self.reload_from_env(
                auto_check=False
            )

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
            if self.busy:
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

            if self.busy:
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
            if not self.catalog_started:
                self.discover_all_catalogs(force=False)

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
                if self.active_run is not None:
                    self.push_screen(ForceStopConfirmation(), self.force_stop_confirmed)

            elif button_id == "open":
                self.action_open_output()

            elif button_id == "resume_run":
                self.resume_selected_query()
            elif button_id == "open_designer":
                self.open_query_designer()
            elif button_id == "refresh_all_catalogs":
                self.discover_all_catalogs(force=True)
            elif button_id in ("applied_remove", "applied_up", "applied_down"):
                self.change_applied_query(button_id)

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
            self.start_selected_query()

        def action_stop_collection(self) -> None:
            if self.active_run is not None:
                (self.active_run / "stop.flag").touch(exist_ok=True)
                self.info("STOPPING: no new chunks/queries will start; current requests may finish.", "warning")
                self.query_one("#status", Static).update("STOPPING / Waiting for active workers")
                self.query_one("#stop", Button).disabled = True
                self.query_one("#force_stop", Button).disabled = False
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

        def force_stop_confirmed(self, approved: bool) -> None:
            if not approved or self.active_run is None:
                return
            try:
                from oic_parallel_collector import force_stop_batch
                count = force_stop_batch(self.active_run)
                self.info(f"FORCE STOP requested: terminated {count} verified batch processes", "warning")
                self.query_one("#status", Static).update("STOPPING / Force termination requested")
            except Exception as exc:
                self.info(f"FORCE STOP failed: {type(exc).__name__}: {exc}", "error")

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
            if self.busy:
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
            return applied_list_path(resolve_output_path(self.output_value()))

        def select_category(self, code: str) -> None:
            if code in QUERY_PROFILES:
                self.selected_code = code

        def on_select_changed(self, event: Select.Changed) -> None:
            if event.select.id == "cube_select" and event.value in QUERY_PROFILES:
                self.select_category(str(event.value))

        def on_data_table_row_highlighted(self, event: DataTable.RowHighlighted) -> None:
            if event.data_table.id == "applied_queries":
                self.selected_query_id = str(event.row_key.value)

        def update_query_summary(self) -> None:
            if not self.is_mounted:
                return
            table = self.query_one("#applied_queries", DataTable)
            previous = self.selected_query_id
            table.clear(columns=False)
            for entry in self.applied_queries:
                config = entry["config"]
                table.add_row(entry["query_id"],
                              f"{config['code']} / {config['cube']}",
                              str(len(config["x"])), str(len(config["y"])),
                              str(len(config["measures"])), str(config["workers"]),
                              f"{config['x_chunk']}/{config['y_chunk']}",
                              entry["task_plan"]["status"], key=entry["query_id"])
            identities = [entry["query_id"] for entry in self.applied_queries]
            self.selected_query_id = previous if previous in identities else (identities[0] if identities else None)
            if self.selected_query_id:
                table.move_cursor(row=identities.index(self.selected_query_id))
            self.query_one("#applied_heading", Static).update(
                f"APPLIED QUERIES — {len(identities)} QUERIES" if identities
                else "NO APPLIED QUERIES / Open Query Designer to add a query")
            for control in ("applied_remove", "applied_up", "applied_down"):
                self.query_one("#" + control, Button).disabled = (
                    not bool(identities) or self.active_run is not None)

        def query_config(self) -> dict:
            if not self.applied_queries:
                raise ValueError("Apply at least one query before Start")
            selected = next((item for item in self.applied_queries
                             if item["query_id"] == self.selected_query_id), self.applied_queries[0])
            return validate_query(dict(selected["config"]))

        def poll_applied_selection(self) -> None:
            path = self.selection_file()
            try:
                stat = path.stat()
                signature = (str(path), stat.st_mtime_ns, stat.st_size)
            except FileNotFoundError:
                signature = None
            except OSError as exc:
                self.info(f"Applied list stat failed: {exc}", "error")
                return
            if signature == self.applied_signature:
                return
            try:
                document = applied_load(path)
            except ValueError as exc:
                if signature != self.applied_error_signature:
                    self.info(str(exc), "error")
                self.applied_error_signature = signature
                return
            self.applied_error_signature = None
            self.applied_signature = signature
            if document["revision"] == self.selection_revision:
                return
            self.applied_queries = document["queries"]
            self.selection_revision = document["revision"]
            self.applied_config = None
            self.update_query_summary()

        def change_applied_query(self, control: str) -> None:
            if self.active_run is not None:
                self.info("Applied queue editing is disabled during an active batch", "warning")
                return
            if not self.selected_query_id:
                self.info("Select a query first", "warning")
                return
            try:
                operation = "remove" if control == "applied_remove" else "move"
                direction = -1 if control == "applied_up" else 1
                query_id, changed = applied_update(self.selection_file(), operation,
                                                   query_id=self.selected_query_id, direction=direction)
                if changed:
                    self.poll_applied_selection()
                    self.info(f"Applied list updated: {query_id}", "success")
            except Exception as exc:
                self.info(f"Applied list update failed: {type(exc).__name__}: {exc}", "error")

        def open_query_designer(self) -> None:
            if sys.platform != "win32":
                self.info(
                    "Dedicated Query Designer requires Windows.",
                    "error",
                )
                return

            from oic_parallel_collector import launch_terminal_window

            selection = self.selection_file().resolve()
            lock_path = selection.with_name(selection.name + ".designer.lock")

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

                launch_terminal_window(
                    [designer_command],
                    title="OIC | Query Designer",
                    cwd=BASE_DIR,
                    tab_titles=["OIC | Query Designer"],
                    keep_open=False,
                    maximized=True,
                )

                self.info(
                    "Query Designer opened in maximized PowerShell.",
                    "success",
                )

            except (OSError, RuntimeError, ValueError) as exc:
                self.info(
                    f"Query Designer launch failed: {exc}",
                    "error",
                )

        def update_catalog_summary(self) -> None:
            status = "  |  ".join(f"{code}: {self.catalog_statuses.get(code,'—')}"
                                  for code in QUERY_PROFILES)
            self.query_one("#catalog_status", Static).update("FY FIELD CATALOGS\n" + status)

        @work(thread=True, group="five-cube-discovery", exclusive=True)
        def discover_all_catalogs(self, force: bool = False) -> None:
            if self.catalog_running or self.cookie_state != "VALID":
                return
            self.catalog_running = True
            self.catalog_started = True
            sequence = self.auth_sequence
            cookie, token = self.cookie_value(), self.token_value()
            for code in QUERY_PROFILES:
                if sequence != self.auth_sequence:
                    break
                existing = load_catalog(code)
                if not force and existing.get("source") == "CACHE" and existing.get("refreshed_at"):
                    try:
                        updated = datetime.fromisoformat(existing["refreshed_at"])
                        if (datetime.now(timezone.utc) - updated).total_seconds() < 86400:
                            self.call_from_thread(self.catalog_discovery_event, code,
                                                  existing, None)
                            continue
                    except (TypeError, ValueError):
                        pass
                try:
                    catalog = discover_fields(code, token, cookie)
                    self.call_from_thread(self.catalog_discovery_event, code, catalog, None)
                except Exception as exc:
                    error = str(exc)[:110]
                    self.call_from_thread(self.catalog_discovery_event, code, None, error)
                    if "SESSION_EXPIRED" in error or "401" in error or "403" in error:
                        break
            self.catalog_running = False

        def catalog_discovery_event(self, code: str, catalog: dict | None,
                                    error: str | None) -> None:
            if error:
                self.catalog_statuses[code] = "ERROR / CACHE KEPT"
                self.info(f"{code} metadata: {error}", "warning")
            else:
                self.catalog_statuses[code] = (
                    f"{catalog['view_count']}+{catalog['data_count']}")
                self.info(f"{code}: {len(catalog['fields'])} field items cached.", "success")
            self.update_catalog_summary()

        # ====================================================
        # SELECTED DATA EXTRACTION / WINDOW HANDOFF
        # ====================================================

        def start_selected_query(self) -> None:
            if self.active_run is not None:
                self.info("Batch collection is already active", "warning")
                return
            if self.cookie_state != "VALID" or jwt_info(self.token_value())["status"] != "VALID":
                self.info("Validate Cookie and JWT before Start", "error")
                return
            if sys.platform != "win32":
                self.info("PowerShell collection requires Windows", "error")
                return
            try:
                # Read the canonical list once: the resulting snapshot stays immutable.
                document = applied_load(self.selection_file())
                if not document["queries"]:
                    raise ValueError("No applied queries")
                queries = []
                for entry in document["queries"]:
                    config = validate_query(dict(entry["config"]))
                    verify_apply_plan(config, entry["query_id"], entry["task_plan"])
                    queries.append((entry["query_id"], config, entry["task_plan"]))
                from oic_parallel_collector import launch_terminal_window, find_powershell
                find_powershell()
                script = BASE_DIR / "oic_parallel_collector.py"
                if not script.is_file():
                    raise FileNotFoundError(str(script))
                root = resolve_output_path(self.output_value())
                root.mkdir(parents=True, exist_ok=True)
                self.save_credentials()
                run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
                batch_dir = root / "Runs" / run_id
                batch_dir.mkdir(parents=True, exist_ok=False)
                items = []
                for query_id, config, plan in queries:
                    config["run_id"] = run_id  # excluded from oic_query.query_hash
                    name = f"{query_id}/query_config.json"
                    local_json_atomic(batch_dir / name, config)
                    local_json_atomic(batch_dir / query_id / "task_plan.json", plan)
                    items.append(dict(query_id=query_id, code=config["code"], cube=config["cube"],
                                      config_path=name, plan_path=f"{query_id}/task_plan.json",
                                      status="PENDING", started_at=None,
                                      finished_at=None, excel_path=None, records=0, error=None))
                manifest = dict(storage_version=1, run_id=run_id,
                                created_at=datetime.now(timezone.utc).isoformat(),
                                started_at=None, finished_at=None, status="STARTING",
                                current_query_id=None, queries=items)
                local_json_atomic(batch_dir / "batch_manifest.json", manifest)
                from oic_parallel_collector import batch_controller_active
                for other in (root / "Runs").glob("*/batch_manifest.json"):
                    if other.parent == batch_dir:
                        continue
                    if batch_controller_active(other.parent):
                        raise RuntimeError("Another batch is active: " + other.parent.name)
                monitor = launch_terminal_window(
                    [[str(Path(sys.executable).resolve()), str(script.resolve()),
                      "--batch-dir", str(batch_dir), "--batch-monitor"]],
                    title="OIC | Live Monitor", cwd=BASE_DIR, keep_open=True)
                try:
                    coordinator_log = (batch_dir / "coordinator.log").open("ab")
                    try:
                        subprocess.Popen(
                            [str(Path(sys.executable).resolve()), str(script.resolve()),
                             "--batch-dir", str(batch_dir), "--launch"],
                            cwd=str(BASE_DIR), stdout=coordinator_log,
                            stderr=subprocess.STDOUT,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                    finally:
                        coordinator_log.close()
                except Exception:
                    if hasattr(monitor, "terminate") and monitor.poll() is None:
                        monitor.terminate()
                    raise
                self.active_run = batch_dir
                self.last_run_path = batch_dir
                self.query_one("#start", Button).disabled = True
                self.query_one("#stop", Button).disabled = False
                self.query_one("#force_stop", Button).disabled = True
                self.update_query_summary()
                self.query_one("#status", Static).update(f"BATCH STARTING / {run_id}")
                self.info(f"Batch launched: {run_id} / {len(items)} queries", "success")
            except Exception as exc:
                self.info(f"Cannot start batch: {type(exc).__name__}: {exc}", "error")

        def resume_selected_query(self) -> None:
            if self.active_run is not None:
                self.info("A batch is already active", "warning")
                return
            if self.cookie_state != "VALID" or jwt_info(self.token_value())["status"] != "VALID":
                self.info("Validate Cookie and JWT before Resume", "warning")
                return
            root = resolve_output_path(self.output_value()) / "Runs"
            path = self.last_run_path
            if path is None and root.is_dir():
                candidates = sorted((x for x in root.iterdir()
                                     if (x / "batch_manifest.json").is_file()), reverse=True)
                from oic_parallel_collector import batch_controller_active
                path = next((x for x in candidates if
                             read_run_status(x) != "COMPLETED" and
                             not batch_controller_active(x)), None)
            if not path or not (path / "batch_manifest.json").is_file():
                self.info("No previous batch to resume", "warning")
                return
            try:
                manifest = json.loads((path / "batch_manifest.json").read_text(encoding="utf-8"))
                if manifest.get("status") == "COMPLETED":
                    self.info("Batch is already completed", "warning")
                    return
                from oic_parallel_collector import (
                    launch_terminal_window, find_powershell, batch_controller_active,
                )
                find_powershell()
                if batch_controller_active(path):
                    self.info("Batch is already running in another monitor", "warning")
                    return
                (path / "stop.flag").unlink(missing_ok=True)
                script = BASE_DIR / "oic_parallel_collector.py"
                monitor = launch_terminal_window(
                    [[str(Path(sys.executable).resolve()), str(script),
                      "--batch-dir", str(path.resolve()), "--batch-monitor"]],
                    title="OIC | Live Monitor", cwd=BASE_DIR, keep_open=True)
                try:
                    with (path / "coordinator.log").open("ab") as log_stream:
                        subprocess.Popen(
                            [str(Path(sys.executable).resolve()), str(script),
                             "--batch-dir", str(path.resolve()), "--launch"],
                            cwd=str(BASE_DIR), stdout=log_stream, stderr=subprocess.STDOUT,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
                except Exception:
                    if hasattr(monitor, "terminate") and monitor.poll() is None:
                        monitor.terminate()
                    raise
                self.active_run = path
                self.last_run_path = path
                self.query_one("#start", Button).disabled = True
                self.query_one("#stop", Button).disabled = False
                self.query_one("#force_stop", Button).disabled = True
                self.update_query_summary()
                self.info(f"Resuming batch {path.name}", "success")
            except Exception as exc:
                self.info(f"Resume failed: {type(exc).__name__}: {exc}", "error")

        def poll_selected_run(self) -> None:
            if self.active_run is None:
                return
            try:
                data = json.loads((self.active_run / "batch_manifest.json").read_text(encoding="utf-8"))
            except (OSError, ValueError):
                return
            items = data.get("queries", [])
            current = next((q for q in items if q.get("query_id") == data.get("current_query_id")), None)
            finished = sum(q.get("status") in ("COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED") for q in items)
            current_text = (f"{current['query_id']} / {current['code']}" if current else "—")
            self.query_one("#status", Static).update(
                f"BATCH {data.get('status', '?')} / {self.active_run.name} | "
                f"{finished}/{len(items)} queries | CURRENT {current_text}")
            if data.get("status") in ("COMPLETED", "PARTIAL", "FAILED", "STOPPED", "SESSION_EXPIRED"):
                self.info(f"Batch ended: {self.active_run.name} / {data['status']}")
                self.last_run_path = self.active_run
                self.active_run = None
                self.query_one("#start", Button).disabled = False
                self.query_one("#stop", Button).disabled = True
                self.query_one("#force_stop", Button).disabled = True
                self.update_query_summary()

    # ========================================================
    # DEDICATED QUERY DESIGNER — same Textual framework, separate process.
    # ========================================================

    from textual.screen import ModalScreen

    class DuplicateConfirmation(ModalScreen[bool]):
        CSS = """
        DuplicateConfirmation { align: center middle; }
        #duplicate_dialog { width: 68; height: 11; padding: 1 2;
            background: #16232e; border: solid #c9b996; }
        #duplicate_message { height: 1fr; }
        #duplicate_buttons { height: 3; align: center middle; }
        #duplicate_buttons Button { margin-right: 2; }
        """

        def __init__(self, query_id: str, config: dict):
            super().__init__()
            self.query_id = query_id
            self.config = config

        def compose(self) -> ComposeResult:
            with Vertical(id="duplicate_dialog"):
                yield Static("DUPLICATE QUERY DETECTED", classes="designer_section")
                yield Static(f"Same configuration as {self.query_id} "
                             f"({self.config['code']} / {self.config['cube']}).",
                             id="duplicate_message")
                with Horizontal(id="duplicate_buttons"):
                    yield Button("ADD ANYWAY", id="duplicate_add", variant="warning")
                    yield Button("CANCEL", id="duplicate_cancel")

        def on_button_pressed(self, event: Button.Pressed) -> None:
            self.dismiss(event.button.id == "duplicate_add")

    class OICQueryDesigner(App):
        TITLE = "OIC | Query Designer"
        SUB_TITLE = "Five-Cube FY Field Catalog"
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
            height: 2;
            min-height: 2;
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
            self.catalog = load_catalog(self.code)
            self.fields = {}
            self.chosen = {"x": [], "y": [], "measures": []}
            self.by_category = {}
            self.highlighted_id = None

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
                        "CANCEL",
                        id="designer_cancel",
                    )

            yield Footer()

        def on_mount(self) -> None:
            self.selection_path.parent.mkdir(parents=True, exist_ok=True)
            self.selection_path.with_name(self.selection_path.name + ".designer.lock").write_text(
                str(os.getpid()), encoding="ascii")
            self._restore_previous_selection()
            self.update_catalog(self.code)

        def on_unmount(self) -> None:
            lock = self.selection_path.with_name(self.selection_path.name + ".designer.lock")
            try:
                if lock.read_text(encoding="ascii") == str(os.getpid()):
                    lock.unlink(missing_ok=True)
            except OSError:
                pass

        def _restore_previous_selection(self) -> None:
            try:
                entries = applied_load(self.selection_path)["queries"]
                for entry in reversed(entries):
                    previous = entry["config"]
                    if previous.get("code") != self.code:
                        continue
                    ids = previous.get("selection_ids", {})
                    self.by_category[self.code] = {
                        key: list(ids.get(key, [])) for key in self.chosen
                    }
                    for name, key in (("workers", "designer_workers"),
                                      ("x_chunk", "designer_x_chunk"),
                                      ("y_chunk", "designer_y_chunk")):
                        self.query_one("#" + key, Input).value = str(previous[name])
                    break
            except (ValueError, TypeError, KeyError) as exc:
                self.show_error(f"Applied list cannot be restored: {exc}")

        def update_catalog(self, code: str) -> None:
            self.code = code
            self.catalog = load_catalog(code)
            self.fields = {field["id"]: field for field in self.catalog["fields"]}
            previous = self.by_category.get(code, {})
            self.chosen = {axis: [item for item in previous.get(axis, [])
                                  if item in self.fields and
                                  self.fields[item].get("status") in ("MAPPED", "VERIFIED")]
                           for axis in ("x", "y", "measures")}
            self.highlighted_id = None
            self.refresh_tree()
            self.refresh_layout()
            self.query_one("#designer_status", Static).update(
                f"CATALOG {code}: {self.catalog.get('view_count', 0)} views / "
                f"{self.catalog.get('data_count', 0)} data  ·  {self.catalog['source']}"
            )

        def on_select_changed(self, event: Select.Changed) -> None:
            if event.select.id != "designer_category":
                return
            code = event.value
            if code in QUERY_PROFILES and code != self.code:
                self.by_category[self.code] = {
                    key: list(value) for key, value in self.chosen.items()}
                self.update_catalog(str(code))

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
            visible = 0

            for field in self.catalog["fields"]:
                caption = field.get(
                    "caption",
                    field.get("display", ""),
                )

                terms = " ".join(
                    str(field.get(key) or "")
                    for key in (
                        "caption",
                        "group_dimension",
                        "dimension",
                        "mdx",
                    )
                )

                terms += " " + " ".join(
                    str(field.get("details", {}).get(key) or "")
                    for key in (
                        "DATAFIELD",
                        "DATATABLE",
                    )
                )

                if search and search not in terms.casefold():
                    continue

                group = field.get("group_dimension") or (
                    "ข้อมูล"
                    if field["type"] == "MEASURE"
                    else "มุมมอง"
                )

                if group not in groups:
                    groups[group] = tree.root.add(
                        group,
                        expand=True,
                    )

                parent = groups[group]
                folder = field.get("dimension") or ""

                if folder:
                    folder_key = (group, folder)

                    if folder_key not in folders:
                        folders[folder_key] = parent.add(
                            folder,
                            expand=True,
                        )

                    parent = folders[folder_key]

                status = field.get("status")

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

                visible += 1

            tree.root.expand()

            self.query_one(
                "#field_description",
                Static,
            ).update(
                f"{visible} visible / "
                f"{len(self.catalog['fields'])} total"
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
                    "A → X/ROWS   S → Y/COLUMNS"
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
                    f"mapping not available ({field.get('status')})"
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
                    f"{field.get('status')} — selection disabled"
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

            widget.set_options(
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
            profile = QUERY_PROFILES[self.code]

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
                    f"CODE: {query['code']} / {query['cube']}\n"
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

        def apply_selection(self) -> None:
            try:
                config = self.build_query()
                identity, changed = applied_update(self.selection_path, "append", config=config)
                if changed:
                    self.query_one("#designer_preview", Static).update(
                        f"APPLIED {identity} / {config['code']} / PLAN READY "
                        f"({config['workers']} workers). Add another query without closing.")
                    self.by_category[self.code] = {
                        key: list(values) for key, values in self.chosen.items()}
                else:
                    # Confirmation must not mutate the list or the Designer's selected fields.
                    self.push_screen(DuplicateConfirmation(identity, config),
                                     lambda approved: self.confirm_duplicate(config) if approved else None)
            except Exception as exc:
                self.show_error(f"APPLY FAILED: {type(exc).__name__}: {exc}")

        def confirm_duplicate(self, config: dict) -> None:
            try:
                # The append operation takes a fresh lock and reads latest storage.
                query_id, _ = applied_update(self.selection_path, "append", config=config,
                                              allow_duplicate=True)
                self.query_one("#designer_preview", Static).update(
                    f"APPLIED {query_id} (duplicate confirmed) / PLAN READY. "
                    "Designer remains open.")
            except Exception as exc:
                self.show_error(f"APPLY FAILED: {type(exc).__name__}: {exc}")

        def catalog_refreshed(self, code: str, catalog: dict | None, error: str | None):
            if error:
                self.show_error("Catalog refresh failed; existing cache kept: " + error)
            elif code == self.code:
                previous = {k: list(v) for k, v in self.chosen.items()}
                self.by_category[code] = previous
                self.update_catalog(code)

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

            elif operation == "designer_preview_button":
                self.preview_query()

            elif operation == "designer_apply":
                self.apply_selection()

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