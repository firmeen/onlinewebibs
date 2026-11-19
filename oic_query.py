#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
OIC QUERY ENGINE
================

WebForm metadata discovery:
    GET  DynamicPivotCube.aspx?Class=<CLASS>
    POST DynamicPivotCube.aspx/GetStructure
    POST DynamicPivotCube.aspx/GET_MENUTEMPLATE_LIST
    POST DynamicPivotCube.aspx/BindChangeTemplate

XMLA metadata discovery:
    MDSCHEMA_HIERARCHIES
    MDSCHEMA_MEASURES

Authentication:
    - Existing OIC_COOKIE is used for the ASP.NET session.
    - WebForm Authorization values are discovered from the
      authenticated page's JavaScript.
    - Every WebForm method uses its own Authorization value.
    - XMLA uses the supplied Bearer JWT.
    - No Authorization values are hardcoded.
    - No browser automation is required.

Catalog:
    AH / FI / MR / MI / MV

Requirements:
    requests

Safety:
    - No network traffic during module import.
    - Original field captions are preserved.
    - Unverified mappings are never executed.
    - Invalid responses never overwrite existing catalog files.
    - Authentication secrets are never written to catalog JSON.
"""

from __future__ import annotations

import hashlib
import json
import re
import time

from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urljoin, urlparse
from xml.etree import ElementTree as ET


# ============================================================
# CONSTANTS
# ============================================================

ROOT = Path(__file__).resolve().parent

CATALOG_DIR = ROOT / "output" / "catalog"

BASE_URL = "https://onlinewebibs.oic.or.th"

WEBFORM = (
    BASE_URL
    + "/WebForm/Dashboard/Pivot/DynamicPivotCube.aspx"
)

PROXY_URL = (
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

SCHEMA_VERSION = 2

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/157.0.0.0 Safari/537.36"
)

WEBFORM_METHODS = (
    "GetStructure",
    "GET_MENUTEMPLATE_LIST",
    "BindChangeTemplate",
)

FIELD_COLUMNS = (
    "DATATABLE",
    "DATAFIELD",
    "CAPTION",
    "DATATYPE",
    "SUMMARYTYPE",
    "FORMAT",
    "AREA",
    "CALCULATE",
    "GROUP_DIMENSION",
    "DIMENSION",
    "SEQ",
)

CATEGORY = {
    "มุมมอง": "DIMENSION",
    "ข้อมูล": "MEASURE",
}

IDENTIFIER = re.compile(
    r"^\[[^\]\r\n]+\]\.\[[^\]\r\n]+\]$"
)

MEASURE_IDENTIFIER = re.compile(
    r"^\[Measures\]\.\[[^\]\r\n]+\]$"
)

KNOWN_MEASURE_NAMES = {
    "ACT POLICY CNT": "จำนวนกรมธรรม์",
    "ACT PREMIUM AMT": "เบี้ยประกันภัย",
    "EARNED PREMIUM": "เบี้ยประกันภัยที่ถือเป็นรายได้",
    "ACT SUM INS AMT": "จำนวนเงินเอาประกันภัย",
    "CLAIM CNT": "จำนวนเคลม",
    "CLAIM PAID": "สินไหมจ่าย",
    "CLAIM OS": "สินไหมค้างจ่าย",
    "CLAIM LI": "ภาระสินไหม",
}

KNOWN_DIMENSION_NAMES = {
    "YEAR": "ปีบัญชี",
    "FY YEAR": "ปีบัญชี",
    "MAIN NAME TH": "ประเภทงานหลัก",
    "SUB NAME TH": "ประเภทงานย่อย",
    "SUBCLASS DESC TH": "ประเภทกรมธรรม์",
    "GROUP DESC TH": "กลุ่มประเภท",
}


# ============================================================
# INSURANCE PROFILES
# ============================================================

PROFILES = {
    "AH": {
        "name": "อุบัติเหตุและสุขภาพ",
        "cls": "AH",
        "catalog": "PIVOT_202609_AH",
        "cube": "AH_AGGREGATE_FY",
        "year": "[VW PV AH AGGREGATE FY].[FY YEAR]",
        "row1": "[DIM MAINSUB CLASS AH].[MAIN NAME TH]",
        "row2": "[DIM MAINSUB CLASS AH].[SUB NAME TH]",
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "EARNED",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "INMTH CLM PAID",
            "ATMTH CLM LI",
        ],
    },
    "FI": {
        "name": "อัคคีภัยและทรัพย์สิน",
        "cls": "FI",
        "catalog": "PIVOT_202609_FI",
        "cube": "FI_AGGREGATE_FY",
        "year": "[VW PV FI AGGREGATE FY].[YEAR]",
        "row1": "[DIM CT SUBCLASS FI].[GROUP DESC TH]",
        "row2": None,
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "EARNED PREMIUM",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "SUM CLM PAID",
            "SUM CLM OS",
            "SUM CLM LI",
        ],
    },
    "MR": {
        "name": "ทางทะเลและขนส่ง",
        "cls": "MR",
        "catalog": "PIVOT_202609_MR",
        "cube": "MR_AGGREGATE_FY",
        "year": None,
        "row1": "[DIM CT SUBCLASS MR].[SUBCLASS DESC TH]",
        "row2": None,
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "CLAIM PAID",
            "CLAIM OS",
            "CLAIM LI",
        ],
    },
    "MI": {
        "name": "เบ็ดเตล็ด",
        "cls": "MI",
        "catalog": "PIVOT_202609_MI",
        "cube": "MI_AGGREGATE_FY",
        "year": "[VW PV MI AGGREGATE FY].[FY YEAR]",
        "row1": "[DIM MAINSUB CLASS].[MAIN NAME TH]",
        "row2": "[DIM MAINSUB CLASS].[SUB NAME TH]",
        "measures": [
            "ACT POLICY CNT",
            "ACT PREMIUM AMT",
            "ACT SUM INS AMT",
            "CLAIM CNT",
            "CLM PAID",
            "CLM OS",
            "CLM LI",
        ],
    },
    "MV": {
        "name": "รถยนต์",
        "cls": "MO",
        "catalog": "PIVOT_202609_MV",
        "cube": "MV_AGGREGATE_FY",
        "year": "[VW PV MV AGGREGATE FY].[YEAR]",
        "row1": "[DIM CT SUBCLASS MV].[SUBCLASS DESC TH]",
        "row2": None,
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
    },
}


# ============================================================
# UTILITIES
# ============================================================

def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def valid_identifier(value: str) -> bool:
    return (
        isinstance(value, str)
        and bool(
            IDENTIFIER.fullmatch(value)
            or MEASURE_IDENTIFIER.fullmatch(value)
        )
    )


def normalize_cookie(value: str) -> str:
    value = (value or "").strip()

    if value.lower().startswith("cookie:"):
        value = value[7:].strip()

    return value


def normalize_token(value: str) -> str:
    value = (value or "").strip()

    if value.lower().startswith("bearer "):
        value = value[7:].strip()

    return value


def field_id(code: str, record: dict) -> str:
    identity = [code] + [
        record.get(key)
        for key in (
            "GROUP_DIMENSION",
            "DIMENSION",
            "DATATABLE",
            "DATAFIELD",
            "CAPTION",
            "SUMMARYTYPE",
            "CALCULATE",
            "SEQ",
        )
    ]

    encoded = json.dumps(
        identity,
        ensure_ascii=False,
        sort_keys=True,
        default=str,
    ).encode("utf-8")

    return hashlib.sha256(encoded).hexdigest()[:20]


# ============================================================
# CATALOG STORAGE
# ============================================================

def cache_path(code: str) -> Path:
    return CATALOG_DIR / f"{code}.json"


def seeded_catalog(code: str) -> dict:
    if code not in PROFILES:
        raise ValueError("Unknown insurance category")

    profile = PROFILES[code]
    fields = []

    dimensions = filter(
        None,
        (
            profile["year"],
            profile["row1"],
            profile["row2"],
        ),
    )

    for index, name in enumerate(dimensions):
        caption = name.rsplit("[", 1)[-1].rstrip("]")

        fields.append({
            "id": f"profile-d-{index}-{code}",
            "mdx": name,
            "display": caption,
            "caption": caption,
            "type": "DIMENSION",
            "group_dimension": "มุมมอง",
            "dimension": "",
            "source": "PROFILE",
            "status": "PROFILE",
            "mapping_status": "PROFILE",
            "seq": index,
        })

    for index, name in enumerate(profile["measures"]):
        mdx = f"[Measures].[{name}]"

        fields.append({
            "id": f"profile-m-{index}-{code}",
            "mdx": mdx,
            "display": name,
            "caption": name,
            "type": "MEASURE",
            "group_dimension": "ข้อมูล",
            "dimension": "",
            "source": "PROFILE",
            "status": "PROFILE",
            "mapping_status": "PROFILE",
            "seq": index,
        })

    return {
        "code": code,
        "web_class": profile["cls"],
        "catalog": profile["catalog"],
        "cube": profile["cube"],
        "source": "PROFILE",
        "refreshed_at": None,
        "fields": fields,
        "view_count": sum(
            f["type"] == "DIMENSION"
            for f in fields
        ),
        "data_count": sum(
            f["type"] == "MEASURE"
            for f in fields
        ),
        "errors": [],
    }


def load_catalog(code: str) -> dict:
    if code not in PROFILES:
        raise ValueError("Unknown insurance category")

    fallback = seeded_catalog(code)

    try:
        cached = json.loads(
            cache_path(code).read_text(
                encoding="utf-8"
            )
        )

        if not isinstance(cached, dict):
            return fallback

        if cached.get("code") != code:
            return fallback

        if cached.get("cube") != fallback["cube"]:
            return fallback

        fields = cached.get("fields")

        if not isinstance(fields, list):
            return fallback

        fields = [
            field
            for field in fields
            if isinstance(field, dict)
            and field.get("type") in (
                "DIMENSION",
                "MEASURE",
            )
            and field.get("id")
        ]

        if not fields:
            return fallback

        cached["fields"] = fields
        cached["source"] = "CACHE"

        return cached

    except (
        OSError,
        UnicodeError,
        ValueError,
        TypeError,
    ):
        return fallback


def save_catalog(catalog: dict) -> None:
    code = catalog.get("code")

    if code not in PROFILES:
        raise ValueError("Invalid catalog category")

    if catalog.get("cube") != PROFILES[code]["cube"]:
        raise ValueError("Invalid catalog cube")

    fields = catalog.get("fields")

    if not isinstance(fields, list) or not fields:
        raise ValueError(
            "Refusing to overwrite catalog with empty fields"
        )

    CATALOG_DIR.mkdir(
        parents=True,
        exist_ok=True,
    )

    destination = cache_path(code)
    temporary = destination.with_suffix(".tmp")

    try:
        temporary.write_text(
            json.dumps(
                catalog,
                ensure_ascii=False,
                indent=2,
                default=str,
            ),
            encoding="utf-8",
        )

        temporary.replace(destination)

    finally:
        temporary.unlink(missing_ok=True)


# ============================================================
# WEBFORM AUTHORIZATION DISCOVERY
# ============================================================

WEBFORM_AUTH_PATTERN = re.compile(
    r"""
    url\s*:\s*
    (?P<quote>['"])
    (?:[^'"]*/)?
    DynamicPivotCube\.aspx/
    (?P<method>
        GetStructure
        |GET_MENUTEMPLATE_LIST
        |BindChangeTemplate
    )
    (?P=quote)
    (?P<between>.{0,2000}?)
    headers\s*:\s*\{
    \s*
    ['"]Authorization['"]
    \s*:\s*
    (?P<value_quote>['"])
    (?P<authorization>[^'"]{20,200})
    (?P=value_quote)
    """,
    re.IGNORECASE | re.DOTALL | re.VERBOSE,
)


def extract_webform_authorizations(
    source: str,
) -> dict[str, list[str]]:
    """
    Extract the per-method Authorization values generated
    by the authenticated ASP.NET page.

    Never log or persist these values.
    """

    result = {
        method: []
        for method in WEBFORM_METHODS
    }

    method_lookup = {
        method.upper(): method
        for method in WEBFORM_METHODS
    }

    for match in WEBFORM_AUTH_PATTERN.finditer(source):
        method_name = method_lookup.get(
            match.group("method").upper()
        )

        if not method_name:
            continue

        authorization = match.group(
            "authorization"
        ).strip()

        if not re.fullmatch(
            r"[A-Fa-f0-9]{32,128}",
            authorization,
        ):
            continue

        if authorization not in result[method_name]:
            result[method_name].append(
                authorization
            )

    return result


def create_webform_session(
    cookie: str,
):
    import requests

    raw_cookie = normalize_cookie(cookie)

    if not raw_cookie:
        raise ValueError(
            "OIC_COOKIE is required for WebForm requests"
        )

    session = requests.Session()

    session.headers.update({
        "User-Agent": USER_AGENT,
        "Accept-Language": "th-TH,th;q=0.9",
    })

    hostname = urlparse(BASE_URL).hostname

    for part in raw_cookie.split(";"):
        name, separator, value = (
            part.strip().partition("=")
        )

        if not separator or not name.strip():
            continue

        session.cookies.set(
            name.strip(),
            value.strip(),
            domain=hostname,
            path="/",
            secure=True,
        )

    return session


def discover_webform_authorizations(
    session,
    code: str,
    *,
    timeout: int = 30,
) -> dict[str, list[str]]:
    """
    Load the real authenticated WebForm page.

    The page supplies a distinct Authorization value
    for each ASP.NET WebMethod.

    No values from another browser session are assumed.
    """

    if code not in PROFILES:
        raise ValueError(
            f"Unknown insurance category: {code}"
        )

    profile = PROFILES[code]

    referer = (
        f"{WEBFORM}?Class={profile['cls']}"
    )

    response = session.get(
        referer,
        headers={
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,*/*;q=0.8"
            ),
            "Referer": BASE_URL + "/",
        },
        timeout=(10, timeout),
        allow_redirects=False,
    )

    if response.status_code in (
        301, 302, 303, 307, 308, 401, 403
    ):
        raise ValueError(
            f"{code}: WebForm page HTTP "
            f"{response.status_code}; "
            "the browser session may require renewal"
        )

    if response.status_code != 200:
        raise ValueError(
            f"{code}: WebForm page HTTP "
            f"{response.status_code}"
        )

    page = response.text

    if not page.strip():
        raise ValueError(
            f"{code}: WebForm returned an empty page"
        )

    auth = extract_webform_authorizations(page)

    required = set(WEBFORM_METHODS)

    if required.issubset(
        method
        for method, values in auth.items()
        if values
    ):
        return auth

    # Only inspect same-origin JavaScript referenced by the page.
    # Inline page scripts are preferred; external scripts are
    # fetched only if required Authorization values are missing.

    script_pattern = re.compile(
        r"""<script\b[^>]*\bsrc\s*=\s*["']([^"']+)["']""",
        re.IGNORECASE,
    )

    urls = []
    seen = set()

    for match in script_pattern.finditer(page):
        url = urljoin(
            referer,
            match.group(1),
        )

        parsed = urlparse(url)

        if parsed.scheme != "https":
            continue

        if parsed.hostname != urlparse(BASE_URL).hostname:
            continue

        if url in seen:
            continue

        seen.add(url)
        urls.append(url)

    for script_url in urls[:20]:
        if all(
            auth[method]
            for method in WEBFORM_METHODS
        ):
            break

        try:
            script_response = session.get(
                script_url,
                headers={
                    "Accept": (
                        "text/javascript,application/javascript,"
                        "*/*;q=0.8"
                    ),
                    "Referer": referer,
                },
                timeout=(10, timeout),
                allow_redirects=False,
            )

            if script_response.status_code != 200:
                continue

            discovered = extract_webform_authorizations(
                script_response.text
            )

            for method, values in discovered.items():
                for value in values:
                    if value not in auth[method]:
                        auth[method].append(value)

        except requests.RequestException:
            continue

    if not auth["GetStructure"]:
        raise ValueError(
            f"{code}: GetStructure Authorization "
            "not found in the authenticated WebForm. "
            "Inspect the browser's Network request headers."
        )

    return auth


# ============================================================
# WEBFORM RESPONSE PARSER
# ============================================================

def parse_webform_response(
    response,
    *,
    method: str = "WebForm",
) -> list:
    """
    ASP.NET WebMethods normally return:

        {"d": "[{...}, {...}]"}

    The d property may also already be decoded.

    Response diagnostics deliberately exclude raw server
    content, cookies, Authorization and session tokens.
    """

    status = response.status_code

    content_type = (
        response.headers.get(
            "Content-Type",
            "",
        ).split(";", 1)[0].strip().lower()
    )

    if status in (301, 302, 303, 307, 308):
        raise ValueError(
            f"{method}: HTTP {status}; "
            "redirected instead of returning metadata"
        )

    if status in (401, 403):
        raise ValueError(
            f"{method}: HTTP {status}; "
            "WebForm authorization or session rejected"
        )

    if status != 200:
        raise ValueError(
            f"{method}: HTTP {status}"
        )

    try:
        wrapper = response.json()

    except ValueError as exc:
        raw_preview = response.text[:250].lstrip().lower()

        if (
            raw_preview.startswith("<!doctype")
            or raw_preview.startswith("<html")
            or raw_preview.startswith("<")
        ):
            description = (
                "HTML returned instead of JSON; "
                "possible login page or security challenge"
            )
        else:
            description = (
                "response is not valid JSON"
            )

        raise ValueError(
            f"{method}: {description}; "
            f"HTTP={status}; "
            f"Content-Type={content_type or 'unknown'}; "
            f"Bytes={len(response.content)}"
        ) from exc

    if not isinstance(wrapper, dict):
        raise ValueError(
            f"{method}: expected a JSON object; "
            f"received {type(wrapper).__name__}"
        )

    if "d" not in wrapper:
        raise ValueError(
            f"{method}: JSON response has no 'd' property; "
            f"keys={list(wrapper)[:10]}"
        )

    payload = wrapper["d"]

    for _ in range(3):
        if not isinstance(payload, str):
            break

        try:
            payload = json.loads(payload)

        except (ValueError, TypeError) as exc:
            raise ValueError(
                f"{method}: 'd' is not a JSON array; "
                "possibly an ASP.NET error or status string"
            ) from exc

    if isinstance(payload, list):
        return payload

    raise ValueError(
        f"{method}: expected an array inside 'd'; "
        f"received {type(payload).__name__}"
    )


# ============================================================
# AUTHENTICATED WEBFORM METHOD REQUEST
# ============================================================

def post_webform_method(
    session,
    method: str,
    payload: dict,
    authorizations: dict[str, list[str]],
    referer: str,
    *,
    timeout: int = 30,
) -> list:
    """
    Send the original ASP.NET WebMethod request.

    Authorization is method-specific and obtained from
    the current authenticated WebForm page.
    """

    if method not in WEBFORM_METHODS:
        raise ValueError(
            f"Unsupported WebForm method: {method}"
        )

    values = authorizations.get(method) or []

    if not values:
        raise ValueError(
            f"{method}: method-specific Authorization "
            "was not found"
        )

    url = f"{WEBFORM}/{method}"

    last_error = None

    for authorization in values:
        headers = {
            "Authorization": authorization,
            "Content-Type": (
                "application/json; charset=UTF-8"
            ),
            "Accept": (
                "application/json, text/javascript, "
                "*/*; q=0.01"
            ),
            "X-Requested-With": "XMLHttpRequest",
            "Origin": BASE_URL,
            "Referer": referer,
        }

        response = session.post(
            url,
            json=payload,
            headers=headers,
            timeout=(10, timeout),
            allow_redirects=False,
        )

        try:
            return parse_webform_response(
                response,
                method=method,
            )

        except ValueError as exc:
            last_error = exc

            # Another Authorization value from the same
            # page may be accepted by this WebMethod.
            # Do not retry on unrelated response errors.

            if response.status_code not in (
                200, 401, 403
            ):
                raise

    raise ValueError(
        f"{method}: all discovered Authorization "
        f"values were rejected; last error: {last_error}"
    )


# ============================================================
# TEMPLATE METADATA NORMALIZATION
# ============================================================

def normalize_template_records(
    code: str,
    menu_records: list,
    binding_records: list,
) -> list[dict]:
    """
    GET_MENUTEMPLATE_LIST:
        TEMPLATE_DETAIL can be null.

    BindChangeTemplate:
        Returns records with NAME and JSON, where JSON
        contains the actual pivot field mappings.

    Combine both formats without inventing MDX mappings.
    """

    profile = PROFILES[code]

    result = []

    for item in menu_records:
        if not isinstance(item, dict):
            continue

        if item.get("CLASS") not in (
            code,
            profile["cls"],
        ):
            continue

        if item.get("CUBE_NAME") != profile["cube"]:
            continue

        detail = item.get("TEMPLATE_DETAIL")

        if not detail:
            continue

        try:
            decoded = (
                json.loads(detail)
                if isinstance(detail, str)
                else detail
            )
        except (ValueError, TypeError):
            continue

        if not isinstance(decoded, list):
            continue

        result.append({
            "CLASS": profile["cls"],
            "CUBE_NAME": profile["cube"],
            "TEMPLATE_NAME": (
                item.get("TEMPLATE_NAME") or ""
            ),
            "TEMPLATE_DETAIL": decoded,
        })

    for item in binding_records:
        if not isinstance(item, dict):
            continue

        detail = (
            item.get("JSON")
            or item.get("TEMPLATE_DETAIL")
        )

        if not detail:
            continue

        try:
            decoded = (
                json.loads(detail)
                if isinstance(detail, str)
                else detail
            )
        except (ValueError, TypeError):
            continue

        if not isinstance(decoded, list):
            continue

        result.append({
            "CLASS": profile["cls"],
            "CUBE_NAME": profile["cube"],
            "TEMPLATE_NAME": (
                item.get("NAME") or ""
            ),
            "TEMPLATE_DETAIL": decoded,
        })

    return result


# ============================================================
# GETSTRUCTURE NORMALIZATION
# ============================================================

def normalize_getstructure(
    code: str,
    raw: object,
    *,
    templates: list | None = None,
    xmla: list | None = None,
) -> dict:
    """
    CAPTION is preserved exactly as returned by OIC.

    Mapping rules:
        - Exact template caption + folder + area.
        - Exact template caption + area if unambiguous.
        - Exact XMLA caption only if unambiguous.

    Unverified and calculated fields remain visible but
    cannot be executed as MDX.
    """

    if code not in PROFILES:
        raise ValueError(
            "Unknown cube category"
        )

    if not isinstance(raw, list) or not raw:
        raise ValueError(
            "GetStructure must return a nonempty array"
        )

    profile = PROFILES[code]

    template_index: dict[
        tuple,
        set[tuple[str, bool]],
    ] = {}

    for item in templates or []:
        if not isinstance(item, dict):
            continue

        if item.get("CLASS") not in (
            code,
            profile["cls"],
        ):
            continue

        if item.get("CUBE_NAME") != profile["cube"]:
            continue

        detail = item.get("TEMPLATE_DETAIL")

        try:
            detail = (
                json.loads(detail)
                if isinstance(detail, str)
                else detail
            )
        except (ValueError, TypeError):
            continue

        if not isinstance(detail, list):
            continue

        for part in detail:
            if not isinstance(part, dict):
                continue

            mdx = part.get("dataField", "")

            if not valid_identifier(mdx):
                continue

            label = part.get("caption", "")

            if not isinstance(label, str) or not label:
                continue

            folder = part.get("displayFolder") or ""

            area = str(
                part.get("area") or ""
            ).lower()

            if area not in (
                "row",
                "column",
                "filter",
                "data",
            ):
                continue

            transformed = bool(
                part.get("summaryDisplayMode")
            )

            entry = (mdx, transformed)

            template_index.setdefault(
                (label, folder, area),
                set(),
            ).add(entry)

            template_index.setdefault(
                (label, "", area),
                set(),
            ).add(entry)

    xmla_index: dict[
        tuple[str, str],
        set[str],
    ] = {}

    for item in xmla or []:
        if not isinstance(item, dict):
            continue

        kind = item.get("type")
        mdx = item.get("mdx")
        display = item.get("display") or ""

        if kind not in (
            "DIMENSION",
            "MEASURE",
        ):
            continue

        if not valid_identifier(mdx):
            continue

        xmla_index.setdefault(
            (kind, display),
            set(),
        ).add(mdx)

    fields = []

    for index, row in enumerate(raw):
        if not isinstance(row, dict):
            raise ValueError(
                f"Invalid GetStructure row {index}"
            )

        if "GROUP_DIMENSION" not in row:
            raise ValueError(
                f"Missing GROUP_DIMENSION in row {index}"
            )

        group = row.get("GROUP_DIMENSION") or ""
        kind = CATEGORY.get(group)

        if not kind:
            continue

        caption = row.get("CAPTION")

        if not isinstance(caption, str) or not caption:
            caption = (
                row.get("DATAFIELD")
                or row.get("DATATABLE")
                or "(unnamed)"
            )

        folder = row.get("DIMENSION") or ""

        areas = (
            ["data"]
            if kind == "MEASURE"
            else ["row", "column", "filter"]
        )

        candidates = set()

        for area in areas:
            candidates.update(
                template_index.get(
                    (caption, folder, area),
                    set(),
                )
            )

        if not candidates:
            for area in areas:
                candidates.update(
                    template_index.get(
                        (caption, "", area),
                        set(),
                    )
                )

        if not candidates:
            candidates = {
                (mdx, False)
                for mdx in xmla_index.get(
                    (kind, caption),
                    set(),
                )
            }

        transformations_only = (
            bool(candidates)
            and all(
                transformed
                for _, transformed in candidates
            )
        )

        mapped_names = {
            mdx
            for mdx, transformed in candidates
            if not transformed
            and (
                bool(MEASURE_IDENTIFIER.fullmatch(mdx))
                if kind == "MEASURE"
                else bool(IDENTIFIER.fullmatch(mdx))
            )
        }

        calculation = row.get("CALCULATE")

        summary = str(
            row.get("SUMMARYTYPE") or ""
        ).strip().lower()

        derived = (
            bool(calculation)
            or summary not in ("", "sum")
            or transformations_only
        )

        if len(mapped_names) == 1 and not derived:
            mdx = next(iter(mapped_names))
            status = "MAPPED"

        elif derived:
            mdx = None
            status = "UNSUPPORTED"

        else:
            mdx = None
            status = "UNVERIFIED"

        retained = {
            key: row.get(key)
            for key in FIELD_COLUMNS
        }

        fields.append({
            "id": field_id(code, row),
            "mdx": mdx,
            "display": caption,
            "caption": caption,
            "type": kind,
            "group_dimension": group,
            "dimension": folder,
            "source": "GETSTRUCTURE",
            "status": status,
            "mapping_status": status,
            "seq": row.get("SEQ"),
            "details": retained,
        })

    if not fields:
        raise ValueError(
            "No usable GetStructure fields"
        )

    return {
        "code": code,
        "web_class": profile["cls"],
        "cube": profile["cube"],
        "catalog": profile["catalog"],
        "source": "LIVE",
        "refreshed_at": datetime.now(
            timezone.utc
        ).isoformat(),
        "fields": fields,
        "view_count": sum(
            field["type"] == "DIMENSION"
            for field in fields
        ),
        "data_count": sum(
            field["type"] == "MEASURE"
            for field in fields
        ),
        "errors": [],
    }


# ============================================================
# XMLA DISCOVER REQUEST BUILDER
# ============================================================

def make_discover_request(
    code: str,
    request_type: str,
) -> bytes:
    supported = {
        "MDSCHEMA_DIMENSIONS",
        "MDSCHEMA_HIERARCHIES",
        "MDSCHEMA_LEVELS",
        "MDSCHEMA_MEASURES",
        "MDSCHEMA_MEMBERS",
    }

    if request_type not in supported:
        raise ValueError(
            "Unsupported XMLA Discover request"
        )

    if code not in PROFILES:
        raise ValueError(
            "Unknown insurance category"
        )

    profile = PROFILES[code]

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
    ).text = profile["catalog"]

    ET.SubElement(
        restrictions,
        f"{{{XMLA_NS}}}CUBE_NAME",
    ).text = profile["cube"]

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
    ).text = profile["catalog"]

    ET.SubElement(
        properties,
        f"{{{XMLA_NS}}}LocaleIdentifier",
    ).text = "9"

    return ET.tostring(
        envelope,
        encoding="utf-8",
    )


# ============================================================
# XMLA METADATA DISCOVERY
# ============================================================

def discover_xmla_metadata(
    session,
    code: str,
    token: str,
    referer: str,
    *,
    timeout: int = 30,
) -> tuple[list[dict], list[str]]:
    import requests

    fields = []
    errors = []

    token = normalize_token(token)

    if not token:
        return fields, errors

    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "text/xml",
        "Accept": "text/plain, */*; q=0.01",
        "Origin": BASE_URL,
        "Referer": referer,
        "X-Requested-With": "XMLHttpRequest",
    }

    for kind in (
        "MDSCHEMA_HIERARCHIES",
        "MDSCHEMA_MEASURES",
    ):
        try:
            response = session.post(
                PROXY_URL,
                data=make_discover_request(
                    code,
                    kind,
                ),
                headers=headers,
                timeout=(10, timeout),
                allow_redirects=False,
            )

            if response.status_code != 200:
                raise ValueError(
                    f"HTTP {response.status_code}"
                )

            root = ET.fromstring(response.content)

            if any(
                local_name(element.tag) == "Fault"
                for element in root.iter()
            ):
                raise ValueError(
                    "XMLA SOAP Fault"
                )

            for row in root.iter():
                if local_name(row.tag) != "row":
                    continue

                attrs = {
                    local_name(element.tag): (
                        element.text or ""
                    )
                    for element in row
                }

                if kind == "MDSCHEMA_HIERARCHIES":
                    mdx = attrs.get(
                        "HIERARCHY_UNIQUE_NAME"
                    )

                    display = attrs.get(
                        "HIERARCHY_CAPTION"
                    )

                    field_type = "DIMENSION"

                else:
                    mdx = attrs.get(
                        "MEASURE_UNIQUE_NAME"
                    )

                    display = attrs.get(
                        "MEASURE_CAPTION"
                    )

                    field_type = "MEASURE"

                if valid_identifier(mdx):
                    fields.append({
                        "mdx": mdx,
                        "display": display or "",
                        "type": field_type,
                    })

        except (
            requests.RequestException,
            ValueError,
            ET.ParseError,
        ) as exc:
            errors.append(
                f"{kind}: {exc}"
            )

    return fields, errors


# ============================================================
# LIVE FIVE-CUBE FIELD DISCOVERY
# ============================================================

def discover_fields(
    code: str,
    token: str,
    cookie: str,
    *,
    timeout: int = 30,
) -> dict:
    """
    Discover the complete field catalog for one OIC cube.

    Workflow:
        1. Create authenticated WebForm session.
        2. GET the correct Class page.
        3. Extract real per-method Authorization values.
        4. POST GetStructure.
        5. POST GET_MENUTEMPLATE_LIST.
        6. POST BindChangeTemplate.
        7. Discover XMLA hierarchy/measure metadata.
        8. Build exact-caption catalog.
        9. Save validated cache.

    Failed mandatory discovery never overwrites cache.
    """

    if code not in PROFILES:
        raise ValueError(
            f"Unknown insurance category: {code}"
        )

    token = normalize_token(token)
    cookie = normalize_cookie(cookie)

    if not cookie:
        raise ValueError(
            f"{code}: OIC_COOKIE is missing"
        )

    profile = PROFILES[code]

    referer = (
        f"{WEBFORM}?Class={profile['cls']}"
    )

    errors = []

    with create_webform_session(cookie) as session:
        # Authentication values must come from the
        # current WebForm page, not an unrelated JWT.

        authorizations = (
            discover_webform_authorizations(
                session,
                code,
                timeout=timeout,
            )
        )

        # ----------------------------------------------------
        # MANDATORY: GETSTRUCTURE
        # ----------------------------------------------------

        structure = post_webform_method(
            session=session,
            method="GetStructure",
            payload={
                "Template": [
                    code,
                    profile["cube"],
                ]
            },
            authorizations=authorizations,
            referer=referer,
            timeout=timeout,
        )

        if not structure:
            raise ValueError(
                f"{code}: GetStructure returned zero fields"
            )

        # ----------------------------------------------------
        # OPTIONAL: TEMPLATE MENU
        # ----------------------------------------------------

        menu_records = []

        try:
            menu_records = post_webform_method(
                session=session,
                method="GET_MENUTEMPLATE_LIST",
                payload={
                    "Template": [
                        profile["cls"],
                    ]
                },
                authorizations=authorizations,
                referer=referer,
                timeout=timeout,
            )

        except Exception as exc:
            errors.append(
                f"GET_MENUTEMPLATE_LIST: {exc}"
            )

        # ----------------------------------------------------
        # OPTIONAL: EXISTING PIVOT MAPPINGS
        # ----------------------------------------------------

        binding_records = []

        try:
            binding_records = post_webform_method(
                session=session,
                method="BindChangeTemplate",
                payload={
                    "Template": [
                        code,
                        profile["cube"],
                    ]
                },
                authorizations=authorizations,
                referer=referer,
                timeout=timeout,
            )

        except Exception as exc:
            errors.append(
                f"BindChangeTemplate: {exc}"
            )

        templates = normalize_template_records(
            code,
            menu_records,
            binding_records,
        )

        # ----------------------------------------------------
        # OPTIONAL: XMLA METADATA
        # ----------------------------------------------------

        xmla_fields, xmla_errors = (
            discover_xmla_metadata(
                session,
                code,
                token,
                referer,
                timeout=timeout,
            )
        )

        errors.extend(xmla_errors)

    # --------------------------------------------------------
    # NORMALIZE CATALOG
    # --------------------------------------------------------

    catalog = normalize_getstructure(
        code,
        structure,
        templates=templates,
        xmla=xmla_fields,
    )

    catalog["errors"] = errors
    catalog["source"] = "LIVE"

    # Never save secrets or partial invalid catalogs.

    save_catalog(catalog)

    return catalog


# ============================================================
# QUERY VALIDATION
# ============================================================

def validate_query(
    config: dict,
    catalog: dict | None = None,
) -> dict:
    code = config.get("code")

    if code not in PROFILES:
        raise ValueError(
            "Select AH/FI/MR/MI/MV"
        )

    profile = PROFILES[code]

    if (
        config.get("catalog") != profile["catalog"]
        or config.get("cube") != profile["cube"]
    ):
        raise ValueError(
            "Catalog/Cube do not match the selected category"
        )

    version = int(
        config.get("schema_version", 1)
    )

    if version not in (1, SCHEMA_VERSION):
        raise ValueError(
            "Unsupported query schema version"
        )

    x = config.get("x")
    y = config.get("y")
    measures = config.get("measures")

    if not all(
        isinstance(value, list) and value
        for value in (x, y, measures)
    ):
        raise ValueError(
            "Select X/ROWS, Y/COLUMNS and at least one Measure"
        )

    if len(x) + len(y) > 5:
        raise ValueError(
            "Maximum five selected hierarchies"
        )

    if (
        len(set(x + y)) != len(x + y)
        or len(set(measures)) != len(measures)
    ):
        raise ValueError(
            "Duplicate hierarchy or Measure"
        )

    for name in x + y:
        if not isinstance(name, str) or not IDENTIFIER.fullmatch(name):
            raise ValueError(
                f"Invalid hierarchy identifier: {name}"
            )

    for name in measures:
        if not isinstance(name, str) or not MEASURE_IDENTIFIER.fullmatch(name):
            raise ValueError(
                f"Invalid Measure identifier: {name}"
            )

    if version == SCHEMA_VERSION:
        identities = config.get(
            "selection_ids",
            {},
        )

        snapshot = config.get(
            "selected_fields",
            [],
        )

        if not isinstance(identities, dict):
            raise ValueError(
                "Missing applied selection identities"
            )

        if not isinstance(snapshot, list):
            raise ValueError(
                "Missing selected fields"
            )

        source = {
            field.get("id"): field
            for field in snapshot
            if isinstance(field, dict)
        }

        if len(source) != len(snapshot):
            raise ValueError(
                "Duplicate selected display identity"
            )

        for axis, names in (
            ("x", x),
            ("y", y),
            ("measures", measures),
        ):
            selected_ids = identities.get(
                axis,
                [],
            )

            if len(selected_ids) != len(names):
                raise ValueError(
                    f"Missing {axis} selection identity"
                )

            for identity, mdx in zip(
                selected_ids,
                names,
            ):
                field = source.get(identity)

                if not field:
                    raise ValueError(
                        f"Unknown field identity: {identity}"
                    )

                if field.get("mdx") != mdx:
                    raise ValueError(
                        f"Field mapping mismatch: {identity}"
                    )

                if field.get("status") not in (
                    "MAPPED",
                    "VERIFIED",
                ):
                    raise ValueError(
                        f"Unverified field in {axis}: {identity}"
                    )

                expected_type = (
                    "MEASURE"
                    if axis == "measures"
                    else "DIMENSION"
                )

                if field.get("type") != expected_type:
                    raise ValueError(
                        "Field type mismatch"
                    )

        if len(
            set(
                identities["x"]
                + identities["y"]
            )
        ) != len(x + y):
            raise ValueError(
                "Duplicate field on both axes"
            )

    else:
        catalog = catalog or load_catalog(code)

        allowed = {
            (
                field.get("mdx"),
                field.get("type"),
            )
            for field in (
                catalog["fields"]
                + seeded_catalog(code)["fields"]
            )
        }

        for name in x + y:
            if (name, "DIMENSION") not in allowed:
                raise ValueError(
                    f"Unknown hierarchy: {name}"
                )

        for name in measures:
            if (name, "MEASURE") not in allowed:
                raise ValueError(
                    f"Unknown Measure: {name}"
                )

    try:
        workers = int(
            config.get("workers", 4)
        )

        x_chunk = int(
            config.get("x_chunk", 1)
        )

        y_chunk = int(
            config.get("y_chunk", 20)
        )

    except (TypeError, ValueError) as exc:
        raise ValueError(
            "Workers and chunk sizes must be integers"
        ) from exc

    if not 1 <= workers <= 6:
        raise ValueError(
            "Workers must be between 1 and 6"
        )

    execution_mode = config.get("execution_mode", "CHUNKED")
    if execution_mode not in ("DIRECT", "CHUNKED"):
        raise ValueError("Unsupported query execution mode")
    if execution_mode == "DIRECT" and (
        version != SCHEMA_VERSION or workers != 1
    ):
        raise ValueError("Direct XMLA requires schema v2 and exactly one worker")
    if execution_mode == "DIRECT":
        response_limit = int(config.get("max_response_mb", 100))
        if not 1 <= response_limit <= 1024:
            raise ValueError("Direct XMLA response limit must be between 1 and 1024 MB")

    if not 1 <= x_chunk <= 20:
        raise ValueError(
            "X chunk size must be between 1 and 20"
        )

    if not 1 <= y_chunk <= 100:
        raise ValueError(
            "Y chunk size must be between 1 and 100"
        )

    if config.get("filters"):
        raise ValueError(
            "Filters require verified metadata"
        )

    for secret in (
        "cookie",
        "token",
        "jwt",
        "authorization",
        "OIC_COOKIE",
        "OIC_BEARER_TOKEN",
    ):
        if secret in config:
            raise ValueError(
                "Authentication secrets must not be "
                "stored in query configuration"
            )

    return config


# ============================================================
# DEFAULT QUERY
# ============================================================

def default_query(
    code: str,
    *,
    workers: int = 1,
    schema_version: int = 1,
) -> dict:
    profile = PROFILES[code]

    result = {
        "code": code,
        "cube": profile["cube"],
        "catalog": profile["catalog"],
        "x": (
            [profile["year"]]
            if profile["year"]
            else []
        ),
        "y": [profile["row1"]],
        "measures": [
            f"[Measures].[{name}]"
            for name in profile["measures"]
        ],
        "filters": [],
        "workers": workers,
        "x_chunk": 1,
        "y_chunk": 20,
    }

    if schema_version != 1:
        result["schema_version"] = schema_version

    return result


# ============================================================
# MDX AXIS BUILDER
# ============================================================

def axis_set(
    fields: list[str],
    offset: int,
    size: int,
) -> str:
    if not fields or offset < 0 or size < 1:
        raise ValueError(
            "Invalid axis pagination"
        )

    members = [
        field + ".AllMembers"
        for field in fields
    ]

    result = members[0]

    for other in members[1:]:
        result = (
            f"CrossJoin({result},{other})"
        )

    return (
        f"Subset({result},{offset},{size})"
    )


# ============================================================
# MDX BUILDER
# ============================================================

def build_mdx(
    config: dict,
    x_offset: int = 0,
    y_offset: int = 0,
) -> str:
    validate_query(config)

    x = axis_set(
        config["x"],
        x_offset,
        int(config.get("x_chunk", 1)),
    )

    y = axis_set(
        config["y"],
        y_offset,
        int(config.get("y_chunk", 20)),
    )

    measures = (
        "{"
        + ",".join(config["measures"])
        + "}"
    )

    if int(
        config.get("schema_version", 1)
    ) == SCHEMA_VERSION:
        # Modern selected query:
        # X = ROWS / Axis1
        # Y = COLUMNS / Axis0

        columns = y
        rows = f"CrossJoin({x},{measures})"

    else:
        # Legacy query format.

        columns = x
        rows = f"CrossJoin({y},{measures})"

    cube = PROFILES[
        config["code"]
    ]["cube"]

    return (
        f"SELECT {columns} "
        "DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,"
        "HIERARCHY_UNIQUE_NAME,"
        "MEMBER_VALUE ON COLUMNS,"
        f"{rows} "
        "DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,"
        "HIERARCHY_UNIQUE_NAME,"
        "MEMBER_VALUE ON ROWS "
        f"FROM [{cube}] "
        "CELL PROPERTIES "
        "VALUE,FORMAT_STRING,LANGUAGE"
    )



# ============================================================
# DIRECT XMLA QUERY (NO OFFSET OR CHUNK PAGINATION)
# ============================================================

def build_direct_mdx(config: dict) -> str:
    """Compile a complete selected query in the main collector's NonEmpty style.

    Keep measures on Axis1 for schema v2 / excel_records compatibility.
    No Subset, speculative member counts, network discovery, or page probes.
    """
    validate_query(config)
    if int(config.get("schema_version", 1)) != SCHEMA_VERSION:
        raise ValueError("Direct queries require the applied v2 axis schema")

    def all_members(fields: list[str]) -> str:
        expression = fields[0] + ".AllMembers"
        for field in fields[1:]:
            expression = f"CrossJoin({expression},{field}.AllMembers)"
        return expression

    measures = "{" + ",".join(config["measures"]) + "}"
    columns = f"NonEmpty({all_members(config['y'])},{measures})"
    rows = f"NonEmpty({all_members(config['x'])},{measures})"
    return (
        f"WITH SET [DX_columns] AS {columns} "
        f"SET [DX_rows] AS {rows} "
        "SELECT [DX_columns] "
        "DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME,MEMBER_VALUE ON COLUMNS, "
        f"CrossJoin([DX_rows],{measures}) "
        "DIMENSION PROPERTIES "
        "PARENT_UNIQUE_NAME,HIERARCHY_UNIQUE_NAME,MEMBER_VALUE ON ROWS "
        f"FROM [{config['cube']}] "
        "CELL PROPERTIES VALUE,FORMAT_STRING,LANGUAGE"
    )


# ============================================================
# XMLA SOAP BUILDER
# ============================================================

def build_soap(
    config: dict,
    x_offset: int = 0,
    y_offset: int = 0,
    *,
    statement: str | None = None,
) -> bytes:
    mdx = statement if statement is not None else build_mdx(
        config,
        x_offset,
        y_offset,
    )

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

    values = {
        "Catalog": config["catalog"],
        "ShowHiddenCubes": "True",
        "Timeout": "3600",
        "LocaleIdentifier": "9",
    }

    for key, value in values.items():
        ET.SubElement(
            properties,
            f"{{{XMLA_NS}}}{key}",
        ).text = value

    return ET.tostring(
        envelope,
        encoding="utf-8",
    )


# ============================================================
# QUERY HASH
# ============================================================

def query_hash(config: dict) -> str:
    canonical = {
        key: value
        for key, value in config.items()
        if key not in (
            "run_id",
            "created_at",
            "output",
        )
    }

    encoded = json.dumps(
        canonical,
        ensure_ascii=False,
        sort_keys=True,
    ).encode("utf-8")

    return hashlib.sha256(
        encoded
    ).hexdigest()[:16]


# ============================================================
# XMLA DATASET PARSER
# ============================================================

def parse_xmla(content: bytes) -> dict:
    root = ET.fromstring(content)

    for element in root.iter():
        if local_name(element.tag) == "Fault":
            message = " ".join(
                (node.text or "")
                for node in element.iter()
                if node.text
            )

            raise ValueError(
                f"SOAP Fault: {message[:500]}"
            )

    dataset = root.find(
        f".//{{{MD_NS}}}root"
    )

    if dataset is None:
        raise ValueError(
            "XMLA multidimensional dataset missing"
        )

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
                        )
                        or ""
                    ),
                    "caption": (
                        member.findtext(
                            f"{{{MD_NS}}}Caption"
                        )
                        or ""
                    ),
                    "level": (
                        member.findtext(
                            f"{{{MD_NS}}}LNum"
                        )
                        or ""
                    ),
                })

            tuples.append(members)

        axes[axis.get("name", "")] = tuples

    if (
        "Axis0" not in axes
        or "Axis1" not in axes
    ):
        raise ValueError(
            "Expected Axis0 and Axis1"
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

        ordinal = int(
            cell.get("CellOrdinal", "-1")
        )

        if ordinal >= 0:
            cells[ordinal] = value.text

    last_update = next(
        (
            element.text
            for element in root.iter()
            if local_name(element.tag)
            == "LastDataUpdate"
        ),
        None,
    )

    return {
        "axes": axes,
        "cells": cells,
        "last_data_update": last_update,
    }


# ============================================================
# EXCEL RECORD ITERATOR
# ============================================================

def excel_records(
    config: dict,
    data: dict,
):
    """
    Long-format records without aggregation.

    Every row preserves the selected dimensional members
    and the original value from the XMLA response.
    """

    columns = data["axes"]["Axis0"]
    rows = data["axes"]["Axis1"]

    width = len(columns)

    modern = (
        int(config.get("schema_version", 1))
        == SCHEMA_VERSION
    )

    if width == 0:
        return

    for row_index, row in enumerate(rows):
        measure = next(
            (
                member
                for member in row
                if member["hierarchy"]
                == "[Measures]"
            ),
            None,
        )

        if measure is None:
            raise ValueError(
                "Missing Measures member on Axis1"
            )

        dimensions = [
            member
            for member in row
            if member["hierarchy"]
            != "[Measures]"
        ]

        for column_index, column in enumerate(
            columns
        ):
            ordinal = (
                column_index
                + width * row_index
            )

            if ordinal not in data["cells"]:
                continue

            yield {
                "X": (
                    dimensions
                    if modern
                    else column
                ),
                "Y": (
                    column
                    if modern
                    else dimensions
                ),
                "MEASURE": measure[
                    "unique_name"
                ],
                "VALUE": data["cells"][
                    ordinal
                ],
                "LAST_DATA_UPDATE": data[
                    "last_data_update"
                ],
            }