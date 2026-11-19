#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared OIC cube catalog, validated query definitions, MDX and XMLA parsing.

No network traffic occurs on import or while building previews.
"""
from __future__ import annotations

import hashlib
import time
import json
import re
from datetime import datetime, timezone
from pathlib import Path
from xml.etree import ElementTree as ET

SOAP_NS = "http://schemas.xmlsoap.org/soap/envelope/"
XMLA_NS = "urn:schemas-microsoft-com:xml-analysis"
MD_NS = "urn:schemas-microsoft-com:xml-analysis:mddataset"
ET.register_namespace("soap", SOAP_NS)
ROOT = Path(__file__).resolve().parent
CATALOG_DIR = ROOT / "output" / "catalog"
PROXY_URL = "https://onlinewebibs.oic.or.th/xmla-proxy/api/xmla-proxy"

# Profiles already used by the existing OIC collector; MR has no confirmed YEAR.
PROFILES = {
    "AH": dict(name="อุบัติเหตุและสุขภาพ", cls="AH", catalog="PIVOT_202609_AH",
               cube="AH_AGGREGATE_FY", year="[VW PV AH AGGREGATE FY].[FY YEAR]",
               row1="[DIM MAINSUB CLASS AH].[MAIN NAME TH]",
               row2="[DIM MAINSUB CLASS AH].[SUB NAME TH]",
               measures=["ACT POLICY CNT", "ACT PREMIUM AMT", "EARNED", "ACT SUM INS AMT",
                         "CLAIM CNT", "INMTH CLM PAID", "ATMTH CLM LI"]),
    "FI": dict(name="อัคคีภัยและทรัพย์สิน", cls="FI", catalog="PIVOT_202609_FI",
               cube="FI_AGGREGATE_FY", year="[VW PV FI AGGREGATE FY].[YEAR]",
               row1="[DIM CT SUBCLASS FI].[GROUP DESC TH]", row2=None,
               measures=["ACT POLICY CNT", "ACT PREMIUM AMT", "EARNED PREMIUM",
                         "ACT SUM INS AMT", "CLAIM CNT", "SUM CLM PAID",
                         "SUM CLM OS", "SUM CLM LI"]),
    "MR": dict(name="ทางทะเลและขนส่ง", cls="MR", catalog="PIVOT_202609_MR",
               cube="MR_AGGREGATE_FY", year=None,
               row1="[DIM CT SUBCLASS MR].[SUBCLASS DESC TH]", row2=None,
               measures=["ACT POLICY CNT", "ACT PREMIUM AMT", "ACT SUM INS AMT",
                         "CLAIM CNT", "CLAIM PAID", "CLAIM OS", "CLAIM LI"]),
    "MI": dict(name="เบ็ดเตล็ด", cls="MI", catalog="PIVOT_202609_MI",
               cube="MI_AGGREGATE_FY", year="[VW PV MI AGGREGATE FY].[FY YEAR]",
               row1="[DIM MAINSUB CLASS].[MAIN NAME TH]",
               row2="[DIM MAINSUB CLASS].[SUB NAME TH]",
               measures=["ACT POLICY CNT", "ACT PREMIUM AMT", "ACT SUM INS AMT",
                         "CLAIM CNT", "CLM PAID", "CLM OS", "CLM LI"]),
    "MV": dict(name="รถยนต์", cls="MO", catalog="PIVOT_202609_MV",
               cube="MV_AGGREGATE_FY", year="[VW PV MV AGGREGATE FY].[YEAR]",
               row1="[DIM CT SUBCLASS MV].[SUBCLASS DESC TH]", row2=None,
               measures=["ACT POLICY CNT", "ACT PREMIUM AMT", "EARNED PREMIUM",
                         "ACT SUM INS AMT", "CLAIM CNT", "CLAIM PAID",
                         "CLAIM OS", "CLAIM LI"]),
}

IDENTIFIER = re.compile(r"^\[[^\]\r\n]+\]\.\[[^\]\r\n]+\]$")
MEASURE_IDENTIFIER = re.compile(r"^\[Measures\]\.\[[^\]\r\n]+\]$")
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
    "YEAR": "ปีบัญชี", "FY YEAR": "ปีบัญชี",
    "MAIN NAME TH": "ประเภทงานหลัก", "SUB NAME TH": "ประเภทงานย่อย",
    "SUBCLASS DESC TH": "ประเภทกรมธรรม์", "GROUP DESC TH": "กลุ่มประเภท",
}


def local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


# UI and query semantics are intentionally separate: the id is a unique display
# item; "mdx" is only populated when its executable mapping has evidence.
WEBFORM = "https://onlinewebibs.oic.or.th/WebForm/Dashboard/Pivot/DynamicPivotCube.aspx"
SCHEMA_VERSION = 2
FIELD_COLUMNS = ("DATATABLE", "DATAFIELD", "CAPTION", "DATATYPE", "SUMMARYTYPE",
                 "FORMAT", "AREA", "CALCULATE", "GROUP_DIMENSION", "DIMENSION", "SEQ")
CATEGORY = {"มุมมอง": "DIMENSION", "ข้อมูล": "MEASURE"}


def field_id(code: str, record: dict) -> str:
    identity = [code] + [record.get(k) for k in
        ("GROUP_DIMENSION", "DIMENSION", "DATATABLE", "DATAFIELD",
         "CAPTION", "SUMMARYTYPE", "CALCULATE", "SEQ")]
    return hashlib.sha256(json.dumps(identity, ensure_ascii=False,
                                      sort_keys=True).encode("utf-8")).hexdigest()[:20]


def seeded_catalog(code: str) -> dict:
    profile = PROFILES[code]
    fields = []
    for index, name in enumerate(filter(None, (profile["year"], profile["row1"], profile["row2"]))):
        caption = name.rsplit("[", 1)[-1].rstrip("]")
        fields.append(dict(id=f"profile-d-{index}-{code}", mdx=name, display=caption,
                           caption=caption, type="DIMENSION", group_dimension="มุมมอง",
                           dimension="", source="PROFILE", status="PROFILE",
                           mapping_status="PROFILE", seq=index))
    for index, name in enumerate(profile["measures"]):
        mdx = f"[Measures].[{name}]"
        fields.append(dict(id=f"profile-m-{index}-{code}", mdx=mdx, display=name,
                           caption=name, type="MEASURE", group_dimension="ข้อมูล",
                           dimension="", source="PROFILE", status="PROFILE",
                           mapping_status="PROFILE", seq=index))
    return dict(code=code, web_class=profile["cls"], catalog=profile["catalog"],
                cube=profile["cube"], source="PROFILE", refreshed_at=None,
                fields=fields, view_count=len([f for f in fields if f["type"]=="DIMENSION"]),
                data_count=len([f for f in fields if f["type"]=="MEASURE"]),
                errors=[])


def cache_path(code: str) -> Path:
    return CATALOG_DIR / f"{code}.json"


def load_catalog(code: str) -> dict:
    if code not in PROFILES:
        raise ValueError("Unknown insurance category")
    fallback = seeded_catalog(code)
    try:
        cached = json.loads(cache_path(code).read_text(encoding="utf-8"))
        if cached.get("code") != code or cached.get("cube") != fallback["cube"]:
            return fallback
        if not isinstance(cached.get("fields"), list) or not cached["fields"]:
            return fallback
        fields = [f for f in cached["fields"] if isinstance(f, dict) and
                  f.get("type") in ("DIMENSION", "MEASURE") and f.get("id")]
        if not fields:
            return fallback
        cached["fields"] = fields
        cached["source"] = "CACHE"
        return cached
    except (OSError, ValueError, TypeError):
        return fallback


def save_catalog(catalog: dict) -> None:
    if catalog["code"] not in PROFILES or catalog["cube"] != PROFILES[catalog["code"]]["cube"]:
        raise ValueError("Invalid cube catalog")
    CATALOG_DIR.mkdir(parents=True, exist_ok=True)
    path = cache_path(catalog["code"])
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(catalog, ensure_ascii=False, indent=2), encoding="utf-8")
    temp.replace(path)


def valid_identifier(value: str) -> bool:
    return isinstance(value, str) and bool(
        IDENTIFIER.fullmatch(value) or MEASURE_IDENTIFIER.fullmatch(value))


def normalize_getstructure(code: str, raw: object, *, templates: list | None = None,
                           xmla: list | None = None) -> dict:
    """Pure parsing. CAPTION is *always* copied verbatim; never translated.

    Unmapped and calculated items remain visible but cannot be executed.
    """
    if not isinstance(raw, list) or not raw:
        raise ValueError("GetStructure must be a nonempty array")
    if code not in PROFILES:
        raise ValueError("Unknown cube category")
    profile = PROFILES[code]
    template_index: dict[tuple, set[tuple[str, bool]]] = {}
    for item in templates or []:
        if not isinstance(item, dict):
            continue
        if item.get("CLASS") != code or item.get("CUBE_NAME") != profile["cube"]:
            continue
        detail = item.get("TEMPLATE_DETAIL")
        try:
            detail = json.loads(detail) if isinstance(detail, str) else detail
        except (ValueError, TypeError):
            continue
        for part in detail if isinstance(detail, list) else []:
            if not isinstance(part, dict):
                continue
            mdx = part.get("dataField", "")
            if not valid_identifier(mdx):
                continue
            label = part.get("caption", "")
            folder = part.get("displayFolder") or ""
            area = part.get("area", "")
            # Display transformations (percent-of-total, etc.) cannot be
            # executed as their raw Measure without changing the meaning.
            value = (mdx, bool(part.get("summaryDisplayMode")))
            template_index.setdefault((label, folder, area), set()).add(value)
            template_index.setdefault((label, "", area), set()).add(value)
    xmla_index: dict[tuple[str,str],set[str]] = {}
    for item in xmla or []:
        if item.get("type") in ("DIMENSION", "MEASURE") and valid_identifier(item.get("mdx", "")):
            key = (item["type"], item.get("display") or "")
            xmla_index.setdefault(key, set()).add(item["mdx"])
    result = []
    for index, row in enumerate(raw):
        if not isinstance(row, dict) or "GROUP_DIMENSION" not in row:
            raise ValueError(f"Invalid GetStructure row {index}")
        group = row.get("GROUP_DIMENSION") or ""
        kind = CATEGORY.get(group)
        if not kind:
            continue
        caption = row.get("CAPTION")
        if not isinstance(caption, str) or not caption:
            caption = (row.get("DATAFIELD") or row.get("DATATABLE") or "(unnamed)")
        folder = row.get("DIMENSION") or ""
        area = "data" if kind == "MEASURE" else "row"
        candidates = set()
        for pos in ([area] if kind == "MEASURE" else ["row", "column", "filter"]):
            candidates |= template_index.get((caption, folder, pos), set())
        # A fallback by exact caption is allowed only if unambiguous. Do not
        # manufacture identifiers from DATATABLE/DATAFIELD.
        if not candidates:
            for pos in ([area] if kind == "MEASURE" else ["row", "column", "filter"]):
                candidates |= template_index.get((caption, "", pos), set())
        if not candidates:
            candidates = {(name, False) for name in xmla_index.get((kind, caption), set())}
        transformations_only = bool(candidates) and all(derived for _, derived in candidates)
        candidates = {mdx for mdx, mode in candidates if not mode and
                      (MEASURE_IDENTIFIER.fullmatch(mdx) if kind == "MEASURE"
                       else IDENTIFIER.fullmatch(mdx))}
        calculation = row.get("CALCULATE")
        summary = str(row.get("SUMMARYTYPE") or "").lower()
        # Calculations and display transformations cannot be replaced with
        # their raw base measure; this would silently export wrong results.
        derived = bool(calculation) or summary not in ("", "sum") or transformations_only
        mdx = next(iter(candidates)) if len(candidates) == 1 and not derived else None
        status = ("UNSUPPORTED" if derived else "MAPPED" if mdx else "UNVERIFIED")
        retained = {key: row.get(key) for key in FIELD_COLUMNS}
        result.append(dict(id=field_id(code, row), mdx=mdx, display=caption,
                           caption=caption, type=kind, group_dimension=group,
                           dimension=folder, source="GETSTRUCTURE",
                           status=status, mapping_status=status, seq=row.get("SEQ"),
                           details=retained))
    if not result:
        raise ValueError("No usable GetStructure rows")
    return dict(code=code, web_class=profile["cls"], cube=profile["cube"],
                catalog=profile["catalog"], source="LIVE", refreshed_at=datetime.now(timezone.utc).isoformat(),
                fields=result, view_count=sum(f["type"]=="DIMENSION" for f in result),
                data_count=sum(f["type"]=="MEASURE" for f in result), errors=[])


def parse_webform_response(response) -> list:
    if response.status_code != 200:
        raise ValueError(f"WebForm HTTP {response.status_code}")
    try:
        wrapper = response.json()
        if not isinstance(wrapper, dict) or "d" not in wrapper:
            raise ValueError("WebForm returned no d property")
        payload = wrapper["d"]
        value = json.loads(payload) if isinstance(payload, str) else payload
        if not isinstance(value, list):
            raise ValueError("WebForm d is not a list")
        return value
    except (TypeError, ValueError, KeyError) as exc:
        raise ValueError("Invalid WebForm JSON response") from exc


def make_discover_request(code: str, request_type: str) -> bytes:
    if request_type not in ("MDSCHEMA_DIMENSIONS", "MDSCHEMA_HIERARCHIES",
                            "MDSCHEMA_LEVELS", "MDSCHEMA_MEASURES", "MDSCHEMA_MEMBERS"):
        raise ValueError("Unsupported XMLA Discover request")
    p = PROFILES[code]
    root = ET.Element(f"{{{SOAP_NS}}}Envelope")
    body = ET.SubElement(root, f"{{{SOAP_NS}}}Body")
    discover = ET.SubElement(body, f"{{{XMLA_NS}}}Discover")
    ET.SubElement(discover, f"{{{XMLA_NS}}}RequestType").text = request_type
    restrictions = ET.SubElement(ET.SubElement(discover, f"{{{XMLA_NS}}}Restrictions"),
                                f"{{{XMLA_NS}}}RestrictionList")
    ET.SubElement(restrictions, f"{{{XMLA_NS}}}CATALOG_NAME").text = p["catalog"]
    ET.SubElement(restrictions, f"{{{XMLA_NS}}}CUBE_NAME").text = p["cube"]
    properties = ET.SubElement(ET.SubElement(discover, f"{{{XMLA_NS}}}Properties"),
                               f"{{{XMLA_NS}}}PropertyList")
    ET.SubElement(properties, f"{{{XMLA_NS}}}Catalog").text = p["catalog"]
    ET.SubElement(properties, f"{{{XMLA_NS}}}LocaleIdentifier").text = "9"
    return ET.tostring(root, encoding="utf-8")


def discover_fields(code: str, token: str, cookie: str, *, timeout: int = 30) -> dict:
    """Explicit/live catalog refresh; GetStructure drives every visible label."""
    import requests
    if code not in PROFILES:
        raise ValueError("Unknown insurance category")
    if not cookie:
        raise ValueError("Valid Cookie required for WebForm metadata")
    p = PROFILES[code]
    referer = f"{WEBFORM}?Class={p['cls']}"
    headers = {"Content-Type": "application/json; charset=utf-8",
               "Accept": "application/json, text/javascript, */*; q=0.01",
               "X-Requested-With": "XMLHttpRequest",
               "Origin": "https://onlinewebibs.oic.or.th", "Referer": referer,
               "Cookie": cookie.removeprefix("Cookie:").strip()}
    errors = []
    with requests.Session() as session:
        def post_webform(method: str, payload: dict) -> list:
            response = session.post(f"{WEBFORM}/{method}", json=payload, headers=headers,
                                    timeout=(10, timeout), allow_redirects=False)
            return parse_webform_response(response)
        # Never reuse an XMLA Bearer header for ASP.NET WebForm without proof.
        structure = post_webform("GetStructure", {"Template": [code, p["cube"]]})
        templates = []
        try:
            templates = post_webform("GET_MENUTEMPLATE_LIST", {"Template": [p["cls"]]})
        except (ValueError, requests.RequestException) as exc:
            errors.append(f"Templates: {exc}")
        xmla = []
        if token:
            xml_headers = {"Authorization": "Bearer " + token.removeprefix("Bearer ").strip(),
                           "Content-Type": "text/xml", "Accept": "text/plain, */*; q=0.01",
                           "Origin": "https://onlinewebibs.oic.or.th", "Referer": referer,
                           "Cookie": headers["Cookie"]}
            for kind in ("MDSCHEMA_HIERARCHIES", "MDSCHEMA_MEASURES"):
                try:
                    response = session.post(PROXY_URL, data=make_discover_request(code, kind),
                                            headers=xml_headers, timeout=(10, timeout),
                                            allow_redirects=False)
                    if response.status_code != 200:
                        raise ValueError(f"HTTP {response.status_code}")
                    root = ET.fromstring(response.content)
                    if any(local_name(e.tag) == "Fault" for e in root.iter()):
                        raise ValueError("SOAP Fault")
                    for row in root.iter():
                        if local_name(row.tag) != "row":
                            continue
                        attrs = {local_name(e.tag): (e.text or "") for e in row}
                        mdx = attrs.get("HIERARCHY_UNIQUE_NAME") if kind.endswith("HIERARCHIES") else attrs.get("MEASURE_UNIQUE_NAME")
                        display = attrs.get("HIERARCHY_CAPTION") if kind.endswith("HIERARCHIES") else attrs.get("MEASURE_CAPTION")
                        typ = "DIMENSION" if kind.endswith("HIERARCHIES") else "MEASURE"
                        if valid_identifier(mdx):
                            xmla.append(dict(mdx=mdx, display=display, type=typ))
                except (requests.RequestException, ValueError, ET.ParseError) as exc:
                    errors.append(f"{kind}: {exc}")
    catalog = normalize_getstructure(code, structure, templates=templates, xmla=xmla)
    catalog["errors"] = errors
    catalog["source"] = "LIVE"
    save_catalog(catalog)
    return catalog


def validate_query(config: dict, catalog: dict | None = None) -> dict:
    code = config.get("code")
    if code not in PROFILES:
        raise ValueError("Select AH/FI/MR/MI/MV")
    p = PROFILES[code]
    if config.get("catalog") != p["catalog"] or config.get("cube") != p["cube"]:
        raise ValueError("Catalog/Cube do not match the selected category")
    version = int(config.get("schema_version", 1))
    if version not in (1, SCHEMA_VERSION):
        raise ValueError("Unsupported query schema version")
    x, y, measures = (config.get(k) for k in ("x", "y", "measures"))
    if not all(isinstance(v, list) and v for v in (x, y, measures)):
        raise ValueError("Select X/ROWS, Y/COLUMNS and at least one Measure")
    if len(x) + len(y) > 5:
        raise ValueError("Maximum five selected hierarchies for safe CrossJoin")
    if len(set(x+y)) != len(x+y) or len(set(measures)) != len(measures):
        raise ValueError("Duplicate hierarchy or Measure in query")
    for name in x+y:
        if not IDENTIFIER.fullmatch(name):
            raise ValueError("Invalid hierarchy identifier: " + str(name))
    for name in measures:
        if not MEASURE_IDENTIFIER.fullmatch(name):
            raise ValueError("Invalid Measure identifier: " + str(name))
    if version == SCHEMA_VERSION:
        ids = config.get("selection_ids", {})
        snapshot = config.get("selected_fields", [])
        if not isinstance(ids, dict) or not isinstance(snapshot, list):
            raise ValueError("Missing applied selection identities")
        source = {f.get("id"): f for f in snapshot if isinstance(f, dict)}
        if len(source) != len(snapshot):
            raise ValueError("Duplicate selected display identity")
        for axis, names in (("x", x), ("y", y), ("measures", measures)):
            identities = ids.get(axis, [])
            if len(identities) != len(names):
                raise ValueError(f"Missing {axis} selection identity")
            for identity, mdx in zip(identities, names):
                field = source.get(identity)
                if not field or field.get("mdx") != mdx or field.get("status") not in ("MAPPED", "VERIFIED"):
                    raise ValueError(f"Unmapped/unsupported field in {axis}: {identity}")
                if field.get("type") != ("MEASURE" if axis == "measures" else "DIMENSION"):
                    raise ValueError("Field type mismatch")
        if len(set(ids["x"]+ids["y"])) != len(x+y):
            raise ValueError("Duplicate display field on both axes")
    else:
        cat = catalog or load_catalog(code)
        allowed = {(f.get("mdx"), f.get("type")) for f in cat["fields"] + seeded_catalog(code)["fields"]}
        for name in x+y:
            if (name, "DIMENSION") not in allowed:
                raise ValueError("Unknown hierarchy: " + name)
        for name in measures:
            if (name, "MEASURE") not in allowed:
                raise ValueError("Unknown Measure: " + name)
    try:
        workers = int(config.get("workers", 4))
        xc = int(config.get("x_chunk", 1))
        yc = int(config.get("y_chunk", 20))
    except (TypeError, ValueError) as exc:
        raise ValueError("Workers and chunk sizes must be integers") from exc
    if not 1 <= workers <= 6 or not 1 <= xc <= 20 or not 1 <= yc <= 100:
        raise ValueError("Workers 1-6; X chunks 1-20; Y chunks 1-100")
    if config.get("filters"):
        raise ValueError("Filters require verified metadata; unsupported in this release")
    for secret in ("cookie", "token", "jwt", "authorization", "OIC_COOKIE", "OIC_BEARER_TOKEN"):
        if secret in config:
            raise ValueError("Authentication secrets must not be in query configuration")
    return config


def default_query(code: str, *, workers: int = 1, schema_version: int = 1) -> dict:
    p = PROFILES[code]
    # Preserve historical v1 defaults for archived runs. The new designer makes
    # a separate v2 configuration from mapped field identities.
    return dict(code=code, cube=p["cube"], catalog=p["catalog"],
                x=[p["year"]] if p["year"] else [], y=[p["row1"]],
                measures=[f"[Measures].[{name}]" for name in p["measures"]],
                filters=[], workers=workers, x_chunk=1, y_chunk=20,
                **({"schema_version": schema_version} if schema_version != 1 else {}))


def axis_set(fields: list[str], offset: int, size: int) -> str:
    if not fields or offset < 0 or size < 1:
        raise ValueError("Invalid pagination")
    members = [field + ".AllMembers" for field in fields]
    result = members[0]
    for other in members[1:]:
        result = f"CrossJoin({result},{other})"
    return f"Subset({result},{offset},{size})"


def build_mdx(config: dict, x_offset: int = 0, y_offset: int = 0) -> str:
    validate_query(config)
    x = axis_set(config["x"], x_offset, int(config.get("x_chunk", 1)))
    y = axis_set(config["y"], y_offset, int(config.get("y_chunk", 20)))
    measures = "{" + ",".join(config["measures"]) + "}"
    if int(config.get("schema_version", 1)) == SCHEMA_VERSION:
        # User-facing X = ROWS / Axis1, Y = COLUMNS / Axis0.
        columns, rows = y, f"CrossJoin({x},{measures})"
    else:
        # Only for previously stored v1 runs: X = Axis0, Y = Axis1.
        columns, rows = x, f"CrossJoin({y},{measures})"
    return (f"SELECT {columns} DIMENSION PROPERTIES PARENT_UNIQUE_NAME,"
            f"HIERARCHY_UNIQUE_NAME,MEMBER_VALUE ON COLUMNS,"
            f"{rows} DIMENSION PROPERTIES PARENT_UNIQUE_NAME,"
            f"HIERARCHY_UNIQUE_NAME,MEMBER_VALUE ON ROWS "
            f"FROM [{PROFILES[config['code']]['cube']}] "
            "CELL PROPERTIES VALUE,FORMAT_STRING,LANGUAGE")


def build_soap(config: dict, x_offset: int = 0, y_offset: int = 0) -> bytes:
    mdx = build_mdx(config, x_offset, y_offset)
    env = ET.Element(f"{{{SOAP_NS}}}Envelope")
    body = ET.SubElement(env, f"{{{SOAP_NS}}}Body")
    execute = ET.SubElement(body, f"{{{XMLA_NS}}}Execute")
    command = ET.SubElement(execute, f"{{{XMLA_NS}}}Command")
    ET.SubElement(command, f"{{{XMLA_NS}}}Statement").text = mdx
    props = ET.SubElement(ET.SubElement(execute, f"{{{XMLA_NS}}}Properties"),
                          f"{{{XMLA_NS}}}PropertyList")
    for key, value in {"Catalog": config["catalog"], "ShowHiddenCubes": "True",
                       "Timeout": "3600", "LocaleIdentifier": "9"}.items():
        ET.SubElement(props, f"{{{XMLA_NS}}}{key}").text = value
    return ET.tostring(env, encoding="utf-8")


def query_hash(config: dict) -> str:
    canonical = {k: v for k, v in config.items() if k not in ("run_id", "created_at", "output")}
    return hashlib.sha256(json.dumps(canonical, ensure_ascii=False, sort_keys=True).encode()).hexdigest()[:16]


def parse_xmla(content: bytes) -> dict:
    """Decode all Axis0 and Axis1 tuple members; CellOrdinal is column-major."""
    root = ET.fromstring(content)
    for element in root.iter():
        if local_name(element.tag) == "Fault":
            raise ValueError("SOAP Fault: " + " ".join((e.text or "") for e in element.iter())[:500])
    dataset = root.find(f".//{{{MD_NS}}}root")
    if dataset is None:
        raise ValueError("XMLA multidimensional dataset missing")
    axes = {}
    for axis in dataset.findall(f"./{{{MD_NS}}}Axes/{{{MD_NS}}}Axis"):
        members = []
        for tup in axis.findall(f"{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple"):
            members.append([dict(hierarchy=member.get("Hierarchy", ""),
                                 unique_name=member.findtext(f"{{{MD_NS}}}UName") or "",
                                 caption=member.findtext(f"{{{MD_NS}}}Caption") or "",
                                 level=member.findtext(f"{{{MD_NS}}}LNum") or "")
                            for member in tup.findall(f"{{{MD_NS}}}Member")])
        axes[axis.get("name", "")] = members
    if "Axis0" not in axes or "Axis1" not in axes:
        raise ValueError("Expected Axis0 and Axis1")
    cells = {}
    for cell in dataset.findall(f"./{{{MD_NS}}}CellData/{{{MD_NS}}}Cell"):
        value = cell.find(f"{{{MD_NS}}}Value")
        if value is not None:
            cells[int(cell.get("CellOrdinal", "-1"))] = value.text
    updates = next((e.text for e in root.iter() if local_name(e.tag) == "LastDataUpdate"), None)
    return dict(axes=axes, cells=cells, last_data_update=updates)


def excel_records(config: dict, data: dict):
    """Long-format records keep every hierarchy and Measure without aggregation."""
    columns = data["axes"]["Axis0"]
    rows = data["axes"]["Axis1"]
    width = len(columns)
    modern = int(config.get("schema_version", 1)) == SCHEMA_VERSION
    for ri, row in enumerate(rows):
        measure = next((m for m in row if m["hierarchy"] == "[Measures]"), None)
        if measure is None:
            raise ValueError("Missing Measures member on Axis1")
        dimensions = [m for m in row if m["hierarchy"] != "[Measures]"]
        for ci, column in enumerate(columns):
            ordinal = ci + width * ri
            if ordinal not in data["cells"]:
                continue
            yield dict(X=dimensions if modern else column,
                       Y=column if modern else dimensions,
                       MEASURE=measure["unique_name"],
                       VALUE=data["cells"][ordinal],
                       LAST_DATA_UPDATE=data["last_data_update"])
