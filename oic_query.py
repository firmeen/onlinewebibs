#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Shared OIC cube catalog, validated query definitions, MDX and XMLA parsing.

No network traffic occurs on import or while building previews.
"""
from __future__ import annotations

import hashlib
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
    "FI": dict(name="ทรัพย์สิน", cls="FI", catalog="PIVOT_202609_FI",
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


def unique(fields: list[dict]) -> list[dict]:
    return list({item["mdx"]: item for item in fields}.values())


def seeded_catalog(code: str) -> dict:
    profile = PROFILES[code]
    dimensions = [profile["year"], profile["row1"], profile["row2"]]
    fields = []
    for name in dimensions:
        if not name:
            continue
        label = name.rsplit("[", 1)[-1].rstrip("]")
        fields.append(dict(mdx=name, display=KNOWN_DIMENSION_NAMES.get(label, label),
                           type="DIMENSION", source="PROFILE", status="KNOWN_PROFILE",
                           group=name.split("].[", 1)[0][1:]))
    for name in profile["measures"]:
        fields.append(dict(mdx=f"[Measures].[{name}]",
                           display=KNOWN_MEASURE_NAMES.get(name, name),
                           type="MEASURE", source="PROFILE", status="KNOWN_PROFILE",
                           group="Measures"))
    return dict(code=code, catalog=profile["catalog"], cube=profile["cube"],
                source="PROFILE", refreshed_at=None, fields=unique(fields))


def cache_path(code: str) -> Path:
    return CATALOG_DIR / f"{code}.json"


def load_catalog(code: str) -> dict:
    if code not in PROFILES:
        raise ValueError("Unknown insurance category")
    fallback = seeded_catalog(code)
    path = cache_path(code)
    try:
        cached = json.loads(path.read_text(encoding="utf-8"))
        if cached.get("code") != code or cached.get("cube") != fallback["cube"]:
            return fallback
        fields = [f for f in cached.get("fields", []) if isinstance(f, dict) and
                  f.get("type") in ("DIMENSION", "MEASURE") and
                  valid_identifier(f.get("mdx", ""))]
        if not fields:
            return fallback
        # Preserve profile fields if offline discovery is incomplete.
        cached["fields"] = unique(fallback["fields"] + fields)
        cached["source"] = "CACHE"
        return cached
    except (OSError, ValueError, TypeError):
        return fallback


def save_catalog(catalog: dict) -> None:
    code = catalog["code"]
    CATALOG_DIR.mkdir(parents=True, exist_ok=True)
    path = cache_path(code)
    tmp = path.with_suffix(".tmp")
    tmp.write_text(json.dumps(catalog, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(path)


def valid_identifier(value: str) -> bool:
    return bool(IDENTIFIER.fullmatch(value) or MEASURE_IDENTIFIER.fullmatch(value))


def make_discover_request(code: str, request_type: str) -> bytes:
    if request_type not in ("MDSCHEMA_HIERARCHIES", "MDSCHEMA_LEVELS", "MDSCHEMA_MEASURES"):
        raise ValueError("Unsupported XMLA Discover request")
    p = PROFILES[code]
    root = ET.Element(f"{{{SOAP_NS}}}Envelope")
    body = ET.SubElement(root, f"{{{SOAP_NS}}}Body")
    discover = ET.SubElement(body, f"{{{XMLA_NS}}}Discover")
    ET.SubElement(discover, f"{{{XMLA_NS}}}RequestType").text = request_type
    restrict = ET.SubElement(ET.SubElement(discover, f"{{{XMLA_NS}}}Restrictions"),
                             f"{{{XMLA_NS}}}RestrictionList")
    ET.SubElement(restrict, f"{{{XMLA_NS}}}CATALOG_NAME").text = p["catalog"]
    ET.SubElement(restrict, f"{{{XMLA_NS}}}CUBE_NAME").text = p["cube"]
    props = ET.SubElement(ET.SubElement(discover, f"{{{XMLA_NS}}}Properties"),
                          f"{{{XMLA_NS}}}PropertyList")
    ET.SubElement(props, f"{{{XMLA_NS}}}Catalog").text = p["catalog"]
    ET.SubElement(props, f"{{{XMLA_NS}}}LocaleIdentifier").text = "9"
    return ET.tostring(root, encoding="utf-8")


def discover_fields(code: str, token: str, cookie: str, *, timeout: int = 30) -> dict:
    """Optional explicit refresh; never called while typing or building a preview."""
    import requests
    token = token.removeprefix("Bearer ").strip()
    if not token or not cookie:
        raise ValueError("Authorized Cookie and JWT are required for live metadata")
    base = seeded_catalog(code)
    fields = base["fields"]
    headers = {
        "Authorization": "Bearer " + token, "Cookie": cookie,
        "Content-Type": "text/xml", "Accept": "text/plain, */*; q=0.01",
        "Origin": "https://onlinewebibs.oic.or.th",
        "X-Requested-With": "XMLHttpRequest",
        "Referer": ("https://onlinewebibs.oic.or.th/WebForm/Dashboard/Pivot/"
                    f"DynamicPivotCube.aspx?Class={PROFILES[code]['cls']}"),
    }
    errors = []
    with requests.Session() as session:
        for kind in ("MDSCHEMA_HIERARCHIES", "MDSCHEMA_MEASURES", "MDSCHEMA_LEVELS"):
            try:
                response = session.post(PROXY_URL, data=make_discover_request(code, kind),
                                        headers=headers, timeout=(10, timeout),
                                        allow_redirects=False)
                if response.status_code != 200:
                    raise ValueError(f"{kind}: HTTP {response.status_code}")
                root = ET.fromstring(response.content)
                faults = [e for e in root.iter() if local_name(e.tag) == "Fault"]
                if faults:
                    raise ValueError(f"{kind}: SOAP Fault")
                rows = [e for e in root.iter() if local_name(e.tag) == "row"]
                if not rows:
                    raise ValueError(f"{kind}: no metadata rows")
                for row in rows:
                    attrs = {local_name(e.tag): (e.text or "") for e in row}
                    if kind == "MDSCHEMA_HIERARCHIES":
                        name = attrs.get("HIERARCHY_UNIQUE_NAME", "")
                        if valid_identifier(name) and not name.startswith("[Measures]"):
                            fields.append(dict(mdx=name, display=attrs.get("HIERARCHY_CAPTION") or
                                               name.rsplit("[", 1)[-1].rstrip("]"),
                                               type="DIMENSION", group=attrs.get("DIMENSION_UNIQUE_NAME", ""),
                                               source="XMLA_HIERARCHIES", status="DISCOVERED"))
                    if kind == "MDSCHEMA_MEASURES":
                        name = attrs.get("MEASURE_UNIQUE_NAME", "")
                        if MEASURE_IDENTIFIER.fullmatch(name):
                            fields.append(dict(mdx=name, display=attrs.get("MEASURE_CAPTION") or
                                               name.rsplit("[", 1)[-1].rstrip("]"),
                                               type="MEASURE", group="Measures",
                                               source="XMLA_MEASURES", status="DISCOVERED"))
            except (requests.RequestException, ValueError, ET.ParseError) as exc:
                errors.append(str(exc))
    # Metadata discovery is not proof that executing every combination will work.
    base.update(fields=unique(fields), source="LIVE" if len(errors) < 3 else "PROFILE",
                refreshed_at=datetime.now(timezone.utc).isoformat(), errors=errors)
    save_catalog(base)
    return base


def validate_query(config: dict, catalog: dict | None = None) -> dict:
    code = config.get("code", "")
    if code not in PROFILES:
        raise ValueError("Select AH, FI, MR, MI or MV")
    cat = catalog or load_catalog(code)
    p = PROFILES[code]
    if config.get("cube") != p["cube"] or config.get("catalog") != p["catalog"]:
        raise ValueError("Cube/Catalog do not match the selected insurance category")
    x = config.get("x", [])
    y = config.get("y", [])
    measures = config.get("measures", [])
    if not isinstance(x, list) or not isinstance(y, list) or not x or not y:
        raise ValueError("Select at least one X and one Y dimension")
    if len(x) + len(y) > 5:
        raise ValueError("Maximum five selected dimensions across X and Y for safe preview")
    if len(x + y) != len(set(x + y)):
        raise ValueError("Duplicate hierarchy across X/Y is not supported")
    allowed = {f["mdx"]: f["type"] for f in cat["fields"]}
    for item in x + y:
        if not IDENTIFIER.fullmatch(item) or allowed.get(item) != "DIMENSION":
            raise ValueError(f"Dimension unavailable in selected catalog: {item}")
    if not measures or len(measures) != len(set(measures)):
        raise ValueError("Select at least one distinct measure")
    for item in measures:
        if not MEASURE_IDENTIFIER.fullmatch(item) or allowed.get(item) != "MEASURE":
            raise ValueError(f"Measure unavailable in selected catalog: {item}")
    workers = int(config.get("workers", 1))
    xs = int(config.get("x_chunk", 1))
    ys = int(config.get("y_chunk", 20))
    if not 1 <= workers <= 6 or not 1 <= xs <= 20 or not 1 <= ys <= 100:
        raise ValueError("Workers: 1–6, X chunk: 1–20, Y chunk: 1–100")
    if config.get("filters"):
        raise ValueError("Filter expressions are not enabled without validated metadata")
    return config


def default_query(code: str, *, workers: int = 1) -> dict:
    p = PROFILES[code]
    return dict(code=code, cube=p["cube"], catalog=p["catalog"],
                x=[p["year"]] if p["year"] else [], y=[p["row1"]],
                measures=[f"[Measures].[{name}]" for name in p["measures"]],
                filters=[], workers=workers, x_chunk=1, y_chunk=20)


def axis_set(fields: list[str], offset: int, size: int) -> str:
    if offset < 0 or size < 1:
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
    m = "{" + ",".join(config["measures"]) + "}"
    rows = f"CrossJoin({y},{m})"
    return (f"SELECT {x} DIMENSION PROPERTIES PARENT_UNIQUE_NAME,"
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
    """Long-format rows keep all X/Y dimensions, members and the selected Measure."""
    columns = data["axes"]["Axis0"]
    rows = data["axes"]["Axis1"]
    width = len(columns)
    for ri, row in enumerate(rows):
        measure = next((m for m in row if m["hierarchy"] == "[Measures]"), None)
        if measure is None:
            raise ValueError("No [Measures] member on Axis1")
        y = [m for m in row if m["hierarchy"] != "[Measures]"]
        for ci, x in enumerate(columns):
            ordinal = ci + width * ri
            if ordinal not in data["cells"]:
                continue
            value = data["cells"][ordinal]
            yield {
                "X": x, "Y": y, "MEASURE": measure["unique_name"],
                "VALUE": value, "LAST_DATA_UPDATE": data["last_data_update"],
            }
