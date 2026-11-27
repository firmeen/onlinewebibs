"""Prepare immutable, executable XMLA requests before starting any worker."""
from __future__ import annotations

import itertools
from copy import deepcopy
from concurrent.futures import ThreadPoolExecutor
import json
import math
import re
import time
from decimal import Decimal, InvalidOperation
import xml.etree.ElementTree as ET
from pathlib import Path

import oic_query as query
from oic_all_dim_db import canonical, digest

SOAP_NS = query.SOAP_NS
XMLA_NS = query.XMLA_NS
MD_NS = query.MD_NS
CATEGORIES = ("AH", "FI", "MR", "MI", "MV")
YEAR_TYPES = ("FY", "UY", "AY")
PLANNER_VERSION = 2
MAX_CELLS = 20000
PROBE_BUDGET = 60
DISCOVERY_PAGE = 100
MAX_DISCOVERY_PAGE = 500
MIN_DISCOVERY_PAGE = 25
MAX_DISCOVERY_PAGES = 10000
DEFAULT_TIMEOUT = 180


def local(tag):
    return tag.rsplit("}", 1)[-1]


def integer(value, default=None):
    try:
        number=Decimal(str(value))
        return int(number) if number.is_finite() and number==number.to_integral_value() else default
    except (InvalidOperation,TypeError,ValueError,OverflowError):
        return default


def checked_root(data):
    try:
        root = ET.fromstring(data)
    except ET.ParseError as exc:
        raise ValueError("Malformed XMLA XML") from exc
    if any(local(e.tag) in ("Fault", "Error") for e in root.iter()):
        raise ValueError("XMLA fault/error")
    return root


def emit_event(emit, event, **data):
    if emit is not None:
        emit(event, **data)


def load_all_catalogs(emit=None):
    result = {}
    for code, kind in itertools.product(CATEGORIES, YEAR_TYPES):
        key = f"{code}/{kind}"
        emit_event(emit, "phase", phase="CATALOG", cube=key, detail="Reading local catalog")
        status = query.cube_status(code, kind)
        catalog = query.load_catalog(code, kind)
        profile = dict(query.cube_profile(code, kind, discovery=True))
        profile.update(code=code, cube_type=kind, availability=status["status"])
        if catalog.get("source") != "PROFILE":
            profile["cube"] = catalog.get("cube") or profile["cube"]
            profile["catalog"] = catalog.get("catalog") or profile["catalog"]
        fields = catalog.get("fields", [])
        # Preserve the original report (including unresolved identifiers/formulas).
        report = query.load_verification_report(code, kind)
        if report:
            original = {}
            for f in report.get("fields", []):
                try:
                    original[str(query._field_identity(code,kind,f))]=f
                except (KeyError,TypeError,ValueError):
                    continue
            fields = [dict(f, original_report=original.get(str(f.get("id")))) for f in fields]
        result[key] = (profile, catalog, fields)
        emit_event(emit, "catalog", cube=key, source=catalog.get("source") or "UNKNOWN",
                   fields=len(fields), availability=profile["availability"])
    return result


def classify_fields(fields):
    dimensions, measures = {}, {}
    for field in fields:
        mdx = field.get("mdx")
        if field.get("status") not in ("VERIFIED", "MAPPED") or not mdx:
            continue
        typ = field.get("type")
        if typ == "DIMENSION" and query.IDENTIFIER.fullmatch(mdx) and not mdx.startswith("[Measures]."):
            dimensions.setdefault(mdx, field)
        elif typ == "MEASURE" and query.MEASURE_IDENTIFIER.fullmatch(mdx):
            details = field.get("details") or {}
            calculated = (str(details.get("DATAFIELD", "")).upper() == "CALCULATE"
                or str(details.get("CALCULATE") or "").lower().strip() not in ("", "0", "false", "none", "null")
                or str(details.get("SUMMARYTYPE") or "").lower().strip() not in ("", "sum")
                or mdx.casefold() == "[measures].[calculate]")
            if not calculated:
                measures.setdefault(mdx, field)
    return dimensions, measures


def resolve_year_hierarchy(profile, dimensions):
    year = profile.get("year")
    if year in dimensions:
        return year
    target = profile["cube_type"] + " YEAR"
    choices = [h for h in dimensions if h.rsplit("[", 1)[-1].rstrip("]").upper() == target]
    if not choices:
        choices = [h for h in dimensions if h.rsplit("[", 1)[-1].rstrip("]").upper() == "YEAR"]
    if len(choices) != 1:
        raise ValueError("Missing or ambiguous year hierarchy")
    return choices[0]


def build_discover_request(profile, request_type, restrictions=None):
    if request_type not in ("MDSCHEMA_DIMENSIONS","MDSCHEMA_HIERARCHIES","MDSCHEMA_LEVELS",
                            "MDSCHEMA_MEMBERS","MDSCHEMA_MEASURES","MDSCHEMA_MEASUREGROUP_DIMENSIONS"):
        raise ValueError("Invalid Discover rowset")
    if request_type=="MDSCHEMA_MEMBERS" and not (restrictions or {}).get("HIERARCHY_UNIQUE_NAME"):
        raise ValueError("Member discovery requires a hierarchy restriction")
    env=ET.Element(f"{{{SOAP_NS}}}Envelope")
    body=ET.SubElement(env,f"{{{SOAP_NS}}}Body")
    node=ET.SubElement(body,f"{{{XMLA_NS}}}Discover")
    ET.SubElement(node,f"{{{XMLA_NS}}}RequestType").text=request_type
    rs=ET.SubElement(ET.SubElement(node,f"{{{XMLA_NS}}}Restrictions"),f"{{{XMLA_NS}}}RestrictionList")
    for k,v in {"CATALOG_NAME":profile["catalog"],"CUBE_NAME":profile["cube"],**(restrictions or {})}.items():
        ET.SubElement(rs,f"{{{XMLA_NS}}}{k}").text=v
    props=ET.SubElement(ET.SubElement(node,f"{{{XMLA_NS}}}Properties"),f"{{{XMLA_NS}}}PropertyList")
    for k,v in {"Catalog":profile["catalog"],"LocaleIdentifier":"9"}.items():
        ET.SubElement(props,f"{{{XMLA_NS}}}{k}").text=v
    return ET.tostring(env,encoding="utf-8")


def parse_discover_rows(data):
    root=checked_root(data)
    rowset=root.find(".//{urn:schemas-microsoft-com:xml-analysis:rowset}root")
    if rowset is None:
        raise ValueError("Missing XMLA rowset")
    return [{local(c.tag):c.text or "" for c in r} for r in rowset if local(r.tag)=="row"]


def discover_rows(transport,profile,kind,restrictions=None,emit=None):
    phase="MEMBER_DISCOVERY" if kind=="MDSCHEMA_MEMBERS" else "METADATA_DISCOVERY"
    detail=kind
    if restrictions:
        detail+=" | "+" / ".join(str(v) for v in restrictions.values())
    emit_event(emit,"phase",phase=phase,detail=detail)
    data=transport(profile,build_discover_request(profile,kind,restrictions),60 if kind=="MDSCHEMA_MEMBERS" else 30)
    emit_event(emit,"phase",phase=phase,detail="Parsing "+detail)
    rows=parse_discover_rows(data)
    emit_event(emit,"activity",message=f"{kind} | {len(rows):,} metadata rows")
    return rows


def mdx_set(values):
    # Allow empty *key* components (&[]) and escaped closing brackets,
    # not empty identifiers or arbitrary MDX expressions.
    identifier=r"\[(?:[^\]\x00-\x1f]|\]\])+\]"
    key=r"&\[(?:[^\]\x00-\x1f]|\]\])*\]"
    pattern=identifier+r"(?:\.(?:"+identifier+r"|"+key+r"(?:"+key+r")*))*"
    if not values or any(not isinstance(v,str) or not re.fullmatch(pattern,v) for v in values):
        raise ValueError("INVALID_MEMBER_UNIQUE_NAME: invalid bracketed identifier/key")
    return "{" + ",".join(values) + "}"


def build_mdx(profile, scope):
    years=mdx_set(scope["years"])
    measures=mdx_set(scope["measures"])
    rows=measures if not scope["hierarchy"] else f"CrossJoin({mdx_set(scope['members'])},{measures})"
    cube=profile["cube"].replace("]","]]")
    return f"SELECT {years} ON COLUMNS, {rows} ON ROWS FROM [{cube}] CELL PROPERTIES VALUE, FORMATTED_VALUE, FORMAT_STRING, LANGUAGE"


def build_soap(profile,mdx):
    env=ET.Element(f"{{{SOAP_NS}}}Envelope")
    execute=ET.SubElement(ET.SubElement(env,f"{{{SOAP_NS}}}Body"),f"{{{XMLA_NS}}}Execute")
    ET.SubElement(ET.SubElement(execute,f"{{{XMLA_NS}}}Command"),f"{{{XMLA_NS}}}Statement").text=mdx
    props=ET.SubElement(ET.SubElement(execute,f"{{{XMLA_NS}}}Properties"),f"{{{XMLA_NS}}}PropertyList")
    for k,v in {"Catalog":profile["catalog"],"ShowHiddenCubes":"True",
                "SspropInitAppName":"Microsoft SQL Server Management Studio",
                "Timeout":"3600","LocaleIdentifier":"9"}.items():
        ET.SubElement(props,f"{{{XMLA_NS}}}{k}").text=v
    return ET.tostring(env,encoding="utf-8")


def normalize_member(row):
    name=row.get("MEMBER_UNIQUE_NAME")
    if not name:
        raise ValueError("MISSING_MEMBER_UNIQUE_NAME")
    mdx_set([name])
    return {"unique_name":name,"caption":row.get("MEMBER_CAPTION") or row.get("MEMBER_NAME"),
            "parent":row.get("PARENT_UNIQUE_NAME") or None,
            "level":row.get("LEVEL_UNIQUE_NAME") or None,
            "level_number":integer(row.get("LEVEL_NUMBER")),
            "member_type":integer(row.get("MEMBER_TYPE")),
            "children_count":integer(row.get("CHILDREN_CARDINALITY")),
            "is_all_member":integer(row.get("MEMBER_TYPE"))==2,
            "is_calculated_member":integer(row.get("MEMBER_TYPE"))==4,
            "metadata":row}


def deduplicate_members(rows, hierarchy):
    result={}
    for row in rows:
        actual=row.get("HIERARCHY_UNIQUE_NAME")
        if not actual or actual.casefold()!=hierarchy.casefold():
            raise ValueError("Member restriction not honored by server: missing or mismatched hierarchy")
        member=normalize_member(row)
        previous=result.get(member["unique_name"])
        if previous is not None and (previous["parent"],previous["level"])!=(member["parent"],member["level"]):
            raise ValueError("CONFLICTING_MEMBER_METADATA: duplicate unique name")
        result[member["unique_name"]]=member
    return sorted(result.values(),key=lambda m:m["unique_name"])


def raise_if_fatal(exc):
    if getattr(exc,"kind","") in ("USER_STOP","AUTH_COOKIE_EXPIRED","AUTH_TOKEN_EXPIRED","AUTH_ENDPOINT_ERROR",
                                  "ACCESS_DENIED","TLS_ERROR","RATE_LIMITED","SERVER_UNAVAILABLE",
                                  "DISK_FULL","STORAGE_ERROR","NETWORK_ERROR","DNS_ERROR","CONNECTION_RESET"):
        raise exc


def count_batch(transport,profile,hierarchies):
    cube=profile["cube"].replace("]","]]")
    if any(not query.IDENTIFIER.fullmatch(h) for h in hierarchies):
        raise ValueError("Invalid hierarchy identifier")
    aliases={f"[Measures].[OICInventoryCount{i}]":h for i,h in enumerate(hierarchies)}
    definitions=" ".join(f"MEMBER {alias} AS Count({h}.AllMembers)" for alias,h in aliases.items())
    mdx=f"WITH {definitions} SELECT {mdx_set(list(aliases))} ON COLUMNS FROM [{cube}] CELL PROPERTIES VALUE"
    root=checked_root(transport(profile,build_soap(profile,mdx),60))
    axis=root.find(f".//{{{MD_NS}}}Axis[@name='Axis0']")
    if axis is None:
        raise ValueError("Missing count axis")
    names=[]
    for item in axis.findall(f"./{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple"):
        members=item.findall(f"{{{MD_NS}}}Member")
        if len(members)!=1:
            raise ValueError("Invalid count tuple")
        names.append(members[0].findtext(f"{{{MD_NS}}}UName"))
    if len(names)!=len(aliases) or set(names)!=set(aliases):
        raise ValueError("Incomplete count axis")
    found={}
    for cell in root.findall(f".//{{{MD_NS}}}CellData/{{{MD_NS}}}Cell"):
        ordinal=integer(cell.get("CellOrdinal"))
        count=integer(cell.findtext(f"{{{MD_NS}}}Value"))
        if ordinal is None or ordinal<0 or ordinal>=len(names) or count is None or count<0 or ordinal in found:
            raise ValueError("Invalid count cell")
        found[ordinal]=count
    if len(found)!=len(names):
        raise ValueError("Incomplete count cells")
    return {aliases[name]:found[i] for i,name in enumerate(names)}


def discover_member_counts(transport,profile,hierarchies,emit=None):
    """Yield successful count groups immediately; shrink failed groups in place."""
    offset=0
    size=2
    previous_rate=None
    while offset<len(hierarchies):
        batch=hierarchies[offset:offset+size]
        emit_event(emit,"phase",phase="CARDINALITY",detail=f"Count {offset+1}-{offset+len(batch)}/{len(hierarchies)} | batch {len(batch)}")
        start=time.monotonic()
        try:
            counts=count_batch(transport,profile,batch)
        except (RuntimeError,ValueError) as exc:
            raise_if_fatal(exc)
            if len(batch)>1:
                size=max(1,len(batch)//2)
                previous_rate=None
                emit_event(emit,"discovery",message=f"Count batch reduced {len(batch)} -> {size} | {type(exc).__name__}")
                continue
            # Let the hierarchy remain unresolved; never infer zero from failure.
            yield {batch[0]:None}
            offset+=1
            continue
        elapsed=max(.001,time.monotonic()-start)
        rate=len(batch)/elapsed
        next_size=size
        if elapsed<30 and (previous_rate is None or rate>=previous_rate*.9):
            next_size=min(8,max(size+1,size*2))
        elif previous_rate is not None and rate<previous_rate*.7:
            next_size=max(1,size//2)
        emit_event(emit,"discovery",message=f"Count validated {offset+len(batch)}/{len(hierarchies)} | batch {len(batch)} | next {next_size} | {elapsed:.2f}s")
        yield counts
        offset+=len(batch)
        size=next_size
        previous_rate=rate


def exact_member_count(transport,profile,hierarchy,emit=None):
    emit_event(emit,"phase",phase="CARDINALITY",detail=f"Individual count | {hierarchy}")
    return count_batch(transport,profile,[hierarchy])[hierarchy]


def read_member_page(transport,profile,hierarchy,offset,size,emit=None):
    cube=profile["cube"].replace("]","]]")
    mdx=(f"SELECT Subset({hierarchy}.AllMembers,{offset},{size}) "
         f"DIMENSION PROPERTIES MEMBER_TYPE,PARENT_UNIQUE_NAME,CHILDREN_CARDINALITY,HIERARCHY_UNIQUE_NAME ON COLUMNS FROM [{cube}]")
    root=checked_root(transport(profile,build_soap(profile,mdx),90))
    axis=root.find(f".//{{{MD_NS}}}Axis[@name='Axis0']")
    if axis is None:
        raise ValueError("Missing member enumeration axis")
    page=axis.findall(f"./{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple")
    if len(page)!=size:
        raise ValueError(f"Truncated member enumeration: requested={size}, returned={len(page)}")
    result=[]
    for item in page:
        nodes=item.findall(f"{{{MD_NS}}}Member")
        if len(nodes)!=1:
            raise ValueError("Unexpected member tuple")
        node=nodes[0]
        props={local(e.tag):e.text for e in node}
        actual=props.get("HIERARCHY_UNIQUE_NAME") or node.get("Hierarchy")
        if actual and actual.casefold()!=hierarchy.casefold():
            raise ValueError("Unexpected member hierarchy: mismatched XMLA identity")
        if not actual:
            name=props.get("UName") or ""
            if not (name.casefold()==hierarchy.casefold() or name.casefold().startswith(hierarchy.casefold()+".")):
                raise ValueError("Unexpected member hierarchy: cannot establish XMLA identity")
        row={"MEMBER_UNIQUE_NAME":props.get("UName"),"MEMBER_CAPTION":props.get("Caption"),
             "LEVEL_UNIQUE_NAME":props.get("LName"),"LEVEL_NUMBER":props.get("LNum"),
             "PARENT_UNIQUE_NAME":props.get("PARENT_UNIQUE_NAME"),"MEMBER_TYPE":props.get("MEMBER_TYPE"),
             "CHILDREN_CARDINALITY":props.get("CHILDREN_CARDINALITY"),"HIERARCHY_UNIQUE_NAME":hierarchy}
        result.append(normalize_member(row))
    if len({m["unique_name"] for m in result})!=len(result):
        raise ValueError("Duplicate member enumeration page")
    return result


def inventory_checkpoint(conn,prep,cube,hierarchy):
    row=conn.execute("SELECT status,metadata_json FROM oic_dimensions WHERE prep=? AND cube=? AND hierarchy=?",
                     (prep,cube,hierarchy)).fetchone()
    if row is None:
        return {"status":"PENDING","saved":0,"expected":None,"error":None}
    status,raw=row
    state=json.loads(raw)
    saved=state.get("next_offset",state.get("discovered_cardinality"))
    if saved is None:
        saved=conn.execute("SELECT count(*) FROM oic_members WHERE prep=? AND cube=? AND hierarchy=?",
                           (prep,cube,hierarchy)).fetchone()[0]
    expected=state.get("expected_count")
    if expected is None and status in ("COMPLETE","EMPTY"):
        expected=saved
    return {"status":status,"saved":saved,"expected":expected,"error":state.get("error")}


def inventory_summary(conn,catalogs,prep=None):
    """Checkpoint-only snapshot; no materialization of the member table."""
    fingerprint=digest([PLANNER_VERSION,catalogs])
    if prep is None and conn is not None:
        row=conn.execute("SELECT id,status FROM oic_preparations WHERE catalog_hash=? ORDER BY created_at DESC LIMIT 1",
                         (fingerprint,)).fetchone()
        if row and row[1] in ("READY","PREPARING"):
            prep=row[0]
    cubes={}
    eligible={}
    for cube,(profile,catalog,fields) in catalogs.items():
        dims,_=classify_fields(fields)
        available=profile["availability"]=="AVAILABLE" and catalog.get("source")!="PROFILE"
        eligible[cube]=set(dims) if available else set()
        cubes[cube]={"available":available,"total":len(eligible[cube]),"hierarchies":{}}
    if conn is not None and prep:
        for cube,hierarchy in conn.execute("SELECT cube,hierarchy FROM oic_dimensions WHERE prep=?",(prep,)):
            if hierarchy in eligible.get(cube,()):
                cubes[cube]["hierarchies"][hierarchy]=inventory_checkpoint(conn,prep,cube,hierarchy)
    return {"prep":prep,"cubes":cubes}


def emit_inventory(db,prep,cube,hierarchy,emit):
    emit_event(emit,"inventory",prep=prep,cube=cube,hierarchy=hierarchy,
               checkpoint=inventory_checkpoint(db.conn,prep,cube,hierarchy))


def discover_members_mdx(db,prep,cube,transport,profile,hierarchy,count,emit=None):
    state=db.load_discovery_checkpoint(prep,cube,hierarchy)
    offset=state.get("next_offset",0)
    size=max(1,min(MAX_DISCOVERY_PAGE,state.get("page_size",DISCOVERY_PAGE)))
    last=state.get("last_member")
    # Resume validates the cached boundary against the live member ordering.
    # This is best effort: the provider does not expose a guaranteed immutable snapshot.
    if offset:
        emit_event(emit,"phase",phase="MEMBER_DISCOVERY",detail=f"Validating saved boundary | {hierarchy} | {offset:,}/{count:,}")
        boundary=read_member_page(transport,profile,hierarchy,offset-1,1,emit)
        if boundary[0]["unique_name"]!=last:
            db.reset_partial_inventory(prep,cube,hierarchy)
            emit_inventory(db,prep,cube,hierarchy,emit)
            offset,last,size=0,None,DISCOVERY_PAGE
            emit_event(emit,"activity",message=f"{cube} | source order changed; reset only {hierarchy}")
    pages=0
    previous_rate=None
    while offset<count:
        pages+=1
        if pages>MAX_DISCOVERY_PAGES:
            raise ValueError(f"Discovery page safety budget reached; resume retained at {offset:,}/{count:,}")
        take=min(size,count-offset)
        overlap=1 if offset else 0
        emit_event(emit,"phase",phase="MEMBER_DISCOVERY",detail=f"MDX page | {hierarchy} | offset {offset:,} | size {take}")
        emit_event(emit,"discovery",message=f"Members committed {offset:,}/{count:,} | page {take} | next offset {offset:,}")
        start=time.monotonic()
        try:
            page=read_member_page(transport,profile,hierarchy,offset-overlap,take+overlap,emit)
            if overlap and page[0]["unique_name"]!=last:
                raise ValueError("Member ordering changed during discovery")
            page=page[overlap:]
        except (RuntimeError,ValueError) as exc:
            raise_if_fatal(exc)
            reason=str(exc)[:240]
            emit_event(emit,"activity",message=f"{cube} | MDX PAGE FAILED | offset {offset:,} | requested {take+overlap} | {reason}")
            # Only scope/length failures can benefit from smaller requests.
            shrink=getattr(exc,"kind","") in ("QUERY_SCOPE_ERROR","DISCOVERY_SCOPE_UNCONFIRMED","RESPONSE_TOO_LARGE","RESPONSE_TIMEOUT") or reason.startswith("Truncated member enumeration")
            if not shrink or take<=1:
                raise
            size=max(1,take//2)
            previous_rate=None
            emit_event(emit,"discovery",message=f"Page reduced {take} -> {size} | unchanged offset {offset:,} | {reason}")
            continue
        elapsed=max(.001,time.monotonic()-start)
        rate=len(page)/elapsed
        next_size=size
        if elapsed<25 and (previous_rate is None or rate>=previous_rate*.9):
            next_size=min(MAX_DISCOVERY_PAGE,math.ceil(size*1.5))
        elif previous_rate is not None and rate<previous_rate*.7:
            next_size=max(1,size//2)
        db.append_member_page(prep,cube,hierarchy,offset,page,next_size)
        emit_inventory(db,prep,cube,hierarchy,emit)
        offset+=len(page)
        last=page[-1]["unique_name"]
        size=next_size
        previous_rate=rate
        emit_event(emit,"discovery",message=f"CHECKPOINT | {hierarchy} | {offset:,}/{count:,} committed | next size {size}")
    if offset and exact_member_count(transport,profile,hierarchy,emit)!=count:
        db.reset_partial_inventory(prep,cube,hierarchy)
        emit_inventory(db,prep,cube,hierarchy,emit)
        raise ValueError("Member cardinality changed during discovery")
    db.finalize_inventory(prep,cube,hierarchy,count)
    emit_inventory(db,prep,cube,hierarchy,emit)
    return db.inventory(prep,cube,hierarchy)


def select_discovery_strategy(count,levels):
    if count<=500:
        return "RESTRICTED"
    if count>2000 and levels and all(
        0<=integer(l.get("LEVEL_CARDINALITY"),-1)<=500 for l in levels):
        return "LEVEL"
    return "MDX_PAGINATION"


def discover_dimension_members(db,prep,cube,transport,profile,hierarchy,metadata,levels,count,emit=None):
    state=db.load_discovery_checkpoint(prep,cube,hierarchy)
    if count is None:
        raise ValueError("Exact cardinality not available")
    if count==0:
        db.finalize_inventory(prep,cube,hierarchy,0)
        emit_inventory(db,prep,cube,hierarchy,emit)
        return []
    strategy=select_discovery_strategy(count,levels)
    capability=db.load_profile(prep,cube,"DISCOVERY_METADATA:MEMBER_CAPABILITY") or {}
    if capability.get("prefer_mdx"):
        strategy="MDX_PAGINATION"
    if not state.get("next_offset") and strategy in ("RESTRICTED","LEVEL"):
        try:
            if strategy=="RESTRICTED":
                rows=discover_rows(transport,profile,"MDSCHEMA_MEMBERS",{"HIERARCHY_UNIQUE_NAME":hierarchy},emit=emit)
            else:
                rows=[]
                for level in levels:
                    name=level.get("LEVEL_UNIQUE_NAME")
                    if not name:
                        raise ValueError("Missing level unique name")
                    part=discover_rows(transport,profile,"MDSCHEMA_MEMBERS",
                                       {"HIERARCHY_UNIQUE_NAME":hierarchy,"LEVEL_UNIQUE_NAME":name},emit=emit)
                    if any(r.get("LEVEL_UNIQUE_NAME")!=name for r in part):
                        raise ValueError("Level restriction not honored")
                    rows.extend(part)
            members=deduplicate_members(rows,hierarchy)
            if len(members)!=count:
                raise ValueError(f"Restricted member count mismatch: expected={count:,}, unique={len(members):,}, rows={len(rows):,}")
            names={m["unique_name"] for m in members}
            for m in members:
                m["parent_present"]=m["parent"] in names if m["parent"] else None
            if not levels:
                levels=list({m["level"]:{"LEVEL_UNIQUE_NAME":m["level"],"LEVEL_NUMBER":m["level_number"],
                    "source":"MEMBER_METADATA"} for m in members if m["level"]}.values())
            db.save_inventory(prep,cube,hierarchy,"COMPLETE",
                {"metadata":metadata,"reported_cardinality":integer(metadata.get("HIERARCHY_CARDINALITY")),
                 "discovered_cardinality":count,"expected_count":count,"next_offset":count,
                 "source":"XMLA_COUNT_VALIDATED","method":strategy},levels,members)
            emit_inventory(db,prep,cube,hierarchy,emit)
            emit_event(emit,"discovery",message=f"CHECKPOINT | {hierarchy} | {count:,}/{count:,} committed")
            return members
        except (RuntimeError,ValueError) as exc:
            raise_if_fatal(exc)
            kind=getattr(exc,"kind","")
            if kind in ("DISCOVER_NOT_SUPPORTED","DISCOVERY_SCOPE_UNCONFIRMED","QUERY_SCOPE_ERROR") or "restriction not honored" in str(exc):
                failures=capability.get("failures",0)+1
                prefer=kind=="DISCOVER_NOT_SUPPORTED" or "restriction not honored" in str(exc) or failures>=2
                db.save_profile(prep,cube,"DISCOVERY_METADATA:MEMBER_CAPABILITY",
                                {"failures":failures,"prefer_mdx":prefer,"reason":kind or "RESTRICTION_NOT_HONORED"})
                if prefer:
                    emit_event(emit,"activity",message=f"{cube} | restricted Discover unreliable; prefer paged MDX")
            emit_event(emit,"activity",message=f"{hierarchy} | {strategy} failed | {str(exc)[:240]} | fallback to paged MDX")
    return discover_members_mdx(db,prep,cube,transport,profile,hierarchy,count,emit)


def workload_group(n):
    return next((label for limit,label in ((10,"XS"),(50,"S"),(200,"M"),(1000,"L"),(5000,"XL")) if n<=limit),"XXL")


def initial_profile(n,y,m):
    ysize=min(y,8)
    msize=min(m,40)
    xsize=min(max(1,n),100,max(1,MAX_CELLS//max(1,ysize*msize)))
    return {"x":xsize,"y":ysize,"m":msize,"timeout":DEFAULT_TIMEOUT,
            "source":"CONSERVATIVE","confidence":"ESTIMATED","samples":[]}


def partition(values, size):
    return [values[i:i+size] for i in range(0,len(values),size)]


def potential(scope):
    return len(scope["members"])*len(scope["years"])*len(scope["measures"])


def request_record(prep,job,cube,profile,scope,timeout,status="PENDING",parent=None):
    mdx=build_mdx(profile,scope)
    payload=build_soap(profile,mdx)
    identity={k:scope[k] for k in ("hierarchy","year_hierarchy","members","years","measures")}
    rid=digest([prep,cube,job,identity,parent,"PROBE" if status=="PROBE" else "EXTRACT",scope.get("probe_kind")])
    return {"id":rid,"prep":prep,"job":job,"cube":cube,"scope":scope,"mdx":mdx,
            "payload":payload,"timeout":timeout,"potential_cells":potential(scope),
            "status":status,"parent":parent}


def subtract_scope(scope,covered):
    # Exact disjoint rectangular subtraction without materializing cells.
    keys=("members","years","measures")
    intersection={k:[v for v in scope[k] if v in set(covered[k])] for k in keys}
    if any(not intersection[k] for k in keys):
        return [scope]
    remainder=dict(scope)
    result=[]
    for k in keys:
        inside=set(intersection[k])
        outside=[v for v in remainder[k] if v not in inside]
        if outside:
            result.append(dict(remainder,**{k:outside}))
        remainder=dict(remainder,**{k:intersection[k]})
    return result


def benchmark_profile(db,prep,job,cube,profile,scope,base,probe,budget,emit=None):
    samples=[]
    if not scope["hierarchy"]:
        return base,budget
    # Compare bounded X/Y/M shapes. Non-overlapping samples are reusable data.
    shapes=[
        (base["x"],base["y"],base["m"]),
        (base["x"]*2,base["y"],base["m"]),
        (base["x"],max(1,base["y"]//2),min(len(scope["measures"]),base["m"]*2)),
        (base["x"]*2,base["y"],max(1,base["m"]//2)),
    ]
    occupied=[]
    for x,y,m in shapes:
        if budget<=0:
            break
        x=max(1,min(x,MAX_CELLS//max(1,y*m)))
        candidates=[scope]
        for previous in occupied:
            candidates=[r for c in candidates for r in subtract_scope(c,previous)]
        if not candidates:
            break
        region=max(candidates,key=potential)
        part=dict(region,members=region["members"][:x],years=region["years"][:y],measures=region["measures"][:m])
        req=request_record(prep,job,cube,profile,part,base["timeout"],"PROBE")
        db.insert_requests([req])
        before=time.monotonic()
        result=probe(req,profile)
        budget-=1
        elapsed=max(0.001,float(result.get("duration",time.monotonic()-before)))
        samples.append({"x":len(part["members"]),"y":len(part["years"]),"m":len(part["measures"]),
                        "seconds":elapsed,"ok":bool(result.get("ok")),
                        "covered_per_second":potential(part)/elapsed if result.get("ok") else 0,
                        "useful_cells":result.get("useful_cells",0)})
        emit_event(emit,"sample",cube=cube,hierarchy=scope["hierarchy"],sample=samples[-1])
        occupied.append(part)
        if result.get("pressure"):
            break
    good=[v for v in samples if v["ok"]]
    if not good:
        return dict(base,samples=samples),budget
    best=max(good,key=lambda v:v["covered_per_second"])
    return dict(base,x=best["x"],y=best["y"],m=best["m"],
                timeout=min(300,max(60,math.ceil(max(v["seconds"] for v in good)*1.5),DEFAULT_TIMEOUT)),
                source="DIRECT_SAMPLE",confidence="LOW",samples=samples),budget


def optimize_concurrency(db,prep,jobs,probe,emit=None):
    emit_event(emit,"phase",phase="CONCURRENCY",detail="Selecting concurrency profile")
    saved=db.load_profile(prep,"GLOBAL","concurrency")
    if saved:
        emit_event(emit,"activity",message="Reusing saved concurrency profile")
        return saved
    # Compare 1, 2, and at most 4 simultaneous requests using disjoint,
    # equal-sized batches of one representative job. Do not benchmark live
    # traffic again after the bounded seven-probe budget.
    used=db.conn.execute("""SELECT count(*) FROM oic_requests
        WHERE prep=? AND json_extract(scope_json,'$.probe_kind')='concurrency'""",(prep,)).fetchone()[0]
    if used:
        result={"limit":1,"source":"CONSERVATIVE_INTERRUPTED_BENCHMARK","samples":[]}
        db.save_profile(prep,"GLOBAL","concurrency",result)
        return result
    selected=None
    for job,cube,profile,scope in jobs:
        if not scope["hierarchy"] or len(scope["members"])<140:
            continue
        remaining=[scope]
        for row in db.conn.execute("SELECT scope_json FROM oic_requests WHERE prep=? AND job=? AND status IN ('COMMITTED','EMPTY')",(prep,job)):
            remaining=[r for c in remaining for r in subtract_scope(c,json.loads(row[0]))]
        for region in remaining:
            if len(region["members"])>=140:
                selected=(job,cube,profile,region)
                break
        if selected:
            break
    if not selected:
        result={"limit":1,"source":"CONSERVATIVE_NO_COMPARABLE_SAMPLE","samples":[]}
        db.save_profile(prep,"GLOBAL","concurrency",result)
        return result
    job,cube,profile,region=selected
    offset=0
    samples=[]
    for concurrency in (1,2,4):
        emit_event(emit,"phase",phase="CONCURRENCY",cube=cube,detail=f"Benchmark {concurrency} simultaneous requests")
        reqs=[]
        for _ in range(concurrency):
            part=dict(region,members=region["members"][offset:offset+20],
                      years=region["years"][:4],measures=region["measures"][:20],probe_kind="concurrency")
            offset+=20
            reqs.append(request_record(prep,job,cube,profile,part,DEFAULT_TIMEOUT,"PROBE"))
        db.insert_requests(reqs)
        start=time.monotonic()
        with ThreadPoolExecutor(max_workers=concurrency) as pool:
            results=list(pool.map(lambda r:probe(r,profile,parallel=True),reqs))
        elapsed=max(0.001,time.monotonic()-start)
        success=all(r.get("ok") for r in results)
        samples.append({"concurrency":concurrency,"seconds":elapsed,"ok":success,
                        "covered_per_second":sum(r["potential_cells"] for r in reqs)/elapsed if success else 0})
        emit_event(emit,"concurrency_sample",sample=samples[-1])
        if not success or any(r.get("pressure") for r in results):
            break
        if len(samples)>1 and samples[-1]["covered_per_second"]<=samples[-2]["covered_per_second"]*1.05:
            break
    good=[v for v in samples if v["ok"]]
    best=max(good,key=lambda v:v["covered_per_second"]) if good else {"concurrency":1}
    result={"limit":best["concurrency"],"source":"BOUNDED_COMPARATIVE_SAMPLE","samples":samples}
    db.save_profile(prep,"GLOBAL","concurrency",result)
    return result


def reprepare_failed_requests(db,prep,log):
    rows=[dict(r) for r in db.conn.execute("SELECT * FROM oic_requests WHERE prep=? AND status='NEEDS_REPREPARE'",(prep,))]
    for row in rows:
        scope=json.loads(row["scope_json"])
        axis=max(("members","years","measures"),key=lambda k:len(scope[k]))
        if len(scope[axis])<2:
            log(f"{row['id'][:12]} remains unresolved: minimum request scope")
            continue
        values=scope[axis]
        mid=len(values)//2
        profile=db.profile(prep,row["cube"])
        children=[request_record(prep,row["job"],row["cube"],profile,dict(scope,**{axis:part}),
                                  row["timeout"],"PENDING",row["id"]) for part in (values[:mid],values[mid:])]
        db.reprepare_request(row,children)
        log(f"REPREPARED {row['id'][:12]} | two disjoint requests")
    db.retry_failed(prep)


def build_recovery_requests(req,profile):
    scope=req["scope"]
    if req["potential_cells"]<1000:
        return []
    axis=max(("members","years","measures"),key=lambda k:len(scope[k]))
    values=scope[axis]
    if len(values)<2:
        return []
    middle=len(values)//2
    return [request_record(req["prep"],req["job"],req["cube"],profile,dict(scope,**{axis:part}),
                           req["timeout"],"STANDBY",req["id"]) for part in (values[:middle],values[middle:])]


def generate_request_plan(db,prep,job,cube,profile,scope,adaptive,emit=None):
    probes=[json.loads(r[0]) for r in db.conn.execute(
        "SELECT scope_json FROM oic_requests WHERE prep=? AND job=? AND status IN ('COMMITTED','EMPTY') AND parent IS NULL",
        (prep,job))]
    # Generation runs only during PREPARING. Previously generated primary
    # requests may exist after interruption; their IDs are deterministic.
    n=0
    recovery=0
    total=0
    batch=[]
    for x,y,m in itertools.product(partition(scope["members"],adaptive["x"]),
                                    partition(scope["years"],adaptive["y"]),
                                    partition(scope["measures"],adaptive["m"])):
        part=dict(scope,members=x,years=y,measures=m,profile_id=adaptive.get("profile_id"))
        remaining=[part]
        for covered in probes:
            remaining=[r for piece in remaining for r in subtract_scope(piece,covered)]
        for piece in remaining:
            req=request_record(prep,job,cube,profile,piece,adaptive["timeout"])
            batch.append(req)
            children=build_recovery_requests(req,profile)
            batch.extend(children)
            recovery+=len(children)
            n+=1
            total+=req["potential_cells"]
            if len(batch)>=100:
                db.insert_requests(batch)
                emit_event(emit,"plan_batch",primary=n,recovery=recovery)
                batch=[]
    if batch:
        db.insert_requests(batch)
    emit_event(emit,"plan_batch",primary=n,recovery=recovery)
    # Probe scopes are constructed disjointly; remaining rectangles are
    # generated by exact subtraction from disjoint Cartesian partitions.
    if total+sum(potential(p) for p in probes)!=potential(scope):
        raise RuntimeError("Prepared request coverage mismatch")
    return n


def prepare_all(db,transport,probe,output,log,emit=None,catalogs=None):
    if catalogs is None:
        catalogs=load_all_catalogs(emit=emit)
    else:
        # The caller owns a stable local catalog snapshot. Resolve mutable
        # year/profile fields on a private copy, never on the UI's cache.
        catalogs=deepcopy(catalogs)
        for cube,(profile,catalog,fields) in catalogs.items():
            emit_event(emit,"catalog",cube=cube,source=catalog.get("source") or "UNKNOWN",
                       fields=len(fields),availability=profile["availability"])
        log("CATALOG | Existing in-memory snapshot reused")
    emit_event(emit,"phase",phase="CATALOG",detail="Checking preparation snapshot")
    fingerprint=digest([PLANNER_VERSION,catalogs])
    current=db.preparation(fingerprint)
    prep=current["id"]
    version_changed=db.conn.execute("SELECT 1 FROM oic_requests WHERE prep=? AND error LIKE 'SOURCE_VERSION_CHANGED%' LIMIT 1",(prep,)).fetchone()
    if version_changed:
        db.conn.execute("UPDATE oic_preparations SET status='STALE' WHERE id=?",(prep,))
        current=db.preparation(fingerprint)
        prep=current["id"]
        log("Source update marker changed; preparing a new snapshot")
    elif current["status"]=="READY":
        emit_event(emit,"inventory_snapshot",snapshot=inventory_summary(db.conn,catalogs,prep))
        emit_event(emit,"phase",phase="REUSING",detail="Reusing prepared inventory and recovering failed requests")
        reprepare_failed_requests(db,prep,log)
        log("Existing READY inventory reused; failed scopes reprepared when possible")
        summary=json.loads(current["summary_json"])
        summary.update(requests=db.progress(prep),coverage=db.coverage(prep))
        db.finalize_preparation(prep,"READY",summary,current["concurrency"])
        (output/"preparation_summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
        emit_event(emit,"summary",summary=summary)
        return not db.conn.execute("SELECT 1 FROM oic_requests WHERE prep=? AND status='NEEDS_REPREPARE' LIMIT 1",(prep,)).fetchone()
    db.recover_preparation(prep)
    emit_event(emit,"inventory_snapshot",snapshot=inventory_summary(db.conn,catalogs,prep))
    errors=[]
    jobs=[]
    total_members=0
    total_fields=0
    unresolved=0
    cube_summary={}
    unavailable=[]
    for cube,(profile,catalog,fields) in catalogs.items():
        log(f"PREPARE {cube} | {len(fields)} catalog fields")
        dimensions,measures=classify_fields(fields)
        emit_event(emit,"cube",cube=cube,status="DISCOVERING",total=len(dimensions),done=0,members=0,
                   phase="METADATA_DISCOVERY",detail="Checking saved inventory")
        total_fields+=len(fields)
        unresolved+=sum(1 for f in fields if f.get("mdx") not in dimensions and f.get("mdx") not in measures)
        if profile["availability"]!="AVAILABLE" or catalog.get("source")=="PROFILE":
            db.save_catalog(prep,cube,profile,fields,measures)
            if profile["availability"]=="NOT_AVAILABLE":
                unavailable.append(cube)
                emit_event(emit,"cube",cube=cube,status="UNAVAILABLE")
            else:
                errors.append(f"{cube}: missing executable catalog")
                emit_event(emit,"cube",cube=cube,status="INCOMPLETE")
            continue
        try:
            profile["year"]=resolve_year_hierarchy(profile,dimensions)
        except ValueError as exc:
            db.save_catalog(prep,cube,profile,fields,measures)
            errors.append(f"{cube}: {exc}")
            emit_event(emit,"cube",cube=cube,status="INCOMPLETE")
            continue
        if not db.conn.execute("SELECT 1 FROM oic_cubes WHERE prep=? AND cube=?",(prep,cube)).fetchone():
            db.save_catalog(prep,cube,profile,fields,measures)
        inventory={}
        for h in dimensions:
            cached=db.inventory(prep,cube,h)
            if cached is not None:
                inventory[h]=cached
        pending=[h for h in dimensions if h not in inventory]
        cube_members=sum(len(v) for v in inventory.values())
        total_members+=cube_members
        emit_event(emit,"cube",cube=cube,done=len(inventory),members=cube_members)
        emit_event(emit,"activity",message=f"{cube} | {len(inventory)} cached hierarchies | {len(pending)} pending")
        hierarchies,levels,live_measures=[],[],[]
        if pending:
            targets=(("MDSCHEMA_HIERARCHIES",hierarchies),("MDSCHEMA_LEVELS",levels),("MDSCHEMA_MEASURES",live_measures))
            for kind,target in targets:
                cached=db.load_profile(prep,cube,"DISCOVERY_METADATA:"+kind)
                if cached is not None:
                    target.extend(cached["rows"])
                    emit_event(emit,"activity",message=f"{cube} | reusing {kind}")
                    continue
                try:
                    target.extend(discover_rows(transport,profile,kind,emit=emit))
                    db.save_profile(prep,cube,"DISCOVERY_METADATA:"+kind,{"rows":target})
                except (ValueError,RuntimeError) as exc:
                    raise_if_fatal(exc)
                    log(f"{cube} {kind} fallback: {type(exc).__name__}")
            native={r.get("MEASURE_UNIQUE_NAME"):r for r in live_measures}
            measures={k:dict(v,live_metadata=native.get(k)) for k,v in measures.items()}
            db.save_catalog(prep,cube,profile,fields,measures)
        else:
            stored=db.conn.execute("SELECT measure,metadata_json FROM oic_measures WHERE prep=? AND cube=?",(prep,cube)).fetchall()
            measures={r[0]:json.loads(r[1]) for r in stored}
        hierarchy_meta={str(r.get("HIERARCHY_UNIQUE_NAME","")).casefold():r for r in hierarchies}
        levels_by_h={}
        for row in levels:
            levels_by_h.setdefault(str(row.get("HIERARCHY_UNIQUE_NAME","")).casefold(),[]).append(row)
        pending.sort(key=lambda h:(h!=profile["year"],
            integer(hierarchy_meta.get(h.casefold(),{}).get("HIERARCHY_CARDINALITY"),10**12)))
        # Each yielded count group is persisted and discovered before fetching the next.
        # Revalidate counts for partial pages on resume, rather than trusting a stale count.
        for counts in discover_member_counts(transport,profile,pending,emit):
            for h,count in counts.items():
                meta=hierarchy_meta.get(h.casefold(),{})
                if count is not None and count>MAX_DISCOVERY_PAGES*MAX_DISCOVERY_PAGE:
                    reported=integer(meta.get("HIERARCHY_CARDINALITY"))
                    try:
                        confirmed=exact_member_count(transport,profile,h,emit)
                    except (RuntimeError,ValueError) as exc:
                        raise_if_fatal(exc)
                        confirmed=None
                        log(f"{cube} | {h} | individual cardinality failed: {str(exc)[:200]}")
                    log(f"{cube} | LARGE CARDINALITY | {h} | batch={count:,} | individual={confirmed} | metadata={reported} | page budget may require resume")
                    if confirmed!=count:
                        counts[h]=count=None
                        log(f"{cube} | {h} | unconfirmed cardinality; retaining previous checkpoint")
                db.save_cardinality(prep,cube,h,meta,levels_by_h.get(h.casefold(),[]),count)
                emit_inventory(db,prep,cube,h,emit)
            for h,count in counts.items():
                emit_event(emit,"phase",phase="MEMBER_DISCOVERY",detail=f"Hierarchy | {h}")
                try:
                    members=discover_dimension_members(db,prep,cube,transport,profile,h,
                        hierarchy_meta.get(h.casefold(),{}),levels_by_h.get(h.casefold(),[]),count,emit)
                    log(f"{cube} {h} | MEMBERS {len(members)}")
                except (ValueError,RuntimeError) as exc:
                    if getattr(exc,"kind","")=="USER_STOP":
                        raise
                    db.mark_discovery_failure(prep,cube,h,str(exc)[:300])
                    emit_inventory(db,prep,cube,h,emit)
                    raise_if_fatal(exc)
                    errors.append(f"{cube} {h}: {str(exc)[:200]}")
                    emit_event(emit,"activity",message=f"{cube} | hierarchy incomplete | {h} | {str(exc)[:240]}")
                    continue
                inventory[h]=members
                total_members+=len(members)
                cube_members+=len(members)
                emit_event(emit,"cube",cube=cube,done=len(inventory),members=cube_members)
        # Year-only totals and one hierarchy at a time; never crossjoin year with itself.
        years=[m["unique_name"] for m in inventory.get(profile["year"],[])
               if not m["is_all_member"] and not m["is_calculated_member"]
               and (m["member_type"] is not None or (m["level_number"] or 0)>0)]
        cube_summary[cube]={"fields":len(fields),"dimensions":len(dimensions),"inventories_complete":len(inventory),
                            "members":sum(len(m) for m in inventory.values()),"years":len(years),"native_measures":len(measures)}
        if not years or not measures:
            errors.append(f"{cube}: no validated years or native measures")
            emit_event(emit,"cube",cube=cube,status="INCOMPLETE")
            continue
        for h in [""]+[h for h in dimensions if h!=profile["year"]]:
            members=[m["unique_name"] for m in inventory.get(h,[])] if h else [""]
            if not members:
                continue
            groups={}
            for name,meta in measures.items():
                group=(meta.get("live_metadata") or {}).get("MEASUREGROUP_NAME") or "UNKNOWN"
                groups.setdefault(group,[]).append(name)
            for group,names in sorted(groups.items()):
                scope={"hierarchy":h,"year_hierarchy":profile["year"],"members":members,
                       "years":years,"measures":sorted(names),"measure_group":group}
                job=digest([prep,cube,h,scope])
                db.save_job(prep,job,cube,h,scope)
                jobs.append((job,cube,profile,scope))
        emit_event(emit,"cube",cube=cube,status="INVENTORY READY" if len(inventory)==len(dimensions) else "INCOMPLETE")
    summary={"preparation":prep,"catalog_fields":total_fields,"discovered_members":total_members,
             "unresolved_fields":unresolved,"jobs":len(jobs),"errors":errors,
             "cube_details":cube_summary,"unavailable_cubes":unavailable,"available_cubes":sum(p[0]["availability"]=="AVAILABLE" for p in catalogs.values())}
    if errors:
        db.finalize_preparation(prep,"PREPARING",summary)
        (output/"preparation_summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
        emit_event(emit,"summary",summary=summary)
        log(f"PREPARE INCOMPLETE | {len(errors)} errors; rerun --prepare")
        return False
    # Spend at most four shape probes per cube, keeping the global budget
    # distributed across all categories instead of exhausting it on the first.
    profiles={}
    for index,(job,cube,profile,scope) in enumerate(sorted(jobs,key=lambda j:len(j[3]["members"]),reverse=True),1):
        emit_event(emit,"phase",phase="ADAPTIVE",cube=cube,detail=f"Job {index}/{len(jobs)} | {scope['hierarchy'] or 'YEAR ONLY'}")
        key=("YEAR_ONLY" if not scope["hierarchy"] else workload_group(len(scope["members"])))+":"+str(len(scope["years"]))+":"+digest(scope["measures"])[:12]
        adaptive=db.load_profile(prep,cube,key)
        if adaptive is None:
            base=initial_profile(len(scope["members"]),len(scope["years"]),len(scope["measures"]))
            used=db.conn.execute("""SELECT count(*) FROM oic_requests WHERE prep=? AND cube=?
                 AND status IN ('PROBE','PROBE_FAILED','COMMITTED','EMPTY') AND json_extract(scope_json,'$.probe_kind') IS NULL""",(prep,cube)).fetchone()[0]
            budget=max(0,min(4-used,PROBE_BUDGET))
            adaptive,_=benchmark_profile(db,prep,job,cube,profile,scope,base,probe,budget,emit=emit)
            db.save_profile(prep,cube,key,adaptive)
        emit_event(emit,"profile",cube=cube,key=key,profile=adaptive)
        profiles[job]=dict(adaptive,profile_id=key)
    concurrency=optimize_concurrency(db,prep,jobs,probe,emit=emit)
    emit_event(emit,"activity",message=f"Selected concurrency: {concurrency['limit']} | {concurrency['source']}")
    for index,(job,cube,profile,scope) in enumerate(jobs,1):
        emit_event(emit,"plan_job",cube=cube,index=index,total=len(jobs),detail=scope["hierarchy"] or "YEAR ONLY")
        generate_request_plan(db,prep,job,cube,profile,scope,profiles[job],emit=emit)
        emit_event(emit,"plan_done")
    emit_event(emit,"phase",phase="VALIDATING",detail="Finalizing coverage and preparation summary")
    db.conn.execute("UPDATE oic_requests SET status='PROBE_FAILED' WHERE prep=? AND status='PROBE'",(prep,))
    summary["requests"]=db.progress(prep)
    summary["coverage"]=db.coverage(prep)
    summary["probe_budget"]=PROBE_BUDGET
    summary["profile_counts"]={}
    for row in db.conn.execute("SELECT metadata_json FROM oic_profiles WHERE prep=? AND cube!='GLOBAL' AND profile NOT LIKE 'DISCOVERY_METADATA:%'",(prep,)):
        source=json.loads(row[0]).get("source","UNKNOWN")
        summary["profile_counts"][source]=summary["profile_counts"].get(source,0)+1
    summary["estimated_extraction_seconds"]=None
    # Only extrapolate when all pending jobs have measured throughput.
    elapsed=0.0
    for job,_,_,scope in jobs:
        good=[sample["covered_per_second"] for sample in profiles[job].get("samples",[]) if sample.get("ok") and sample.get("covered_per_second",0)>0]
        remaining=db.conn.execute("SELECT coalesce(sum(potential_cells),0) FROM oic_requests WHERE prep=? AND job=? AND status='PENDING'",(prep,job)).fetchone()[0]
        if remaining and not good:
            break
        if remaining:
            elapsed+=remaining/max(good)
    else:
        summary["estimated_extraction_seconds"]=round(elapsed/max(1,concurrency["limit"]))
    summary["recommended_concurrency"]=concurrency["limit"]
    summary["concurrency_profile"]=concurrency
    db.finalize_preparation(prep,"READY",summary,concurrency=concurrency["limit"])
    (output/"preparation_summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
    emit_event(emit,"summary",summary=summary)
    log(f"PREPARE READY | {len(jobs)} jobs | {summary['requests']}")
    return True
