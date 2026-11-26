"""Prepare immutable, executable XMLA requests before starting any worker."""
from __future__ import annotations

import itertools
import json
import math
import re
import time
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
DISCOVERY_PAGE = 500
MAX_DISCOVERY_PAGES = 10000
DEFAULT_TIMEOUT = 180


def local(tag):
    return tag.rsplit("}", 1)[-1]


def integer(value, default=None):
    try:
        return int(value)
    except (TypeError, ValueError):
        return default


def checked_root(data):
    root = ET.fromstring(data)
    if any(local(e.tag) in ("Fault", "Error") for e in root.iter()):
        raise ValueError("XMLA fault/error")
    return root


def load_all_catalogs():
    result = {}
    for code, kind in itertools.product(CATEGORIES, YEAR_TYPES):
        key = f"{code}/{kind}"
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
            original = {str(f.get("id")): f for f in report.get("fields", []) if f.get("id")}
            fields = [dict(f, original_report=original.get(str(f.get("id")))) for f in fields]
        result[key] = (profile, catalog, fields)
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


def discover_rows(transport,profile,kind,restrictions=None):
    return parse_discover_rows(transport(profile,build_discover_request(profile,kind,restrictions),DEFAULT_TIMEOUT))


def mdx_set(values):
    # Accept only bracketed unique names returned by the server, not captions.
    pattern=r"(?:\[(?:[^\]\r\n]|\]\])+\])(?:\.(?:&?\[(?:[^\]\r\n]|\]\])+\])(?:&\[(?:[^\]\r\n]|\]\])+\])*)*"
    if not values or any(not re.fullmatch(pattern,v) for v in values):
        raise ValueError("Invalid or empty MDX unique-name set")
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
        raise ValueError("Member without unique name")
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
            continue
        member=normalize_member(row)
        result[member["unique_name"]]=member
    return sorted(result.values(),key=lambda m:m["unique_name"])


def exact_member_count(transport, profile, hierarchy):
    if not query.IDENTIFIER.fullmatch(hierarchy):
        raise ValueError("Invalid hierarchy")
    cube=profile["cube"].replace("]","]]")
    mdx=(f"WITH MEMBER [Measures].[OICInventoryCount] AS Count({hierarchy}.AllMembers) "
         f"SELECT {{[Measures].[OICInventoryCount]}} ON COLUMNS FROM [{cube}]")
    root=checked_root(transport(profile,build_soap(profile,mdx),DEFAULT_TIMEOUT))
    values=root.findall(f".//{{{MD_NS}}}CellData/{{{MD_NS}}}Cell/{{{MD_NS}}}Value")
    if len(values)!=1:
        raise ValueError("Missing inventory count")
    count=integer(values[0].text)
    if count is None or count<0:
        raise ValueError("Invalid inventory count")
    return count


def discover_members_mdx(transport,profile,hierarchy,count):
    result={}
    cube=profile["cube"].replace("]","]]")
    if math.ceil(count/DISCOVERY_PAGE)>MAX_DISCOVERY_PAGES:
        raise ValueError("Inventory exceeds discovery safety budget")
    for offset in range(0,count,DISCOVERY_PAGE):
        size=min(DISCOVERY_PAGE,count-offset)
        mdx=(f"SELECT Subset({hierarchy}.AllMembers,{offset},{size}) "
             f"DIMENSION PROPERTIES MEMBER_TYPE,PARENT_UNIQUE_NAME,CHILDREN_CARDINALITY ON COLUMNS FROM [{cube}]")
        root=checked_root(transport(profile,build_soap(profile,mdx),DEFAULT_TIMEOUT))
        axis=root.find(f".//{{{MD_NS}}}Axis[@name='Axis0']")
        if axis is None:
            raise ValueError("Missing member enumeration axis")
        page=axis.findall(f"./{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple")
        if len(page)!=size:
            raise ValueError("Truncated member enumeration")
        for item in page:
            nodes=item.findall(f"{{{MD_NS}}}Member")
            if len(nodes)!=1:
                raise ValueError("Unexpected member tuple")
            node=nodes[0]
            props={local(e.tag):e.text for e in node}
            row={"MEMBER_UNIQUE_NAME":props.get("UName"),"MEMBER_CAPTION":props.get("Caption"),
                 "LEVEL_UNIQUE_NAME":props.get("LName"),"LEVEL_NUMBER":props.get("LNum"),
                 "PARENT_UNIQUE_NAME":props.get("PARENT_UNIQUE_NAME"),"MEMBER_TYPE":props.get("MEMBER_TYPE"),
                 "CHILDREN_CARDINALITY":props.get("CHILDREN_CARDINALITY"),"HIERARCHY_UNIQUE_NAME":hierarchy}
            member=normalize_member(row)
            if member["unique_name"] in result:
                raise ValueError("Duplicate member enumeration page")
            result[member["unique_name"]]=member
    if len(result)!=count:
        raise ValueError("Member count mismatch")
    return sorted(result.values(),key=lambda m:m["unique_name"])


def discover_dimension_members(transport,profile,hierarchy,bulk,metadata,levels):
    members=deduplicate_members(bulk,hierarchy)
    reported=integer(metadata.get("HIERARCHY_CARDINALITY"))
    # Verify completeness against an actual MDX count rather than trusting
    # schema cardinality estimates or a short/empty response.
    count=exact_member_count(transport,profile,hierarchy)
    if len(members)!=count:
        try:
            rows=discover_rows(transport,profile,"MDSCHEMA_MEMBERS",{"HIERARCHY_UNIQUE_NAME":hierarchy})
            members=deduplicate_members(rows,hierarchy)
        except (ValueError, RuntimeError):
            members=[]
    if len(members)!=count and levels:
        try:
            rows=[]
            for level in levels:
                if level.get("LEVEL_UNIQUE_NAME"):
                    rows.extend(discover_rows(transport,profile,"MDSCHEMA_MEMBERS",
                        {"HIERARCHY_UNIQUE_NAME":hierarchy,"LEVEL_UNIQUE_NAME":level["LEVEL_UNIQUE_NAME"]}))
            members=deduplicate_members(rows,hierarchy)
        except (ValueError,RuntimeError):
            members=[]
    if len(members)!=count:
        members=discover_members_mdx(transport,profile,hierarchy,count)
    names={m["unique_name"] for m in members}
    for member in members:
        member["parent_present"]=member["parent"] in names if member["parent"] else None
    return members,{"reported_cardinality":reported,"discovered_cardinality":count,
                    "metadata":metadata,"source":"XMLA_COUNT_VALIDATED"}


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
    rid=digest([prep,cube,job,scope,parent])
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


def benchmark_profile(db,prep,job,cube,profile,scope,base,probe,budget):
    samples=[]
    if not scope["hierarchy"]:
        return base,budget
    # Disjoint member samples allow successful probe data to remain committed.
    # Compare within a workload group; estimates are deliberately low confidence.
    offset=0
    for factor in (0.5,1,2):
        if budget<=0:
            break
        x=max(1,int(base["x"]*factor))
        x=min(x,MAX_CELLS//max(1,base["y"]*base["m"]))
        members=scope["members"][offset:offset+x]
        if not members:
            break
        part=dict(scope,members=members,years=scope["years"][:base["y"]],measures=scope["measures"][:base["m"]])
        req=request_record(prep,job,cube,profile,part,base["timeout"],"PROBE")
        db.insert_requests([req])
        before=time.monotonic()
        result=probe(req,profile)
        budget-=1
        elapsed=max(0.001,float(result.get("duration",time.monotonic()-before)))
        samples.append({"x":len(members),"y":len(part["years"]),"m":len(part["measures"]),
                        "seconds":elapsed,"ok":bool(result.get("ok")),
                        "covered_per_second":potential(part)/elapsed if result.get("ok") else 0,
                        "useful_cells":result.get("useful_cells",0)})
        offset+=len(members)
        if result.get("pressure"):
            break
    good=[s for s in samples if s["ok"]]
    if not good:
        return dict(base,samples=samples),budget
    best=max(good,key=lambda s:s["covered_per_second"])
    # Few samples cannot establish a statistically meaningful P95.
    return dict(base,x=best["x"],y=best["y"],m=best["m"],
                timeout=min(300,max(60,math.ceil(max(s["seconds"] for s in good)*1.5),DEFAULT_TIMEOUT)),
                source="DIRECT_SAMPLE",confidence="LOW",samples=samples),budget


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


def generate_request_plan(db,prep,job,cube,profile,scope,adaptive):
    probes=[json.loads(r[0]) for r in db.conn.execute(
        "SELECT scope_json FROM oic_requests WHERE prep=? AND job=? AND status IN ('COMMITTED','EMPTY') AND parent IS NULL",
        (prep,job))]
    # Generation runs only during PREPARING. Previously generated primary
    # requests may exist after interruption; their IDs are deterministic.
    n=0
    total=0
    batch=[]
    for x,y,m in itertools.product(partition(scope["members"],adaptive["x"]),
                                    partition(scope["years"],adaptive["y"]),
                                    partition(scope["measures"],adaptive["m"])):
        part=dict(scope,members=x,years=y,measures=m)
        remaining=[part]
        for covered in probes:
            remaining=[r for piece in remaining for r in subtract_scope(piece,covered)]
        for piece in remaining:
            req=request_record(prep,job,cube,profile,piece,adaptive["timeout"])
            batch.append(req)
            batch.extend(build_recovery_requests(req,profile))
            n+=1
            total+=req["potential_cells"]
            if len(batch)>=100:
                db.insert_requests(batch)
                batch=[]
    if batch:
        db.insert_requests(batch)
    # Probe scopes are constructed disjointly; remaining rectangles are
    # generated by exact subtraction from disjoint Cartesian partitions.
    if total+sum(potential(p) for p in probes)!=potential(scope):
        raise RuntimeError("Prepared request coverage mismatch")
    return n


def prepare_all(db,transport,probe,output,log):
    catalogs=load_all_catalogs()
    fingerprint=digest([PLANNER_VERSION,catalogs])
    current=db.preparation(fingerprint)
    prep=current["id"]
    if current["status"]=="READY":
        log("Existing READY preparation reused; no rediscovery")
        return True
    db.recover_preparation(prep)
    errors=[]
    jobs=[]
    total_members=0
    total_fields=0
    unresolved=0
    for cube,(profile,catalog,fields) in catalogs.items():
        log(f"PREPARE {cube} | {len(fields)} catalog fields")
        dimensions,measures=classify_fields(fields)
        total_fields+=len(fields)
        unresolved+=sum(1 for f in fields if f.get("mdx") not in dimensions and f.get("mdx") not in measures)
        if profile["availability"]!="AVAILABLE" or catalog.get("source")=="PROFILE":
            db.save_catalog(prep,cube,profile,fields,measures)
            errors.append(f"{cube}: missing executable catalog")
            continue
        try:
            profile["year"]=resolve_year_hierarchy(profile,dimensions)
        except ValueError as exc:
            db.save_catalog(prep,cube,profile,fields,measures)
            errors.append(f"{cube}: {exc}")
            continue
        db.save_catalog(prep,cube,profile,fields,measures)
        pending=[h for h in dimensions if db.inventory(prep,cube,h) is None]
        hierarchies,levels,bulk=[],[],[]
        if pending:
            for kind,target in (("MDSCHEMA_HIERARCHIES",hierarchies),("MDSCHEMA_LEVELS",levels),("MDSCHEMA_MEMBERS",bulk)):
                try:
                    target.extend(discover_rows(transport,profile,kind))
                except (ValueError,RuntimeError) as exc:
                    log(f"{cube} {kind} fallback: {type(exc).__name__}")
        inventory={}
        for h in dimensions:
            members=db.inventory(prep,cube,h)
            if members is None:
                try:
                    meta=next((r for r in hierarchies if r.get("HIERARCHY_UNIQUE_NAME")==h),{})
                    own_levels=[r for r in levels if r.get("HIERARCHY_UNIQUE_NAME")==h]
                    members,metadata=discover_dimension_members(transport,profile,h,bulk,meta,own_levels)
                    db.save_inventory(prep,cube,h,"COMPLETE" if members else "EMPTY",metadata,own_levels,members)
                    log(f"{cube} {h} | MEMBERS {len(members)}")
                except (ValueError,RuntimeError) as exc:
                    db.save_inventory(prep,cube,h,"FAILED",{"error":str(exc)[:300]},[],[])
                    errors.append(f"{cube} {h}: {str(exc)[:200]}")
                    continue
            inventory[h]=members
            total_members+=len(members)
        # Year-only totals and one hierarchy at a time; never crossjoin year with itself.
        years=[m["unique_name"] for m in inventory.get(profile["year"],[])
               if not m["is_all_member"] and not m["is_calculated_member"]
               and (m["member_type"] is not None or (m["level_number"] or 0)>0)]
        if not years or not measures:
            errors.append(f"{cube}: no validated years or native measures")
            continue
        for h in [""]+[h for h in dimensions if h!=profile["year"]]:
            members=[m["unique_name"] for m in inventory.get(h,[])] if h else [""]
            if not members:
                continue
            scope={"hierarchy":h,"year_hierarchy":profile["year"],"members":members,
                   "years":years,"measures":sorted(measures)}
            job=digest([prep,cube,h,scope])
            db.save_job(prep,job,cube,h,scope)
            jobs.append((job,cube,profile,scope))
    summary={"preparation":prep,"catalog_fields":total_fields,"discovered_members":total_members,
             "unresolved_fields":unresolved,"jobs":len(jobs),"errors":errors}
    if errors:
        db.finalize_preparation(prep,"PREPARING",summary)
        (output/"preparation_summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
        log(f"PREPARE INCOMPLETE | {len(errors)} errors; rerun --prepare")
        return False
    budget=PROBE_BUDGET
    for job,cube,profile,scope in jobs:
        key=workload_group(len(scope["members"]))+f":{len(scope['years'])}:{len(scope['measures'])}"
        adaptive=db.load_profile(prep,cube,key)
        if adaptive is None:
            base=initial_profile(len(scope["members"]),len(scope["years"]),len(scope["measures"]))
            adaptive,budget=benchmark_profile(db,prep,job,cube,profile,scope,base,probe,budget)
            db.save_profile(prep,cube,key,adaptive)
        generate_request_plan(db,prep,job,cube,profile,scope,adaptive)
    summary["requests"]=db.progress(prep)
    summary["probe_budget"]=PROBE_BUDGET
    summary["recommended_concurrency"]=1
    summary["concurrency_source"]="CONSERVATIVE_NO_LIVE_CONCURRENCY_BENCHMARK"
    db.finalize_preparation(prep,"READY",summary,concurrency=1)
    (output/"preparation_summary.json").write_text(json.dumps(summary,ensure_ascii=False,indent=2),encoding="utf-8")
    log(f"PREPARE READY | {len(jobs)} jobs | {summary['requests']}")
    return True
