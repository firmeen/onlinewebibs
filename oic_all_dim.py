"""OIC All Dimension V2: prepare once, execute immutable XMLA requests.

Commands: --prepare, default run/resume, --status, --workers.
All extraction workers are visible Windows Terminal PowerShell tabs.
"""
from __future__ import annotations

import argparse
import base64
import hashlib
import json
import logging
import os
import re
import shutil
import sqlite3
import subprocess
import sys
import threading
import time
import uuid
import xml.etree.ElementTree as ET
from collections import deque
from contextlib import contextmanager
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation
from email.utils import parsedate_to_datetime
from pathlib import Path
from urllib.parse import urlparse

PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))

from curl_cffi import requests as curl_requests
from curl_cffi.const import CurlInfo, CurlOpt
from curl_cffi.requests.exceptions import RequestException
import psutil

from oic_all_dim_db import OICDatabase, digest
from oic_all_dim_prepare import prepare_all, MD_NS, checked_root, local

OUTPUT = PROJECT_ROOT / "output" / "oic_all_dim"
DB_PATH = OUTPUT / "oic.db"
BASE_URL = "https://onlinewebibs.oic.or.th"
XMLA_URL = BASE_URL + "/xmla-proxy/api/xmla-proxy"
PAGE_PATH = "/WebForm/Dashboard/Pivot/DynamicPivotCube.aspx"
TOKEN_URL = BASE_URL + "/WebForm/Dashboard/Pivot/PivotCube.asmx/GenToken"
DEFAULT_WORKERS = 4
MAX_RESPONSE_BYTES = 100 * 1024 * 1024
MAX_ATTEMPTS = 3
STARTUP_TIMEOUT = 60
TOKEN_REFRESH_MARGIN = 60
SESSION_RECHECK_SECONDS = 120
AUTH_LOCK_TIMEOUT = 90
USER_AGENT = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
              "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/157.0.0.0 Safari/537.36")
LOGGER = logging.getLogger("oic-all-dim")


def configure_log(path=None, console_output=True):
    for handler in LOGGER.handlers[:]:
        LOGGER.removeHandler(handler)
        handler.close()
    LOGGER.setLevel(logging.INFO)
    fmt=logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s",datefmt="%H:%M:%S")
    console=logging.StreamHandler(sys.stdout)
    console.setFormatter(fmt)
    if console_output:
        LOGGER.addHandler(console)
    if path:
        path.parent.mkdir(parents=True,exist_ok=True)
        handler=logging.FileHandler(path,encoding="utf-8")
        handler.setFormatter(fmt)
        LOGGER.addHandler(handler)


def log(message):
    LOGGER.info(message)


class PrepareMonitor:
    """Read-only observation of preparation; the Live refresh owns no DB connection."""
    def __init__(self,on_progress=None):
        from rich.console import Console
        self.console=Console()
        self.on_progress=on_progress
        self.interactive=on_progress is None and self.console.is_terminal
        self.lock=threading.RLock()
        self.started=time.monotonic()
        self.phase="INITIALIZING"
        self.cube="-"
        self.detail="Opening preparation"
        self.catalogs={}
        self.cubes={f"{c}/{y}":{"status":"PENDING","done":0,"total":None,"members":0}
                    for c in ("AH","FI","MR","MI","MV") for y in ("FY","UY","AY")}
        self.active={}
        self.events=deque(maxlen=8)
        self.samples=deque(maxlen=4)
        self.concurrency_samples=deque(maxlen=3)
        self.profiles={}
        self.selected=None
        self.plan={"index":0,"total":0,"primary":0,"recovery":0,"current_primary":0,"current_recovery":0}
        self.stats={"ok":0,"failed":0,"retries":0,"bytes":0,"http":0,"duration":0.0}
        self.summary={}
        self.discovery={}
        self.live=None
        self.handler=None

    def start(self):
        from rich.live import Live
        configure_log(OUTPUT/"prepare.log",console_output=not self.interactive and self.on_progress is None)
        if self.interactive or self.on_progress is not None:
            monitor=self
            class ActivityHandler(logging.Handler):
                def emit(self,record):
                    message=record.getMessage()
                    monitor.add_event(message)
                    if monitor.on_progress is not None:
                        monitor.on_progress("log",message=message)
            self.handler=ActivityHandler()
            LOGGER.addHandler(self.handler)
        if self.interactive:
            # Rich's independent refresh thread calls render during blocking CURL I/O.
            self.live=Live(get_renderable=self.render,console=self.console,refresh_per_second=3,
                           transient=True,redirect_stdout=False,redirect_stderr=False,
                           vertical_overflow="ellipsis")
            self.live.start(refresh=True)

    def add_event(self,message):
        with self.lock:
            self.events.append((time.strftime("%H:%M:%S"),str(message)[:240]))

    def emit(self,event,**data):
        message=None
        with self.lock:
            if event=="phase":
                self.phase=data["phase"]
                self.cube=data.get("cube",self.cube)
                self.detail=data.get("detail","")
                message=f"{self.phase} | {self.cube} | {self.detail}"
            elif event=="catalog":
                self.catalogs[data["cube"]]=dict(data)
                message=f"{data['cube']} | catalog {data['source']} | {data['fields']} fields"
            elif event=="cube":
                cube=data["cube"]
                self.cubes[cube].update({k:v for k,v in data.items() if k in ("status","total","done","members")})
                if "phase" in data:
                    self.phase=data["phase"]
                    self.cube=cube
                    self.detail=data.get("detail","")
                if "status" in data:
                    message=f"{cube} | {data['status']}"
            elif event=="sample":
                self.samples.append(dict(data["sample"],cube=data["cube"]))
                v=data["sample"]
                message=f"{data['cube']} | probe X={v['x']} Y={v['y']} M={v['m']} | {v['seconds']:.2f}s | {'OK' if v['ok'] else 'FAILED'}"
            elif event=="concurrency_sample":
                self.concurrency_samples.append(dict(data["sample"]))
                v=data["sample"]
                message=f"Concurrency {v['concurrency']} | {v['covered_per_second']:.1f} cells/s | {'OK' if v['ok'] else 'FAILED'}"
            elif event=="profile":
                key=(data["cube"],data["key"])
                self.profiles[key]=dict(data["profile"])
                self.selected=(data["cube"],dict(data["profile"]))
            elif event=="plan_job":
                self.phase="REQUEST_GENERATION"
                self.cube=data["cube"]
                self.detail=data["detail"]
                self.plan.update(index=data["index"],total=data["total"],current_primary=0,current_recovery=0)
            elif event=="plan_batch":
                self.plan.update(current_primary=data["primary"],current_recovery=data["recovery"])
            elif event=="plan_done":
                self.plan["primary"]+=self.plan["current_primary"]
                self.plan["recovery"]+=self.plan["current_recovery"]
                self.plan.update(current_primary=0,current_recovery=0)
            elif event=="summary":
                self.summary=dict(data["summary"])
                for cube,info in self.summary.get("cube_details",{}).items():
                    total=info["dimensions"]
                    done=info["inventories_complete"]
                    self.cubes[cube].update(total=total,done=done,members=info["members"],
                        status="INVENTORY READY" if done==total else "INCOMPLETE")
                for cube in self.summary.get("unavailable_cubes",[]):
                    self.cubes[cube]["status"]="UNAVAILABLE"
            elif event=="discovery":
                self.discovery=dict(data)
                message=data.get("message")
            elif event=="activity":
                message=data["message"]
        if message:
            # Logging handler acquires the monitor lock only after it is released.
            # File logging remains independent of the live renderer.
            log(message)
        if self.on_progress is not None and event in ("phase","cube","plan_job","summary","inventory_snapshot","inventory"):
            if event=="phase" and "cube" not in data:
                data=dict(data,cube=self.context_cube())
            self.on_progress(event,**data)

    def context(self):
        with self.lock:
            return self.phase,self.detail

    def context_cube(self):
        with self.lock:
            return self.cube

    def request_started(self,key,cube,detail,timeout,attempt=1,kind="DISCOVERY"):
        now=time.monotonic()
        with self.lock:
            self.active[key]={"cube":cube,"detail":detail,"kind":kind,"timeout":timeout,
                "attempt":attempt,"started":now,"stage_at":now,"stage":"INITIALIZING","bytes":0}
        log(f"{cube} | {kind} | {detail} | attempt {attempt} | START")

    def request_updated(self,key,stage,**data):
        with self.lock:
            row=self.active.get(key)
            if row is None:
                return
            if row["stage"]!=stage:
                row["stage"]=stage
                row["stage_at"]=time.monotonic()
            row.update(data)
        if stage not in ("RECEIVING",):
            suffix=f" | HTTP {data['http_status']}" if "http_status" in data else ""
            if "remaining_seconds" in data:
                suffix+=f" | JWT remaining {data['remaining_seconds']}s"
            log(f"{key[:12]} | {stage}{suffix}")

    def request_finished(self,key,ok,metrics=None,error=None):
        metrics=metrics or {}
        with self.lock:
            row=self.active.pop(key,None)
            if row is None:
                return
            if row["kind"]!="RECOVERY":
                self.stats["ok" if ok else "failed"]+=1
            if metrics.get("http_status")==200 and not metrics.get("replayed"):
                self.stats["http"]+=1
                self.stats["bytes"]+=metrics.get("response_bytes",0)
                self.stats["duration"]+=metrics.get("duration",0)
        timing=f" | {metrics['response_bytes']:,} bytes" if "response_bytes" in metrics else ""
        if "response_wait_estimate" in metrics:
            timing+=f" | response wait estimate {metrics['response_wait_estimate']:.2f}s"
        log(f"{row['cube']} | {row['kind']} | {'OK' if ok else 'FAILED'} | "
            f"{time.monotonic()-row['started']:.2f}s{timing} | {error or ''}")

    def retry(self,key,seconds):
        with self.lock:
            self.stats["retries"]+=1
        self.request_updated(key,"COOLDOWN",deadline=time.monotonic()+seconds)

    def render(self):
        from rich.console import Group
        from rich.panel import Panel
        from rich.table import Table
        from rich.text import Text
        from rich.progress_bar import ProgressBar
        # Copy only the bounded presentation state, not catalogs or payloads.
        with self.lock:
            phase,cube,detail=self.phase,self.cube,self.detail
            cubes={k:dict(v) for k,v in self.cubes.items()}
            catalogs={k:dict(v) for k,v in self.catalogs.items()}
            active=[dict(v) for v in self.active.values()]
            events=list(self.events)
            samples=list(self.samples)
            concurrency=list(self.concurrency_samples)
            selected=self.selected
            plan=dict(self.plan)
            stats=dict(self.stats)
            profile_count=len(self.profiles)
            summary=self.summary
            discovery=dict(self.discovery)
        now=time.monotonic()
        elapsed=int(now-self.started)
        total_fields=sum(v["fields"] for v in catalogs.values())
        ready=sum(v["status"]=="INVENTORY READY" for v in cubes.values())
        done=sum(v["done"] for v in cubes.values())
        members=sum(v["members"] for v in cubes.values())
        current=cubes.get(cube,{})
        source=catalogs.get(cube,{}).get("source","-")
        overview=Text()
        overview.append(f"{phase}  |  {cube}  |  {elapsed//3600:02}:{elapsed//60%60:02}:{elapsed%60:02}\n",style="bold cyan")
        overview.append(f"Catalogs {len(catalogs)}/15  |  Fields {total_fields:,}  |  Source {source}\n")
        overview.append(f"Inventories ready {ready}/15  |  Validated hierarchies {done:,}  |  Members {members:,}\n")
        overview.append(detail,style="yellow")
        height=self.console.height
        compact=height<40
        overview.no_wrap=True
        overview.overflow="ellipsis"
        current_total=current.get("total") or 0
        progress=Text(f"Current hierarchy coverage: {current.get('done',0)}/{current_total}" if current_total else "Current hierarchy coverage: pending")
        header=[overview,progress]
        if discovery and phase in ("CARDINALITY","MEMBER_DISCOVERY"):
            header.append(Text(discovery.get("message",""),no_wrap=True,overflow="ellipsis",style="cyan"))
        if current_total and not compact:
            header.append(ProgressBar(total=current_total,completed=current.get("done",0),width=40))
        parts=[Panel(Group(*header),title="OIC ALL DIMENSION V2 / PREPARE",border_style="cyan")]
        grid=Table(expand=True,show_edge=False,box=None,padding=(0,1))
        for h in ("Category","FY","UY","AY"):
            grid.add_column(h)
        for c in ("AH","FI","MR","MI","MV"):
            cells=[]
            for y in ("FY","UY","AY"):
                v=cubes[f"{c}/{y}"]
                status=v["status"]
                value=f"{v['done']}/{v['total']}" if status=="DISCOVERING" and v["total"] is not None else status
                color="green" if status=="INVENTORY READY" else "red" if status=="INCOMPLETE" else "yellow" if status=="DISCOVERING" else "dim"
                cells.append(Text(value,style=color))
            grid.add_row(c,*cells)
        coverage_panel=Panel(grid,title="CUBE INVENTORY COVERAGE")
        requests=[]
        for row in active:
            age=now-row["stage_at"]
            stage=row["stage"]
            line=Text(f"{row['cube']} | {row['kind']} | attempt {row['attempt']} | {stage}",style="bold yellow")
            requests.append(line)
            if not compact or len(active)==1:
                requests.append(Text(row["detail"],overflow="ellipsis",no_wrap=True))
            if stage=="COOLDOWN":
                requests.append(Text(f"Retry in {max(0,row.get('deadline',now)-now):.1f}s",style="yellow"))
            elif stage in ("REQUEST IN FLIGHT","AUTH REFRESH","AUTH LOCK WAIT"):
                limit=row.get("stage_timeout",row["timeout"])
                requests.append(Text(f"{age:.1f}s elapsed / {limit}s configured timeout (client wait, not server progress)"))
                if not compact or len(active)==1:
                    requests.append(ProgressBar(total=max(1,limit),completed=min(age,limit),width=40))
            else:
                requests.append(Text(f"{now-row['started']:.1f}s operation | {row['bytes']:,} bytes received"))
        if not requests:
            requests=[Text("No HTTP operation active — see current preparation stage.",style="dim")]
        parts.append(Panel(Group(*requests),title=f"ACTIVE OPERATIONS ({len(active)})",border_style="yellow" if active else "blue"))
        avg=f"{stats['duration']/stats['http']:.2f}s" if stats["http"] else "-"
        status=Text(f"Operations OK {stats['ok']} | Failed {stats['failed']} | Retries {stats['retries']}\n"
                    f"HTTP 200 recorded {stats['http']} | Avg {avg} | Received {stats['bytes']/1048576:.2f} MiB\n"
                    f"Profiles {profile_count} | Job {plan['index']}/{plan['total']} | "
                    f"Primary planned {plan['primary']+plan['current_primary']:,} | Recovery {plan['recovery']+plan['current_recovery']:,}")
        parts.append(Panel(status,title="STATISTICS & PLAN"))
        if not compact:
            parts.append(coverage_panel)
        if phase in ("ADAPTIVE","CONCURRENCY") and (height>=40 or not active):
            bench=Table(expand=True,box=None)
            for h in ("Cube","X/Y/M","Time","Coverage/s","Useful/s","Result"):
                bench.add_column(h)
            for v in samples:
                bench.add_row(v["cube"],f"{v['x']}/{v['y']}/{v['m']}",f"{v['seconds']:.1f}s",
                              f"{v['covered_per_second']:.1f}",f"{v['useful_cells']/max(.001,v['seconds']):.1f}",
                              "OK" if v["ok"] else "FAILED")
            lines=[bench]
            if selected:
                c,v=selected
                lines.append(Text(f"Selected {c}: X={v['x']} Y={v['y']} M={v['m']} | {v['timeout']}s | {v['source']} / {v['confidence']}",style="green"))
            for v in concurrency:
                lines.append(Text(f"Concurrency {v['concurrency']} | {v['seconds']:.1f}s | {v['covered_per_second']:.1f} cells/s | {'OK' if v['ok'] else 'FAILED'}"))
            parts.append(Panel(Group(*lines),title="04 / ADAPTIVE PROFILING"))
        # Prioritize active request visibility on short terminal windows.
        if compact and not active and phase not in ("ADAPTIVE","CONCURRENCY"):
            parts.append(coverage_panel)
        if height>=48:
            lines=[Text(f"{stamp} | {message}",overflow="ellipsis",no_wrap=True) for stamp,message in events[-(6 if self.console.height>=48 else 3):]]
            parts.append(Panel(Group(*(lines or [Text("No activity yet",style="dim")])),title="RECENT ACTIVITY"))
        parts.append(Text("Ctrl+C to stop | Checkpoints retained | prepare.log",style="dim"))
        return Group(*parts)

    def stop(self,status):
        from rich.table import Table
        with self.lock:
            self.phase=status
            summary=dict(self.summary)
            stats=dict(self.stats)
        if self.live is not None:
            self.live.stop()
        if self.handler is not None:
            LOGGER.removeHandler(self.handler)
            self.handler.close()
        if self.on_progress is not None:
            log(f"PREPARE {status} | checkpoints retained")
            self.on_progress("finished",status=status,summary=summary)
            configure_log(OUTPUT/"prepare.log",console_output=False)
            return
        configure_log(OUTPUT/"prepare.log")
        table=Table(title=f"PREPARATION {status}",show_header=False)
        table.add_column("Metric")
        table.add_column("Value")
        table.add_row("Elapsed",f"{time.monotonic()-self.started:.1f}s")
        for label,key in (("Catalog fields","catalog_fields"),("Available cubes","available_cubes"),
                          ("Discovered members","discovered_members"),("Logical jobs","jobs"),
                          ("Unresolved fields","unresolved_fields")):
            if key in summary:
                table.add_row(label,str(summary[key]))
        for name,info in summary.get("requests",{}).items():
            table.add_row(name,str(info["requests"]))
        if "coverage" in summary:
            table.add_row("Coverage",str(summary["coverage"]))
        table.add_row("Operations OK / failed",f"{stats['ok']} / {stats['failed']}")
        table.add_row("Database",str(DB_PATH))
        self.console.print(table)
        if status=="READY":
            log("READY FOR EXTRACTION | python oic_all_dim.py")
        else:
            log(f"PREPARE {status} | rerun --prepare to resume")
        for message in summary.get("errors",[])[:8]:
            log(f"UNRESOLVED | {message}")
        if len(summary.get("errors",[]))>8:
            log("See preparation_summary.json for remaining errors")


def read_env():
    values=dict(os.environ)
    path=PROJECT_ROOT/".env"
    if not path.exists():
        return values
    for line in path.read_text(encoding="utf-8-sig").splitlines():
        line=line.strip()
        if line.startswith("export "):
            line=line[7:].strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key,value=line.split("=",1)
        value=value.strip()
        if len(value)>=2 and value[0]==value[-1] and value[0] in ('"',"'"):
            value=value[1:-1]
        values[key.strip()]=value
    return values


def credential_values():
    env=read_env()
    return (env.get("OIC_COOKIE","").removeprefix("Cookie:").strip(),
            env.get("OIC_BEARER_TOKEN","").removeprefix("Bearer ").strip())


@contextmanager
def authentication_lock(stop=lambda:False):
    """Serialize GenToken across local workers/probes without sharing HTTP sessions."""
    OUTPUT.mkdir(parents=True,exist_ok=True)
    deadline=time.monotonic()+AUTH_LOCK_TIMEOUT
    with (OUTPUT/"auth.lock").open("a+b") as file:
        if file.seek(0,os.SEEK_END)==0:
            file.write(b"0")
            file.flush()
        locked=False
        try:
            while not locked:
                if stop():
                    raise CollectionError("USER_STOP")
                file.seek(0)
                try:
                    if os.name=="nt":
                        import msvcrt
                        msvcrt.locking(file.fileno(),msvcrt.LK_NBLCK,1)
                    else:
                        import fcntl
                        fcntl.flock(file.fileno(),fcntl.LOCK_EX|fcntl.LOCK_NB)
                    locked=True
                except OSError:
                    if time.monotonic()>=deadline:
                        raise CollectionError("AUTH_ENDPOINT_ERROR","Local authentication lock wait exceeded",5)
                    time.sleep(.2)
            yield
        finally:
            if locked:
                file.seek(0)
                if os.name=="nt":
                    msvcrt.locking(file.fileno(),msvcrt.LK_UNLCK,1)
                else:
                    fcntl.flock(file.fileno(),fcntl.LOCK_UN)


def publish_credentials(expected,cookie,token):
    """Update only the two credential keys, preserving all unrelated .env lines.

    Called under authentication_lock. Refuse to overwrite a concurrently
    changed credential pair (e.g. the user or the original Main Collector).
    """
    path=PROJECT_ROOT/".env"
    if not path.is_file() or credential_values()!=expected:
        return False
    if any("\r" in v or "\n" in v for v in (cookie,token)):
        raise CollectionError("AUTH_ENDPOINT_ERROR","Unexpected credential format",5)
    original=path.read_bytes()
    contents=original.decode("utf-8-sig")
    updates={"OIC_COOKIE":cookie,"OIC_BEARER_TOKEN":token}
    found=set()
    lines=[]
    newline="\r\n" if "\r\n" in contents else "\n"
    for line in contents.splitlines():
        match=re.match(r"^(\s*(?:export\s+)?)([A-Za-z_][A-Za-z0-9_]*)(\s*=).*$",line)
        if not match or match.group(2) not in updates:
            lines.append(line)
            continue
        key=match.group(2)
        if key not in found:
            lines.append(match.group(1)+key+match.group(3)+updates[key])
            found.add(key)
    lines.extend(key+"="+value for key,value in updates.items() if key not in found)
    encoded=(newline.join(lines)+newline).encode("utf-8")
    if original.startswith(b"\xef\xbb\xbf"):
        encoded=b"\xef\xbb\xbf"+encoded
    temporary=path.with_name(".env.oic."+uuid.uuid4().hex+".tmp")
    try:
        # Credentials are never written into the database or logs.
        fd=os.open(temporary,os.O_WRONLY|os.O_CREAT|os.O_EXCL,0o600)
        with os.fdopen(fd,"wb") as out:
            out.write(encoded)
            out.flush()
            os.fsync(out.fileno())
        if path.read_bytes()!=original or credential_values()!=expected:
            return False
        os.replace(temporary,path)
        return True
    finally:
        temporary.unlink(missing_ok=True)


def jwt_exp(token):
    try:
        payload=token.removeprefix("Bearer ").split(".")[1]
        return int(json.loads(base64.urlsafe_b64decode(payload+"="*(-len(payload)%4))).get("exp",0))
    except (ValueError,IndexError,TypeError,KeyError):
        return 0


def retry_after(value):
    if not value:
        return 0
    try:
        return max(0,float(value))
    except ValueError:
        try:
            return max(0,(parsedate_to_datetime(value)-datetime.now(timezone.utc)).total_seconds())
        except (ValueError,TypeError,OverflowError):
            return 0


class CollectionError(RuntimeError):
    def __init__(self,kind,message="",delay=0,*,endpoint=None,http_status=None):
        super().__init__(f"{kind}: {message}" if message else kind)
        self.kind=kind
        self.delay=delay
        self.endpoint=endpoint
        self.http_status=http_status


def extract_error_details(body):
    """Classify bounded error content; never return arbitrary server text."""
    text=body[:65536].decode("utf-8",errors="replace")
    lower=text.casefold()
    try:
        root=ET.fromstring(text)
    except ET.ParseError:
        # Recognize explicit small non-XML auth responses, not arbitrary data.
        if len(body)<=4096 and not lower.lstrip().startswith("<"):
            if "session_expired" in lower or "session expired" in lower:
                return "AUTH_COOKIE_EXPIRED"
            if "token expired" in lower or "expired token" in lower:
                return "AUTH_TOKEN_EXPIRED"
        return None
    faults=[e for e in root.iter() if local(e.tag) in ("Fault","Error")]
    if not faults:
        return None
    lower=" ".join(ET.tostring(e,encoding="unicode").casefold() for e in faults)
    if "session_expired" in lower or "session expired" in lower:
        return "AUTH_COOKIE_EXPIRED"
    if "token expired" in lower or "expired token" in lower:
        return "AUTH_TOKEN_EXPIRED"
    # Interpret known phrases only inside a structured XMLA fault.
    if any(v in lower for v in ("not supported","not implemented","unsupported rowset","invalid restriction")):
        return "DISCOVER_NOT_SUPPORTED"
    if any(v in lower for v in ("resource limit","query too complex","maximum number of cells",
                                 "memory limit","response too large","query timeout","query timed out")):
        return "QUERY_SCOPE_ERROR"
    if any(v in lower for v in ("access denied","permission denied","not authorized")):
        return "ACCESS_DENIED"
    if any(v in lower for v in ("syntax error","cannot be found","not found in the cube",
                                 "unknown member","invalid member","invalid measure")):
        return "QUERY_FAULT"
    return "XMLA_FAULT_UNKNOWN"


def bounded_error_body(response):
    result=bytearray()
    for chunk in response.iter_content():
        remaining=65536-len(result)
        result.extend(chunk[:remaining])
        if len(result)>=65536:
            break
    return bytes(result)


def classify_http(status,headers,endpoint="XMLA",body=b""):
    delay=max(5,retry_after(headers.get("Retry-After")))
    diagnostic=extract_error_details(body)
    if status in (301,302,303,307,308) and endpoint=="GEN_TOKEN":
        kind="AUTH_COOKIE_EXPIRED"
    elif status==401:
        kind="ACCESS_DENIED" if endpoint=="GEN_TOKEN" else "AUTH_TOKEN_EXPIRED"
    elif status==403:
        kind="ACCESS_DENIED"
    elif status==429:
        kind="RATE_LIMITED"
    elif status in (408,425,502,503,504):
        kind="SERVER_UNAVAILABLE"
    elif endpoint=="GEN_TOKEN":
        kind=diagnostic if diagnostic in ("AUTH_COOKIE_EXPIRED","AUTH_TOKEN_EXPIRED","ACCESS_DENIED") else "AUTH_ENDPOINT_ERROR"
    elif status==500:
        kind=diagnostic if diagnostic not in (None,"XMLA_FAULT_UNKNOWN") else "HTTP_500_UNKNOWN"
    else:
        kind=diagnostic or "INVALID_QUERY"
    return CollectionError(kind,f"{endpoint} HTTP {status}"+(f" | {diagnostic}" if diagnostic else ""),
                           delay,endpoint=endpoint,http_status=status)


def classify_network(exc):
    code=int(getattr(exc,"code",0) or 0)
    kind={5:"DNS_ERROR",6:"DNS_ERROR",7:"CONNECTION_RESET",28:"RESPONSE_TIMEOUT",
          35:"TLS_ERROR",51:"TLS_ERROR",60:"TLS_ERROR",56:"CONNECTION_RESET",
          55:"CONNECTION_RESET",18:"CONNECTION_RESET"}.get(code,"NETWORK_ERROR")
    # Do not log exception text: HTTP libraries may include auth headers.
    return CollectionError(kind,f"CURL code {code}")


class OICClient:
    def __init__(self,periodic=False,stop=None):
        self.session=None
        self.observer=None
        self.cookie=""
        self.token=""
        self.web_class=None
        self.env_signature=None
        self.env_token=""
        self.last_auth_check=time.monotonic()
        self.periodic=periodic
        self.stop=stop or (lambda:False)
        self.reload_credentials(force=True)

    def observe(self,stage,**data):
        if self.observer is not None:
            self.observer(stage,**data)

    def reload_credentials(self,force=False):
        path=PROJECT_ROOT/".env"
        st=path.stat() if path.exists() else None
        signature=(st.st_mtime_ns,st.st_size) if st else None
        if not force and signature==self.env_signature:
            return False
        cookie,token=credential_values()
        changed=cookie!=self.cookie
        token_changed=token!=self.env_token
        self.env_signature=signature
        self.cookie=cookie
        self.env_token=token
        if changed or token_changed or force:
            self.token=token
            self.last_auth_check=time.monotonic()
        if changed or self.session is None:
            self.reconnect(preserve_cookies=False)
        if (changed or token_changed) and not force:
            log("AUTH | external credential update loaded")
        return changed or token_changed

    def load_cookie(self,raw):
        hostname=urlparse(BASE_URL).hostname
        for part in raw.split(";"):
            name,separator,value=part.strip().partition("=")
            if separator and name.strip():
                self.session.cookies.set(name.strip(),value.strip(),domain=hostname,path="/",secure=True)

    def cookie_header(self):
        hostname=urlparse(BASE_URL).hostname
        values={}
        for cookie in self.session.cookies.jar:
            if cookie.domain.lstrip(".")==hostname and not cookie.is_expired():
                values[cookie.name]=cookie.value
        return "; ".join(f"{key}={value}" for key,value in values.items())

    def reconnect(self,preserve_cookies=True):
        previous=None
        if self.session is not None:
            if preserve_cookies:
                # Preserve full cookie attributes rather than re-parsing a header.
                previous=list(self.session.cookies.jar)
            self.session.close()
        info_names=("NAMELOOKUP_TIME","CONNECT_TIME","APPCONNECT_TIME","PRETRANSFER_TIME","STARTTRANSFER_TIME","TOTAL_TIME")
        infos=[getattr(CurlInfo,n) for n in info_names if hasattr(CurlInfo,n)]
        self.session=curl_requests.Session(verify=True,curl_infos=infos,
            curl_options={CurlOpt.TCP_KEEPALIVE:1})
        self.session.headers.update({"User-Agent":USER_AGENT,"Accept-Language":"th-TH,th;q=0.9"})
        if previous is not None:
            for cookie in previous:
                self.session.cookies.jar.set_cookie(cookie)
        else:
            self.load_cookie(self.cookie)

    def set_cube(self,profile):
        cls=profile.get("cls") or profile.get("web_class")
        if not cls:
            raise CollectionError("INVALID_QUERY","Missing Web Class")
        # Reference collector reuses a valid JWT across categories and FY/UY/AY.
        # The selected class changes Referer/GenToken context, not token validity.
        self.web_class=cls

    def referer(self):
        return BASE_URL+PAGE_PATH+"?Class="+self.web_class

    def refresh_token(self,rejected_token=None,periodic=False):
        before=self.token
        self.observe("AUTH LOCK WAIT",stage_timeout=AUTH_LOCK_TIMEOUT)
        with authentication_lock(self.stop):
            self.reload_credentials()
            if self.stop():
                raise CollectionError("USER_STOP")
            now=time.time()
            # Another worker may already have renewed and published credentials.
            if self.token!=before and self.token!=rejected_token and jwt_exp(self.token)>now+TOKEN_REFRESH_MARGIN:
                self.observe("REUSING SHARED TOKEN",remaining_seconds=int(jwt_exp(self.token)-now))
                return
            if rejected_token is None and not periodic and jwt_exp(self.token)>now+TOKEN_REFRESH_MARGIN:
                return
            if (periodic and rejected_token is None and jwt_exp(self.token)>now+TOKEN_REFRESH_MARGIN
                    and time.monotonic()-self.last_auth_check<SESSION_RECHECK_SECONDS):
                return
            self.observe("AUTH REFRESH",stage_timeout=45)
            if not self.cookie:
                raise CollectionError("AUTH_COOKIE_EXPIRED","Set OIC_COOKIE in .env")
            expected=(self.cookie,self.env_token)
            try:
                response=self.session.post(TOKEN_URL,
                    json={"_tk":"Init","_pageUrl":PAGE_PATH+"?Class="+self.web_class},
                    headers={"Accept":"application/json, text/javascript, */*; q=0.01",
                             "Content-Type":"application/json; charset=UTF-8","X-Requested-With":"XMLHttpRequest",
                             "Origin":BASE_URL,"Referer":self.referer()},
                    timeout=(15,45),allow_redirects=False)
            except RequestException as exc:
                error=classify_network(exc)
                if error.kind=="TLS_ERROR":
                    raise CollectionError("TLS_ERROR","GEN_TOKEN TLS verification failure",endpoint="GEN_TOKEN") from None
                raise CollectionError("AUTH_ENDPOINT_ERROR",f"GEN_TOKEN | {error.kind}",5,endpoint="GEN_TOKEN") from None
            if response.status_code!=200:
                raise classify_http(response.status_code,response.headers,"GEN_TOKEN",response.content[:65536])
            try:
                token=response.json().get("d","")
            except (ValueError,AttributeError):
                raise CollectionError("AUTH_ENDPOINT_ERROR","GenToken did not return usable JSON",5,endpoint="GEN_TOKEN") from None
            if token=="SESSION_EXPIRED":
                raise CollectionError("AUTH_COOKIE_EXPIRED","Update OIC_COOKIE in .env",endpoint="GEN_TOKEN")
            if not isinstance(token,str):
                raise CollectionError("AUTH_ENDPOINT_ERROR","GenToken returned unexpected token type",5,endpoint="GEN_TOKEN")
            token=token.removeprefix("Bearer ").strip()
            if len(token.split("."))!=3 or jwt_exp(token)<=time.time()+30:
                raise CollectionError("AUTH_ENDPOINT_ERROR","GenToken returned unusable JWT",5,endpoint="GEN_TOKEN")
            self.token=token
            self.last_auth_check=time.monotonic()
            current_cookie=self.cookie_header()
            if not current_cookie:
                raise CollectionError("AUTH_COOKIE_EXPIRED","GenToken session contains no usable cookies",endpoint="GEN_TOKEN")
            try:
                published=publish_credentials(expected,current_cookie,token)
            except OSError:
                published=False
                log("AUTH | JWT renewed locally; .env update unavailable")
            if published:
                self.cookie=current_cookie
                self.env_token=token
                st=(PROJECT_ROOT/".env").stat()
                self.env_signature=(st.st_mtime_ns,st.st_size)
            elif credential_values()!=expected:
                self.reload_credentials(force=True)
                if jwt_exp(self.token)<=time.time()+10:
                    raise CollectionError("AUTH_ENDPOINT_ERROR","Credential changed during refresh; retry with latest .env",5)
            log(f"AUTH | GenToken successful | JWT remaining {max(0,int(jwt_exp(self.token)-time.time()))}s | "+("shared .env updated" if published else "local credential retained"))

    def ensure_token(self):
        remaining=jwt_exp(self.token)-time.time()
        periodic_due=self.periodic and time.monotonic()-self.last_auth_check>=SESSION_RECHECK_SECONDS
        if remaining>TOKEN_REFRESH_MARGIN and not periodic_due:
            self.observe("REUSING VALID TOKEN",remaining_seconds=int(remaining))
            return
        try:
            self.refresh_token(periodic=periodic_due)
        except CollectionError as exc:
            # A transient renewal outage does not invalidate an existing JWT.
            if exc.kind=="AUTH_ENDPOINT_ERROR" and jwt_exp(self.token)>time.time()+TOKEN_REFRESH_MARGIN:
                self.last_auth_check=time.monotonic()
                self.observe("REUSING VALID TOKEN",remaining_seconds=int(jwt_exp(self.token)-time.time()))
                return
            raise

    def execute_once(self,profile,payload,timeout,destination=None):
        if self.stop():
            raise CollectionError("USER_STOP")
        self.reload_credentials()
        self.set_cube(profile)
        overall=time.monotonic()
        token_start=time.monotonic()
        self.ensure_token()
        auth_seconds=time.monotonic()-token_start
        for attempt in range(2):
            if self.stop():
                raise CollectionError("USER_STOP")
            started=time.monotonic()
            self.observe("REQUEST IN FLIGHT",stage_timeout=timeout,bytes=0)
            try:
                with self.session.stream("POST",XMLA_URL,data=payload,
                    headers={"Accept":"text/plain, */*; q=0.01","Content-Type":"text/xml",
                             "X-Requested-With":"XMLHttpRequest","Origin":BASE_URL,"Referer":self.referer(),
                             "Authorization":"Bearer "+self.token},
                    timeout=(15,timeout),allow_redirects=False) as response:
                    headers_at=time.monotonic()
                    self.observe("HEADERS RECEIVED",http_status=response.status_code)
                    if response.status_code!=200:
                        error=classify_http(response.status_code,response.headers,"XMLA",bounded_error_body(response))
                        if error.kind in ("AUTH_TOKEN_EXPIRED","ACCESS_DENIED") and attempt==0:
                            retry_auth=True
                            rejected_token=self.token
                        else:
                            raise error
                    else:
                        retry_auth=False
                        size=0
                        body=bytearray() if destination is None else None
                        temporary=None
                        stream=None
                        write_seconds=0.0
                        try:
                            if destination:
                                destination.parent.mkdir(parents=True,exist_ok=True)
                                temporary=destination.with_name(destination.name+"."+uuid.uuid4().hex+".partial")
                                stream=temporary.open("wb")
                            for chunk in response.iter_content():
                                if not chunk:
                                    continue
                                size+=len(chunk)
                                self.observe("RECEIVING",bytes=size)
                                if size>MAX_RESPONSE_BYTES:
                                    raise CollectionError("RESPONSE_TOO_LARGE","100 MiB limit")
                                if stream:
                                    t=time.monotonic()
                                    stream.write(chunk)
                                    write_seconds+=time.monotonic()-t
                                else:
                                    body.extend(chunk)
                            received=time.monotonic()
                            if stream:
                                stream.flush()
                                os.fsync(stream.fileno())
                                stream.close()
                                stream=None
                                temporary.replace(destination)
                            metrics={"duration":time.monotonic()-overall,"http_status":200,
                                     "response_bytes":size,"token_refresh_seconds":auth_seconds,
                                     "header_wait_seconds":headers_at-started,
                                     "receive_seconds":received-headers_at,"file_write_seconds":write_seconds}
                            infos=getattr(response,"infos",{}) or {}
                            for name,key in (("NAMELOOKUP_TIME","dns_seconds"),("CONNECT_TIME","connect_seconds"),
                                             ("APPCONNECT_TIME","tls_complete_seconds"),("PRETRANSFER_TIME","pretransfer_seconds"),
                                             ("STARTTRANSFER_TIME","ttfb_seconds"),("TOTAL_TIME","curl_total_seconds")):
                                value=infos.get(getattr(CurlInfo,name,None))
                                if isinstance(value,(int,float)) and value>=0:
                                    metrics[key]=value
                            if "ttfb_seconds" in metrics and "pretransfer_seconds" in metrics:
                                metrics["response_wait_estimate"]=max(0,metrics["ttfb_seconds"]-metrics["pretransfer_seconds"])
                            return (bytes(body) if body is not None else None),metrics
                        finally:
                            if stream:
                                stream.close()
                            if temporary:
                                temporary.unlink(missing_ok=True)
                if retry_auth:
                    t=time.monotonic()
                    self.refresh_token(rejected_token=rejected_token)
                    auth_seconds+=time.monotonic()-t
            except RequestException as exc:
                raise classify_network(exc) from None
            except OSError as exc:
                raise CollectionError("DISK_FULL" if exc.errno==28 else "STORAGE_ERROR",type(exc).__name__) from None
        raise CollectionError("ACCESS_DENIED","Authentication retry exhausted")

    def discover(self,profile,payload,timeout):
        data,_=self.execute_once(profile,payload,timeout)
        return data

    def close(self):
        if self.session is not None:
            self.session.close()


def parse_xmla(data,scope):
    try:
        root=checked_root(data)
    except (ET.ParseError,ValueError) as exc:
        kind="SOAP_FAULT" if "fault/error" in str(exc).lower() else "INVALID_XML"
        raise CollectionError(kind,"Invalid XMLA dataset or server fault") from None
    dataset=root.find(f".//{{{MD_NS}}}root")
    if dataset is None:
        raise CollectionError("INVALID_XML","Missing multidimensional dataset")
    axes={}
    for axis in dataset.findall(f"./{{{MD_NS}}}Axes/{{{MD_NS}}}Axis"):
        tuples=[]
        for item in axis.findall(f"{{{MD_NS}}}Tuples/{{{MD_NS}}}Tuple"):
            members=[]
            for node in item.findall(f"{{{MD_NS}}}Member"):
                members.append((node.get("Hierarchy",""),node.findtext(f"{{{MD_NS}}}UName") or ""))
            tuples.append(members)
        axes[axis.get("name")]=tuples
    if "Axis0" not in axes or "Axis1" not in axes:
        raise CollectionError("INVALID_XML","Missing expected axes")
    years=[]
    for row in axes["Axis0"]:
        if len(row)!=1 or row[0][1] not in scope["years"]:
            raise CollectionError("INVALID_XML","Unexpected year tuple")
        years.append(row[0][1])
    rows=[]
    for row in axes["Axis1"]:
        measures=[name for _,name in row if name in scope["measures"]]
        members=[name for _,name in row if scope["hierarchy"] and name in scope["members"]]
        if len(measures)!=1 or len(row)!=(2 if scope["hierarchy"] else 1) or len(members)!=(1 if scope["hierarchy"] else 0):
            raise CollectionError("INVALID_XML","Unexpected measure/member tuple")
        rows.append((members[0] if members else "",measures[0]))
    expected={(m,v) for m in scope["members"] for v in scope["measures"]}
    # No NON EMPTY in generated requests: missing axes may indicate truncation.
    if len(years)!=len(set(years)) or set(years)!=set(scope["years"]) or len(rows)!=len(set(rows)) or set(rows)!=expected:
        raise CollectionError("INVALID_XML","Response axes do not cover the prepared request")
    cells={}
    width=len(years)
    for cell in dataset.findall(f"./{{{MD_NS}}}CellData/{{{MD_NS}}}Cell"):
        try:
            ordinal=int(cell.get("CellOrdinal","-1"))
        except ValueError:
            raise CollectionError("INVALID_XML","Invalid CellOrdinal") from None
        if ordinal<0 or ordinal>=width*len(rows) or ordinal in cells:
            raise CollectionError("INVALID_XML","Duplicate/out-of-range CellOrdinal")
        value=cell.find(f"{{{MD_NS}}}Value")
        formatted=cell.findtext(f"{{{MD_NS}}}FmtValue")
        cells[ordinal]=(value.text if value is not None else None,formatted)
    updated=next((n.text for n in root.iter() if local(n.tag)=="LastDataUpdate"),None)
    return years,rows,cells,updated


def iter_database_records(parsed,scope,all_members):
    years,rows,cells,updated=parsed
    width=len(years)
    for i,(member,measure) in enumerate(rows):
        for j,year in enumerate(years):
            ordinal=i*width+j
            raw,formatted=cells.get(ordinal,(None,None))
            numeric=None
            state="MISSING" if ordinal not in cells else "NULL" if raw is None else "VALID"
            if raw is not None:
                try:
                    value=Decimal(raw)
                    if value.is_finite():
                        numeric=str(value)
                        if value==0:
                            state="ZERO"
                except InvalidOperation:
                    pass
            coordinates=[(scope["year_hierarchy"],year)]
            if scope["hierarchy"]:
                coordinates.append((scope["hierarchy"],member))
            yield {"coordinates":coordinates,"grain":scope["hierarchy"] or "YEAR_ONLY",
                   "is_total":not scope["hierarchy"] or member in all_members,
                   "measure":measure,"raw":raw,"numeric":numeric,"formatted":formatted,
                   "state":state,"last_data_update":updated}


def raw_path(r):
    return OUTPUT/"raw"/r["prep"]/r["cube"]/r["id"]/(r["payload_hash"]+"."+r["lease"]+".xml")


def replay_raw_response(r):
    candidates=[]
    if r.get("response_path") and r.get("response_hash"):
        candidates.append((Path(r["response_path"]),r["response_hash"]))
    # Recover a fully renamed response if the previous worker died after the
    # atomic file write but before registering it in the database.
    folder=raw_path(r).parent
    if folder.is_dir():
        candidates.extend((p,None) for p in sorted(folder.glob(r["payload_hash"]+".*.xml")))
    for path,expected_hash in candidates:
        if not path.resolve().is_relative_to(OUTPUT.resolve()) or not path.is_file() or path.stat().st_size>MAX_RESPONSE_BYTES:
            continue
        if not path.name.startswith(r["payload_hash"]+"."):
            continue
        data=path.read_bytes()
        sha=digest(data)
        if expected_hash and sha!=expected_hash:
            continue
        try:
            parsed=parse_xmla(data,json.loads(r["scope_json"]))
        except CollectionError:
            continue
        return path,parsed,sha
    return None


def stopped(db,rid):
    run=db.run(rid)
    return not run or bool(run["stop"]) or run["state"]!="RUNNING"


def wait_seconds(seconds,check=lambda:False):
    end=time.monotonic()+seconds
    while time.monotonic()<end:
        if check():
            raise CollectionError("USER_STOP")
        time.sleep(min(0.5,max(0,end-time.monotonic())))


def process_request(db,client,r,rid=None,owner=None):
    scope=json.loads(r["scope_json"])
    profile=db.profile(r["prep"],r["cube"])
    if digest(r["payload"])!=r["payload_hash"]:
        raise CollectionError("INVALID_QUERY","Prepared payload hash mismatch")
    started=time.monotonic()
    client.observe("CHECKING RAW RESPONSE")
    replay=replay_raw_response(r)
    if replay:
        path,parsed,sha=replay
        metrics=json.loads(r.get("metrics_json") or "{}")
        metrics["replayed"]=True
        db.register_response(r,path,sha,metrics)
    else:
        path=raw_path(r)
        if rid:
            while not db.acquire_slot(rid,owner):
                if stopped(db,rid):
                    raise CollectionError("USER_STOP")
                time.sleep(0.5)
        try:
            log(f"{r['cube']} | {r['id'][:12]} | REQUESTING X={len(scope['members'])} Y={len(scope['years'])} M={len(scope['measures'])}")
            _,metrics=client.execute_once(profile,r["payload"],r["timeout"],path)
        finally:
            if rid:
                db.release_slot(owner)
        client.observe("PARSING")
        data=path.read_bytes()
        db.register_response(r,path,digest(data),metrics)
        t=time.monotonic()
        parsed=parse_xmla(data,scope)
        metrics["parse_seconds"]=time.monotonic()-t
    all_members=set()
    if scope["hierarchy"]:
        all_members={m["unique_name"] for m in (db.inventory(r["prep"],r["cube"],scope["hierarchy"]) or []) if m["is_all_member"]}
    updated=parsed[3]
    if updated:
        prior=db.conn.execute("SELECT last_data_update FROM oic_observations WHERE prep=? AND cube=? AND last_data_update IS NOT NULL LIMIT 1",
                              (r["prep"],r["cube"])).fetchone()
        if prior and prior[0]!=updated:
            raise CollectionError("SOURCE_VERSION_CHANGED","Cube update marker changed; reprepare required")
    metrics["covered_cells"]=r["potential_cells"]
    metrics["useful_cells"]=sum(v[0] is not None for v in parsed[2].values())
    metrics["source_update"]=updated
    t=time.monotonic()
    client.observe("COMMITTING")
    # A busy write must not result in another HTTP request.
    for attempt in range(3):
        try:
            count=db.commit_request(r,iter_database_records(parsed,scope,all_members),metrics)
            break
        except ValueError as exc:
            if str(exc)=="SOURCE_VERSION_CHANGED":
                raise CollectionError("SOURCE_VERSION_CHANGED","Cube update marker changed; reprepare required") from None
            raise
        except sqlite3.OperationalError as exc:
            if "locked" not in str(exc).lower() and "busy" not in str(exc).lower():
                raise CollectionError("DATABASE_WRITE_ERROR",type(exc).__name__) from None
            if attempt==2:
                raise CollectionError("DATABASE_BUSY") from None
            time.sleep(1+attempt)
    metrics["db_commit_seconds"]=time.monotonic()-t
    metrics["wall_seconds"]=time.monotonic()-started
    try:
        db.conn.execute("UPDATE oic_requests SET metrics_json=? WHERE id=? AND status IN ('COMMITTED','EMPTY')",
                        (json.dumps(metrics,ensure_ascii=False),r["id"]))
    except sqlite3.Error:
        pass  # Completion is already committed; diagnostics are best effort.
    log(f"{r['cube']} | {r['id'][:12]} | COMMITTED {count:,} coordinates | useful={metrics['useful_cells']:,} | HTTP={metrics.get('duration','N/A')}s")
    return metrics


def handle_request_failure(db,client,r,exc,rid,probe=False):
    kind=exc.kind if isinstance(exc,CollectionError) else "UNKNOWN_ERROR"
    message=str(exc) if isinstance(exc,CollectionError) else type(exc).__name__
    log(f"{r['cube']} | {r['id'][:12]} | {message}")
    if probe:
        db.fail(r,message,"PROBE" if kind=="USER_STOP" else "PROBE_FAILED")
        return
    if kind=="USER_STOP":
        db.fail(r,message,"RETRY")
    elif kind=="AUTH_COOKIE_EXPIRED":
        db.fail(r,message,"RETRY",5)
    elif kind in ("ACCESS_DENIED","AUTH_TOKEN_EXPIRED"):
        # Persistent rejection after bounded refresh requires user action.
        db.fail(r,message,"RETRY",5)
        db.stop(rid)
    elif kind in ("RATE_LIMITED","SERVER_UNAVAILABLE","HTTP_500_UNKNOWN","AUTH_ENDPOINT_ERROR"):
        minimum=min(120,5*(2**r["attempts"])) if kind in ("RATE_LIMITED","SERVER_UNAVAILABLE") else 5
        delay=max(minimum,exc.delay,2**r["attempts"])
        db.pressure(rid,delay)
        db.fail(r,message,"RETRY" if r["attempts"]<MAX_ATTEMPTS else "FAILED",delay)
    elif kind in ("RESPONSE_TOO_LARGE","SOAP_FAULT","QUERY_FAULT","QUERY_SCOPE_ERROR","RESPONSE_TIMEOUT","INVALID_QUERY"):
        if kind=="RESPONSE_TIMEOUT" and r["attempts"]<2:
            db.pressure(rid,10)
            db.fail(r,message,"RETRY",10)
        elif not db.activate_recovery_requests(r,message):
            db.fail(r,message,"NEEDS_REPREPARE")
    elif kind in ("NETWORK_ERROR","DNS_ERROR","CONNECTION_RESET","INVALID_XML","DATABASE_BUSY"):
        if kind in ("NETWORK_ERROR","DNS_ERROR","CONNECTION_RESET"):
            client.reconnect()
        db.fail(r,message,"RETRY" if r["attempts"]<MAX_ATTEMPTS else "FAILED",2**r["attempts"])
    elif kind=="SOURCE_VERSION_CHANGED":
        db.fail(r,message,"NEEDS_REPREPARE")
        db.stop(rid)
    else:
        db.fail(r,message,"FAILED")
        if kind in ("STORAGE_ERROR","DISK_FULL","DATABASE_WRITE_ERROR"):
            db.stop(rid)


def execute_resilient(db,client,r,rid=None,owner=None,probe=False):
    try:
        metrics=process_request(db,client,r,rid,owner)
        return {"ok":True,**metrics}
    except (CollectionError,sqlite3.Error,OSError,ValueError) as exc:
        handle_request_failure(db,client,r,exc,rid,probe)
        return {"ok":False,"pressure":getattr(exc,"kind","") in ("RATE_LIMITED","SERVER_UNAVAILABLE","HTTP_500_UNKNOWN","AUTH_ENDPOINT_ERROR","RESPONSE_TIMEOUT"),
                "error":getattr(exc,"kind",type(exc).__name__),"delay":getattr(exc,"delay",0)}


def process_alive(pid,created):
    try:
        p=psutil.Process(pid)
        return abs(p.create_time()-created)<0.01 and p.is_running() and p.status()!=psutil.STATUS_ZOMBIE
    except psutil.NoSuchProcess:
        return False
    except psutil.AccessDenied:
        # Do not reclaim a request merely because process inspection is denied.
        return True


def heartbeat_loop(owner,done):
    db=OICDatabase(DB_PATH)
    try:
        while not done.wait(2):
            try:
                db.heartbeat(owner)
            except sqlite3.Error:
                pass
    finally:
        db.close()


def run_worker(rid,worker):
    configure_log(OUTPUT/"runs"/rid/f"worker_{worker:02}.log")
    db=OICDatabase(DB_PATH)
    owner=uuid.uuid4().hex
    done=threading.Event()
    client=None
    thread=None
    registered=False
    final="FAILED"
    try:
        run=db.run(rid)
        if not run or run["state"]!="RUNNING":
            raise RuntimeError("Unknown active run")
        client=OICClient(stop=lambda:stopped(db,rid))
        if not client.cookie:
            raise RuntimeError("OIC_COOKIE not configured")
        db.register_worker(rid,worker,owner,os.getpid(),psutil.Process().create_time())
        registered=True
        thread=threading.Thread(target=heartbeat_loop,args=(owner,done),daemon=True)
        thread.start()
        log(f"W{worker:02} READY | curl_cffi | waiting for barrier")
        while not db.run(rid)["barrier"]:
            if stopped(db,rid):
                final="STOPPED"
                return 0
            time.sleep(0.25)
        while not stopped(db,rid):
            r=db.claim_next_request(rid,owner)
            if r is None:
                if db.all_work_finished(run["prep"]):
                    final="COMPLETED"
                    return 0
                time.sleep(0.5)
                continue
            result=execute_resilient(db,client,r,rid,owner)
            if result.get("error")=="AUTH_COOKIE_EXPIRED":
                log("WAITING FOR CREDENTIAL | update local .env")
                while not stopped(db,rid):
                    if client.reload_credentials():
                        break
                    time.sleep(2)
        final="STOPPED"
        return 0
    except KeyboardInterrupt:
        final="STOPPED"
        return 130
    finally:
        done.set()
        if thread:
            thread.join(timeout=5)
        if client:
            client.close()
        if registered:
            # Supervisor fences and requeues remaining work if this process dies.
            db.finish_worker(owner,final)
        db.close()


def powershell_quote(value):
    return "'"+str(value).replace("'","''")+"'"


def worker_command(rid,worker):
    command=("& "+powershell_quote(sys.executable)+" "+powershell_quote(__file__)+
             " --worker-id "+str(worker)+" --run-id "+powershell_quote(rid)+"; exit $LASTEXITCODE")
    return base64.b64encode(command.encode("utf-16-le")).decode("ascii")


def launch_worker_tabs(rid,workers):
    wt=shutil.which("wt.exe") or shutil.which("wt")
    ps=shutil.which("pwsh.exe") or shutil.which("powershell.exe")
    if os.name!="nt" or not wt or not ps:
        raise RuntimeError("Visible Windows Terminal and PowerShell are required; no hidden fallback")
    args=[wt,"-w","OIC-"+rid[:12]]
    for index,worker in enumerate(workers):
        if index:
            args.append(";")
        args.extend(["new-tab","--startingDirectory",str(PROJECT_ROOT),"--title",f"OIC W{worker:02}",
                     ps,"-NoLogo","-NoProfile","-EncodedCommand",worker_command(rid,worker)])
    result=subprocess.run(args,check=False,timeout=30)
    if result.returncode:
        raise RuntimeError(f"Windows Terminal launcher exited {result.returncode}")


def wait_workers_ready(db,rid,count):
    end=time.monotonic()+STARTUP_TIMEOUT
    while time.monotonic()<end:
        rows=db.workers(rid)
        if len(rows)==count and all(w["status"]=="READY" and process_alive(w["pid"],w["created"]) for w in rows):
            db.barrier(rid)
            log(f"START BARRIER RELEASED | {count} visible workers")
            return
        time.sleep(0.5)
    raise RuntimeError("Worker startup deadline exceeded")


def recover_old_runs(db):
    for run in db.active_runs():
        if process_alive(run["supervisor_pid"],run["supervisor_created"]):
            raise RuntimeError("An extraction supervisor is already active")
        db.stop(run["id"])
        for w in db.workers(run["id"]):
            if process_alive(w["pid"],w["created"]):
                raise RuntimeError("Previous visible workers are stopping; run again after they exit")
            db.recover_worker(w["owner"])
        db.finish_run(run["id"],"INTERRUPTED")


def supervise_workers(db,rid,count):
    run=db.run(rid)
    replacements={i:0 for i in range(1,count+1)}
    launch_deadlines={}
    next_log=0
    while not stopped(db,rid):
        rows={w["worker"]:w for w in db.workers(rid)}
        for number,w in rows.items():
            if process_alive(w["pid"],w["created"]):
                continue
            db.recover_worker(w["owner"])
            if not db.all_work_finished(run["prep"]):
                if replacements[number]>=2:
                    raise RuntimeError(f"W{number:02} replacement budget exhausted")
                replacements[number]+=1
                launch_worker_tabs(rid,[number])
                launch_deadlines[number]=time.monotonic()+STARTUP_TIMEOUT
        for number,deadline in list(launch_deadlines.items()):
            if any(w["worker"]==number for w in db.workers(rid)):
                del launch_deadlines[number]
            elif time.monotonic()>deadline:
                raise RuntimeError(f"W{number:02} replacement failed to start")
        if time.monotonic()>=next_log:
            log(f"PROGRESS | coverage={db.coverage(run['prep'])} | {db.progress(run['prep'])}")
            next_log=time.monotonic()+10
        if db.all_work_finished(run["prep"]):
            return
        time.sleep(0.5)


def shutdown_workers(db,rid):
    db.stop(rid)
    deadline=time.monotonic()+310
    while time.monotonic()<deadline:
        live=[w for w in db.workers(rid) if process_alive(w["pid"],w["created"])]
        if not live:
            break
        time.sleep(0.5)
    for w in db.workers(rid):
        if not process_alive(w["pid"],w["created"]):
            db.recover_worker(w["owner"])
    # Retain RUNNING state if an in-flight worker survives the deadline.
    return not any(process_alive(w["pid"],w["created"]) for w in db.workers(rid))


@contextmanager
def supervisor_lock():
    OUTPUT.mkdir(parents=True,exist_ok=True)
    with (OUTPUT/"supervisor.lock").open("a+b") as file:
        file.seek(0,os.SEEK_END)
        if file.tell()==0:
            file.write(b"0")
            file.flush()
        file.seek(0)
        if os.name=="nt":
            import msvcrt
            try:
                msvcrt.locking(file.fileno(),msvcrt.LK_NBLCK,1)
            except OSError:
                raise RuntimeError("Another preparation/extraction is active") from None
            try:
                yield
            finally:
                file.seek(0)
                msvcrt.locking(file.fileno(),msvcrt.LK_UNLCK,1)
        else:
            import fcntl
            try:
                fcntl.flock(file.fileno(),fcntl.LOCK_EX|fcntl.LOCK_NB)
            except OSError:
                raise RuntimeError("Another preparation/extraction is active") from None
            try:
                yield
            finally:
                fcntl.flock(file.fileno(),fcntl.LOCK_UN)


def execute_discovery_request(collector_client,profile,payload,timeout,observe,renew=False):
    """Send the prepared SOAP using the Main Collector's requests.Session/auth."""
    import requests
    code=profile["code"]
    started=time.monotonic()

    def ensure(force=False):
        if force or jwt_exp(collector_client.token)<=time.time()+120:
            observe("AUTH REFRESH",stage_timeout=45)
        else:
            observe("REUSING VALID TOKEN",remaining_seconds=int(jwt_exp(collector_client.token)-time.time()))
        try:
            collector_client.ensure_token(code,force=force)
        except RuntimeError as exc:
            # The collector may be running as __main__; don't import a second
            # copy of its exception classes. Expose only safe error categories.
            name=type(exc).__name__
            if name=="OICStopped":
                raise CollectionError("USER_STOP") from None
            if name=="OICSessionExpired":
                raise CollectionError("AUTH_COOKIE_EXPIRED","GenToken session expired",endpoint="GEN_TOKEN") from None
            match=re.search(r"GenToken HTTP (\d{3})",str(exc))
            if match:
                raise classify_http(int(match.group(1)),{},"GEN_TOKEN") from None
            if "SSLError" in str(exc):
                raise CollectionError("TLS_ERROR","GenToken TLS verification failed",endpoint="GEN_TOKEN") from None
            raise CollectionError("AUTH_ENDPOINT_ERROR","Main Collector GenToken failed",5,endpoint="GEN_TOKEN") from None

    def check():
        if collector_client.stop_event.is_set():
            raise CollectionError("USER_STOP")

    auth_start=time.monotonic()
    check()
    ensure(renew)
    auth_seconds=time.monotonic()-auth_start
    for attempt in range(2):
        check()
        headers={"Accept":"text/plain, */*; q=0.01","Content-Type":"text/xml",
                 "Origin":BASE_URL,"Referer":BASE_URL+collector_client.page_path(code),
                 "X-Requested-With":"XMLHttpRequest","Authorization":"Bearer "+collector_client.token}
        sent=time.monotonic()
        observe("REQUEST IN FLIGHT",stage_timeout=timeout,bytes=0)
        retry_auth=False
        cookie_before=collector_client.cookie_header()
        try:
            with collector_client.session.post(XMLA_URL,data=payload,headers=headers,
                    timeout=(15,timeout),stream=True,allow_redirects=False) as response:
                headers_at=time.monotonic()
                observe("HEADERS RECEIVED",http_status=response.status_code)
                if response.status_code in (401,403) and attempt==0:
                    retry_auth=True
                elif response.status_code in (401,403):
                    raise CollectionError("ACCESS_DENIED",f"XMLA HTTP {response.status_code} after JWT refresh",endpoint="XMLA",http_status=response.status_code)
                elif response.status_code!=200:
                    body=bytearray()
                    for chunk in response.iter_content(chunk_size=8192):
                        body.extend(chunk[:65536-len(body)])
                        if len(body)>=65536:
                            break
                    raise classify_http(response.status_code,response.headers,"XMLA",bytes(body))
                else:
                    body=bytearray()
                    for chunk in response.iter_content(chunk_size=256*1024):
                        check()
                        if not chunk:
                            continue
                        if len(body)+len(chunk)>MAX_RESPONSE_BYTES:
                            raise CollectionError("RESPONSE_TOO_LARGE","100 MiB limit")
                        body.extend(chunk)
                        observe("RECEIVING",bytes=len(body))
                    received=time.monotonic()
            if retry_auth:
                auth_start=time.monotonic()
                ensure(True)
                auth_seconds+=time.monotonic()-auth_start
                continue
            # Keep cookie rotations from successful XMLA responses available to
            # the later curl benchmark clients through the existing callback.
            current_cookie=collector_client.cookie_header()
            if current_cookie and current_cookie!=cookie_before and collector_client.on_auth:
                collector_client.on_auth(collector_client.token,current_cookie)
            return bytes(body),{"duration":time.monotonic()-started,"http_status":200,
                "response_bytes":len(body),"token_refresh_seconds":auth_seconds,
                "header_wait_seconds":headers_at-sent,"receive_seconds":received-headers_at}
        except requests.exceptions.SSLError:
            raise CollectionError("TLS_ERROR","XMLA TLS verification failed",endpoint="XMLA") from None
        except requests.exceptions.Timeout:
            raise CollectionError("RESPONSE_TIMEOUT","XMLA requests timeout",endpoint="XMLA") from None
        except requests.exceptions.ConnectionError:
            raise CollectionError("CONNECTION_RESET","XMLA connection failure",endpoint="XMLA") from None
        except requests.RequestException:
            raise CollectionError("NETWORK_ERROR","XMLA requests transport failure",endpoint="XMLA") from None
    raise CollectionError("ACCESS_DENIED","Authentication retry exhausted",endpoint="XMLA")


def load_member_inventory_summary(catalogs):
    """Read the matching saved snapshot without opening a writable DB."""
    from oic_all_dim_prepare import inventory_summary
    if not DB_PATH.is_file():
        return inventory_summary(None,catalogs)
    conn=sqlite3.connect(DB_PATH.resolve().as_uri()+"?mode=ro",uri=True,timeout=2)
    try:
        conn.execute("BEGIN")
        return inventory_summary(conn,catalogs)
    finally:
        conn.close()


def prepare_from_collector(catalogs,collector_client,on_progress,stop_event):
    """Use the existing supervisor and DB lifecycle without starting a new UI."""
    with supervisor_lock():
        db=OICDatabase(DB_PATH)
        try:
            db.initialize_schema()
            recover_old_runs(db)
            return run_prepare(db,catalogs=catalogs,collector_client=collector_client,
                               on_progress=on_progress,stop_event=stop_event)
        finally:
            db.close()


def run_prepare(db,catalogs=None,collector_client=None,on_progress=None,stop_event=None):
    monitor=PrepareMonitor(on_progress=on_progress)
    client=None
    benchmark_client=None
    status="FAILED"
    check_stop=stop_event.is_set if stop_event is not None else lambda:False

    def ensure_running():
        if check_stop():
            raise CollectionError("USER_STOP")

    def emit(event,**data):
        if event!="summary":
            ensure_running()
        monitor.emit(event,**data)
    try:
        monitor.start()
        if collector_client is None:
            client=OICClient(periodic=True,stop=check_stop)
            if not client.cookie:
                raise RuntimeError("Set OIC_COOKIE in local .env")
        elif not collector_client.cookie_header():
            raise RuntimeError("Set OIC_COOKIE in Main Collector")

        health_checks={}
        unknown_streak=0
        log("AUTH | Cookie Jar active | reuse valid JWT | periodic renewal at safe request boundaries")
        log("TRANSPORT | "+("Main Collector requests.Session" if collector_client is not None else "curl_cffi"))
        initial_token=collector_client.token if collector_client is not None else client.token
        log(f"AUTH | Initial JWT remaining {max(0,int(jwt_exp(initial_token)-time.time()))}s")
        last_auth_check=time.monotonic()
        preflight_complete=False

        def monitored_call(profile,payload,timeout,phase,detail,attempt):
            nonlocal last_auth_check
            ensure_running()
            cube=f"{profile['code']}/{profile['cube_type']}"
            key=uuid.uuid4().hex
            monitor.request_started(key,cube,detail,timeout,attempt,phase)
            observer=lambda stage,**data: monitor.request_updated(key,stage,**data)
            if client is not None:
                client.observer=observer
            try:
                if collector_client is None:
                    data,metrics=client.execute_once(profile,payload,timeout)
                else:
                    token_before=collector_client.token
                    renew=time.monotonic()-last_auth_check>=SESSION_RECHECK_SECONDS
                    data,metrics=execute_discovery_request(collector_client,profile,payload,timeout,observer,renew)
                    if renew or collector_client.token!=token_before:
                        last_auth_check=time.monotonic()
                # HTTP 200 can still carry a SOAP fault or session error.
                diagnostic=extract_error_details(data)
                if diagnostic:
                    raise CollectionError(diagnostic,"XMLA error response",5,endpoint="XMLA",http_status=200)
                monitor.request_finished(key,True,metrics)
                return data
            except BaseException as exc:
                monitor.request_finished(key,False,error=str(exc) if isinstance(exc,CollectionError) else type(exc).__name__)
                raise
            finally:
                if client is not None:
                    client.observer=None

        def cooldown(profile,delay,reason):
            cube=f"{profile['code']}/{profile['cube_type']}"
            key=uuid.uuid4().hex
            monitor.request_started(key,cube,reason,delay,kind="RECOVERY")
            monitor.retry(key,delay)
            try:
                wait_seconds(delay,check_stop)
            finally:
                monitor.request_finished(key,True,error="Cooldown ended")

        def check_cube_health(profile):
            from oic_all_dim_prepare import build_discover_request,parse_discover_rows
            cube=f"{profile['code']}/{profile['cube_type']}"
            last=health_checks.get(cube)
            if last is not None and time.monotonic()-last<60:
                return True
            try:
                data=monitored_call(profile,build_discover_request(profile,"MDSCHEMA_HIERARCHIES"),
                                    30,"HEALTH_CHECK","Small hierarchy metadata request",1)
                if not parse_discover_rows(data):
                    raise ValueError("Empty health response")
            except (RuntimeError,ValueError) as exc:
                if getattr(exc,"kind","") in ("AUTH_COOKIE_EXPIRED","AUTH_TOKEN_EXPIRED","ACCESS_DENIED","TLS_ERROR","AUTH_ENDPOINT_ERROR","RATE_LIMITED"):
                    raise
                raise CollectionError("SERVER_UNAVAILABLE","XMLA health check failed",5,endpoint="XMLA") from exc
            health_checks[cube]=time.monotonic()
            return True

        def transport(profile,payload,timeout):
            nonlocal unknown_streak,preflight_complete
            phase,detail=monitor.context()
            for attempt in range(1,MAX_ATTEMPTS+1):
                try:
                    data=monitored_call(profile,payload,timeout,phase,detail,attempt)
                    unknown_streak=0
                    if not preflight_complete:
                        # Reuse the first real discovery response as the connection
                        # check instead of issuing an additional preflight request.
                        from oic_all_dim_prepare import parse_discover_rows
                        if "MDSCHEMA_" in detail:
                            try:
                                rows=parse_discover_rows(data)
                                if rows:
                                    preflight_complete=True
                                    log("CONNECTION CHECK | XMLA HTTP 200 | valid metadata rowset")
                            except ValueError:
                                pass
                    return data
                except CollectionError as exc:
                    kind=exc.kind
                    if kind in ("QUERY_SCOPE_ERROR","QUERY_FAULT","DISCOVER_NOT_SUPPORTED","RESPONSE_TOO_LARGE"):
                        raise
                    if kind in ("HTTP_500_UNKNOWN","RESPONSE_TIMEOUT","XMLA_FAULT_UNKNOWN") and exc.endpoint!="GEN_TOKEN":
                        # Do not repeat the same large request three times.
                        # A responsive small query permits a bounded local fallback,
                        # but is not proof the original error was size-related.
                        unknown_streak+=1
                        cooldown(profile,max(5,exc.delay),"Unclassified XMLA failure; checking endpoint")
                        if unknown_streak>3:
                            raise CollectionError("SERVER_UNAVAILABLE","Repeated unexplained XMLA failures; stopping",5) from exc
                        check_cube_health(profile)
                        raise CollectionError("DISCOVERY_SCOPE_UNCONFIRMED",
                                              f"{kind}; small metadata request succeeded",5,endpoint="XMLA",
                                              http_status=exc.http_status) from exc
                    if kind not in ("RATE_LIMITED","SERVER_UNAVAILABLE","NETWORK_ERROR","DNS_ERROR",
                                    "CONNECTION_RESET","AUTH_ENDPOINT_ERROR","AUTH_TOKEN_EXPIRED") or attempt>=MAX_ATTEMPTS:
                        raise
                    if kind in ("NETWORK_ERROR","DNS_ERROR","CONNECTION_RESET") and client is not None:
                        client.reconnect()
                    minimum=5*(2**attempt) if kind in ("RATE_LIMITED","SERVER_UNAVAILABLE") else 5
                    cooldown(profile,max(2**attempt,exc.delay,minimum),str(exc))
            raise RuntimeError("Discovery retry budget exhausted")

        def probe(req,profile,parallel=False):
            nonlocal benchmark_client
            ensure_running()
            # Concurrent preparation probes own their database and CURL
            # connection. Extraction workers remain separate visible processes.
            probe_db=OICDatabase(DB_PATH) if parallel else db
            probe_client=None
            key=req["id"]
            try:
                if parallel:
                    probe_client=OICClient(stop=check_stop)
                elif client is not None:
                    probe_client=client
                else:
                    if benchmark_client is None:
                        benchmark_client=OICClient(stop=check_stop)
                    probe_client=benchmark_client
                existing=probe_db.get_request(key)
                scope=req["scope"]
                detail=f"{scope['hierarchy'] or 'YEAR ONLY'} | X={len(scope['members'])} Y={len(scope['years'])} M={len(scope['measures'])}"
                if existing["status"] in ("COMMITTED","EMPTY"):
                    monitor.emit("activity",message=f"{req['cube']} | Reusing committed probe {key[:12]}")
                    return {"ok":True,**json.loads(existing["metrics_json"] or "{}")}
                if existing["status"]=="PROBE_FAILED":
                    return {"ok":False}
                monitor.request_started(key,req["cube"],detail,req["timeout"],kind="CONCURRENCY PROBE" if parallel else "SHAPE PROBE")
                probe_client.observer=lambda stage,**data: monitor.request_updated(key,stage,**data)
                claimed=probe_db.claim_probe(key)
                result=execute_resilient(probe_db,probe_client,claimed,probe=True)
                if result.get("pressure"):
                    delay=max(10,result.get("delay",0))
                    monitor.retry(key,delay)
                    wait_seconds(delay,check_stop)
                monitor.request_finished(key,result.get("ok",False),result,error=result.get("error"))
                if result.get("error") in ("USER_STOP","AUTH_COOKIE_EXPIRED","ACCESS_DENIED","TLS_ERROR","DISK_FULL","SOURCE_VERSION_CHANGED","AUTH_ENDPOINT_ERROR"):
                    raise CollectionError(result["error"],"Preparation probe stopped")
                return result
            except BaseException as exc:
                monitor.request_finished(key,False,error=getattr(exc,"kind",type(exc).__name__))
                raise
            finally:
                if probe_client is not None:
                    probe_client.observer=None
                if parallel:
                    if probe_client is not None:
                        probe_client.close()
                    probe_db.close()

        ready=prepare_all(db,transport,probe,OUTPUT,log,emit=emit,catalogs=catalogs)
        status="READY" if ready else "INCOMPLETE"
        return 0 if ready else 1
    except CollectionError as exc:
        if exc.kind!="USER_STOP":
            raise
        status="STOPPED"
        return 130
    except KeyboardInterrupt:
        status="STOPPED"
        raise
    finally:
        try:
            if client is not None:
                client.close()
            if benchmark_client is not None:
                benchmark_client.close()
        finally:
            monitor.stop(status)


def run_extraction(db,workers):
    prep=db.latest_preparation()
    if not prep or prep["status"]!="READY":
        raise RuntimeError("Run --prepare successfully before extraction")
    if db.all_work_finished(prep["id"]):
        progress=db.progress(prep["id"])
        log(f"No pending work | {progress}")
        return int(any(k in progress for k in ("FAILED","NEEDS_REPREPARE")))
    rid=db.start_run(prep["id"],workers,min(workers,prep["concurrency"]),os.getpid(),psutil.Process().create_time())
    configure_log(OUTPUT/"runs"/rid/"run.log")
    code=0
    state="FAILED"
    try:
        launch_worker_tabs(rid,list(range(1,workers+1)))
        wait_workers_ready(db,rid,workers)
        supervise_workers(db,rid,workers)
        progress=db.progress(prep["id"])
        if stopped(db,rid):
            state="STOPPED"
        elif any(k in progress for k in ("FAILED","NEEDS_REPREPARE")):
            state="PARTIAL"
        else:
            state="COMPLETED"
        code=int(state!="COMPLETED")
    except KeyboardInterrupt:
        state="STOPPED"
        code=130
    finally:
        if shutdown_workers(db,rid):
            db.finish_run(rid,state)
        else:
            log("Workers still stopping; resume only after their processes exit")
        log(f"RUN {state} | {db.progress(prep['id'])}")
    return code


def build_parser():
    parser=argparse.ArgumentParser(description=__doc__)
    group=parser.add_mutually_exclusive_group()
    group.add_argument("--prepare",action="store_true")
    group.add_argument("--status",action="store_true")
    parser.add_argument("--workers",type=int,default=DEFAULT_WORKERS)
    parser.add_argument("--worker-id",type=int,help=argparse.SUPPRESS)
    parser.add_argument("--run-id",help=argparse.SUPPRESS)
    return parser


def main():
    args=build_parser().parse_args()
    configure_log()
    if not 1<=args.workers<=16:
        raise RuntimeError("Worker count must be between 1 and 16")
    if args.worker_id is not None or args.run_id is not None:
        if args.worker_id is None or not args.run_id or args.prepare or args.status:
            raise RuntimeError("Invalid internal worker invocation")
        return run_worker(args.run_id,args.worker_id)
    if args.status:
        db=OICDatabase(DB_PATH)
        try:
            db.initialize_schema()
            prep=db.latest_preparation()
            if not prep:
                log("No preparation yet")
            else:
                log(f"PREPARATION {prep['status']} | {prep['id']}")
                print(json.dumps(json.loads(prep["summary_json"]),ensure_ascii=False,indent=2))
                print(json.dumps({"requests":db.progress(prep["id"]),"coverage":db.coverage(prep["id"])},ensure_ascii=False,indent=2))
            return 0
        finally:
            db.close()
    with supervisor_lock():
        db=OICDatabase(DB_PATH)
        try:
            db.initialize_schema()
            recover_old_runs(db)
            return run_prepare(db) if args.prepare else run_extraction(db,args.workers)
        finally:
            db.close()


if __name__=="__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        raise SystemExit(130)
    except (RuntimeError,OSError,sqlite3.Error) as exc:
        LOGGER.error("%s",exc)
        raise SystemExit(1)
