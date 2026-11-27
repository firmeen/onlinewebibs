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
from oic_all_dim_prepare import prepare_all, MD_NS, checked_root, local, build_recovery_requests

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


def log(message,auth=False):
    LOGGER.info(message,extra={"auth_event":auth})


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
        self.auth_state=None
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
                    # Keep auth diagnostics in prepare.log, not the collection activity feed.
                    if getattr(record,"auth_event",False) or message.startswith("AUTH |"):
                        return
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

    def auth_status(self,state,operation):
        with self.lock:
            if self.auth_state==state:
                return
            self.auth_state=state
        if self.on_progress is not None:
            self.on_progress("auth_status",state=state,operation=operation)

    def request_updated(self,key,stage,**data):
        with self.lock:
            row=self.active.get(key)
            if row is None:
                return
            if row["stage"]!=stage:
                row["stage"]=stage
                row["stage_at"]=time.monotonic()
            row.update(data)
        auth_stage=stage in ("AUTH REFRESH","AUTH LOCK WAIT","REUSING VALID TOKEN","REUSING SHARED TOKEN","AUTH RETRY")
        if self.on_progress is not None:
            if stage in ("AUTH REFRESH","AUTH LOCK WAIT","AUTH RETRY"):
                self.auth_status("REFRESHING" if stage!="AUTH RETRY" else "RETRYING",stage)
            elif stage in ("REUSING VALID TOKEN","REUSING SHARED TOKEN","REQUEST IN FLIGHT"):
                with self.lock:
                    waiting=any(v["stage"] in ("AUTH REFRESH","AUTH LOCK WAIT","AUTH RETRY")
                                for v in self.active.values())
                if not waiting:
                    self.auth_status("VALID","SESSION ACTIVE")
        if stage not in ("RECEIVING",):
            suffix=f" | HTTP {data['http_status']}" if "http_status" in data else ""
            if "remaining_seconds" in data:
                suffix+=f" | JWT remaining {data['remaining_seconds']}s"
            log(f"{key[:12]} | {stage}{suffix}",auth=auth_stage)

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
            self.observe("REQUEST IN FLIGHT",stage_timeout=timeout,bytes=0,
                         remaining_seconds=max(0,int(jwt_exp(self.token)-time.time())))
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
                                temporary=destination.with_name(uuid.uuid4().hex[:16]+".partial")
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
                                self.observe("SAVING RAW XML",bytes=size)
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
                raise CollectionError("DISK_FULL" if exc.errno==28 else "STORAGE_ERROR",f"{type(exc).__name__}: {exc}") from None
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
    return OUTPUT/"raw"/r["prep"][:12]/r["id"][:24]/(r["payload_hash"][:16]+"."+r["lease"][:12]+".xml")


def replay_raw_response(r):
    candidates=[]
    if r.get("response_path") and r.get("response_hash"):
        candidates.append((Path(r["response_path"]),r["response_hash"]))
    # Recover a fully renamed response if the previous worker died after the
    # atomic file write but before registering it in the database.
    folder=raw_path(r).parent
    if folder.is_dir():
        candidates.extend((p,None) for p in sorted(folder.glob(r["payload_hash"][:16]+".*.xml")))
    legacy=OUTPUT/"raw"/r["prep"]/r["cube"]/r["id"]
    if legacy.is_dir():
        candidates.extend((p,None) for p in sorted(legacy.glob(r["payload_hash"]+".*.xml")))
    for path,expected_hash in candidates:
        if not path.resolve().is_relative_to(OUTPUT.resolve()) or not path.is_file() or path.stat().st_size>MAX_RESPONSE_BYTES:
            continue
        if not path.name.startswith((r["payload_hash"]+".",r["payload_hash"][:16]+".")):
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


EXCEL_DATA_ROWS = 1_048_575


def export_consolidated_excel(db,prep,rid=None,on_progress=None):
    """Supervisor-only, offline export of the entire committed preparation.

    Rebuild from validated raw XML, including prior runs, so resume is
    idempotent. No worker writes XLSX; split workbooks only at Excel's row limit.
    """
    root=OUTPUT/"excel"/prep[:12]
    staging=root/(".export_"+uuid.uuid4().hex[:12])
    total=db.conn.execute("SELECT count(*) FROM oic_requests WHERE prep=? AND status IN ('COMMITTED','EMPTY')",(prep,)).fetchone()[0]
    done=rows_written=0
    files=[]
    wb=sheet=None
    part_rows=0
    last_report=0.0

    def report(status,error=None):
        nonlocal last_report
        last_report=time.monotonic()
        state={"run_id":rid,"status":status,"requests":done,"total_requests":total,
               "rows":rows_written,"files":len(files),"error":error}
        db.conn.execute("""UPDATE oic_preparations SET summary_json=json_set(
            coalesce(summary_json,'{}'),'$.excel_export',json(?)) WHERE id=?""",
            (json.dumps(state,ensure_ascii=False),prep))
        if rid:
            notify_collection_progress(db,rid,on_progress)

    try:
        report("EXPORTING")
        if not total:
            report("EMPTY")
            return 0
        from openpyxl import Workbook
        from openpyxl.cell import WriteOnlyCell
        root.mkdir(parents=True,exist_ok=True)
        staging.mkdir()
        measures=sorted({m for (raw,) in db.conn.execute("SELECT scope_json FROM oic_jobs WHERE prep=?",(prep,))
                         for m in json.loads(raw)["measures"]})
        headers=["CATEGORY","YEAR_TYPE","CUBE","HIERARCHY","YEAR","YEAR_UNIQUE_NAME",
                 "MEMBER","MEMBER_UNIQUE_NAME","REQUEST_ID","SOURCE_UPDATE",*measures]
        if len(headers)>16384:
            raise RuntimeError("Combined measure columns exceed Excel column limit")
        captions={(cube,hierarchy,name):caption for cube,hierarchy,name,caption in db.conn.execute(
            "SELECT cube,hierarchy,unique_name,caption FROM oic_members WHERE prep=?",(prep,))}

        def value(raw):
            if raw is None:
                return None
            try:
                d=Decimal(raw)
                if d.is_finite() and len(d.as_tuple().digits)<=15 and abs(d)<Decimal("1e15"):
                    if d==d.to_integral_value():
                        return int(d)
                    number=float(d)
                    if number!=0:
                        return number
            except (InvalidOperation,OverflowError,ValueError):
                pass
            return str(raw)

        def append(target,items):
            cells=[]
            for item in items:
                if isinstance(item,str):
                    cell=WriteOnlyCell(target,value=item)
                    cell.data_type="s"
                    cells.append(cell)
                else:
                    cells.append(item)
            target.append(cells)

        def begin_part():
            nonlocal wb,sheet,part_rows
            wb=Workbook(write_only=True)
            sheet=wb.create_sheet("DATA")
            append(sheet,headers)
            part_rows=0

        def save_part():
            nonlocal wb,sheet
            path=staging/f"OIC_DATA_{len(files)+1:03d}.xlsx"
            info=wb.create_sheet("INFO")
            for key,item in (("PREPARATION",prep),("PART",len(files)+1),("ROWS",part_rows),
                             ("COMMITTED_REQUESTS",total),("DATA_TYPE","OLAP AGGREGATE"),
                             ("NOTICE","Do not sum All members with detail members"),
                             ("SCOPE","Committed requests only; may be partial collection")):
                append(info,[key,item])
            wb.save(path)
            wb.close()
            wb=sheet=None
            files.append(path)

        requests=db.conn.execute("""SELECT r.id,r.cube,r.scope_json,r.response_path,r.response_hash
            FROM oic_requests r JOIN oic_jobs j ON j.prep=r.prep AND j.id=r.job
            WHERE r.prep=? AND r.status IN ('COMMITTED','EMPTY')
            ORDER BY json_extract(j.scope_json,'$.job_index'),
                     json_extract(r.scope_json,'$.year_offset'),
                     json_extract(r.scope_json,'$.member_offset'),r.id""",(prep,))
        for r in requests:
            scope=json.loads(r["scope_json"])
            raw=Path(r["response_path"]) if r["response_path"] else None
            if (not raw or not raw.resolve().is_relative_to(OUTPUT.resolve())
                    or not raw.is_file() or not r["response_hash"]):
                raise RuntimeError(f"Committed raw XML unavailable: {r['id'][:12]}")
            data=raw.read_bytes()
            if digest(data)!=r["response_hash"]:
                raise RuntimeError(f"Raw response hash mismatch: {r['id'][:12]}")
            years,axes,cells,updated=parse_xmla(data,scope)
            grouped={}
            for idx,(member,measure) in enumerate(axes):
                for yi,year in enumerate(years):
                    grouped.setdefault((year,member),{})[measure]=value(cells.get(idx*len(years)+yi,(None,None))[0])
            category,kind=r["cube"].split("/")
            for (year,member),values in grouped.items():
                if wb is None:
                    begin_part()
                elif part_rows>=EXCEL_DATA_ROWS:
                    save_part()
                    begin_part()
                append(sheet,[category,kind,r["cube"],scope["hierarchy"],
                    captions.get((r["cube"],scope["year_hierarchy"],year)) or year,year,
                    captions.get((r["cube"],scope["hierarchy"],member)) or member,member,
                    r["id"],updated,*[values.get(m) for m in measures]])
                part_rows+=1
                rows_written+=1
            done+=1
            if time.monotonic()-last_report>=1:
                report("EXPORTING")
        if wb is not None:
            save_part()
        if done!=total:
            raise RuntimeError("Committed request count changed during export")
        # Only publish after every raw response has been validated and written.
        for path in files:
            os.replace(path,root/path.name)
        # A rebuilt snapshot may have fewer parts; remove only our own old files.
        names={path.name for path in files}
        for path in root.glob("OIC_DATA_*.xlsx"):
            if re.fullmatch(r"OIC_DATA_\d{3,}\.xlsx",path.name) and path.name not in names:
                path.unlink()
        report("SAVED")
        log(f"EXCEL SAVED | {len(files)} file(s) | {rows_written:,} rows | {root}")
        return 0
    except (Exception,KeyboardInterrupt) as exc:
        error=f"{type(exc).__name__}: {exc}"[:300]
        log(f"EXCEL ERROR | {error}")
        report("FAILED",error)
        return 1
    finally:
        if wb is not None:
            try:
                if sheet is not None:
                    sheet.close()
                wb.close()
            except Exception:
                pass
        shutil.rmtree(staging,ignore_errors=True)


def process_request(db,client,r,rid=None,owner=None):
    scope=json.loads(r["scope_json"])
    profile=db.profile(r["prep"],r["cube"])
    if digest(r["payload"])!=r["payload_hash"]:
        raise CollectionError("INVALID_QUERY","Prepared payload hash mismatch")
    started=time.monotonic()
    client.observe("CHECKING RAW RESPONSE")
    replay=replay_raw_response(r)
    if replay:
        client.observe("REPLAYING SAVED XML")
        path,parsed,sha=replay
        metrics=json.loads(r.get("metrics_json") or "{}")
        metrics["replayed"]=True
        db.register_response(r,path,sha,metrics)
    else:
        path=raw_path(r)
        if rid:
            client.observe("WAITING FOR HTTP SLOT")
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


def handle_request_failure(db,client,r,exc,rid,probe=False,late_reset=False):
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
    elif kind in ("RESPONSE_TOO_LARGE","SOAP_FAULT","QUERY_FAULT","QUERY_SCOPE_ERROR","RESPONSE_TIMEOUT","INVALID_QUERY") or late_reset:
        scope=json.loads(r["scope_json"])
        children=build_recovery_requests({**r,"scope":scope},db.profile(r["prep"],r["cube"]))
        if not children or not db.activate_recovery_requests(r,message,children):
            db.fail(r,message,"FAILED")
        else:
            log(f"{r['cube']} | {r['id'][:12]} | SPLIT X={len(scope['members'])} -> {len(scope['members'])//2} + {len(scope['members'])-len(scope['members'])//2} | Y/M unchanged")
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
    started=time.monotonic()
    try:
        metrics=process_request(db,client,r,rid,owner)
        return {"ok":True,**metrics}
    except (CollectionError,sqlite3.Error,OSError,ValueError) as exc:
        late_reset=(isinstance(exc,CollectionError) and exc.kind=="CONNECTION_RESET"
                    and time.monotonic()-started>=45 and r["attempts"]>=2)
        if late_reset:
            client.reconnect()
            log("Repeated late connection reset; trying smaller member scope")
        handle_request_failure(db,client,r,exc,rid,probe,late_reset=late_reset)
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


def elapsed_label(seconds):
    seconds=max(0,int(seconds))
    return f"{seconds//3600:02}:{seconds//60%60:02}:{seconds%60:02}"


def collection_snapshot(conn,rid):
    """Small read-only aggregates; replaced parents never count toward coverage."""
    run=conn.execute("SELECT * FROM oic_runs WHERE id=?",(rid,)).fetchone()
    if run is None:
        return None
    run=dict(run)
    prep=run["prep"]
    groups=conn.execute("""SELECT cube,job,status,count(*) AS n,
        coalesce(sum(potential_cells),0) AS cells,
        coalesce(sum(attempts),0) AS attempts
        FROM oic_requests WHERE prep=? GROUP BY cube,job,status""",(prep,)).fetchall()
    cubes={}
    jobs={}
    totals={"planned":0,"covered":0,"pending":0,"running":0,"committed":0,
            "failed":0,"splits":0,"attempts":0}
    for row in groups:
        cube,job,status,n,cells,attempts=row
        totals["attempts"]+=attempts
        if status=="REPLACED":
            totals["splits"]+=n
        if status in ("STANDBY","REPLACED","PROBE","PROBE_FAILED"):
            continue
        c=cubes.setdefault(cube,{"planned":0,"covered":0})
        j=jobs.setdefault(job,{"planned":0,"covered":0,"failed":False})
        for item in (c,j,totals):
            item["planned"]+=cells
            if status in ("COMMITTED","EMPTY"):
                item["covered"]+=cells
        if status in ("COMMITTED","EMPTY"):
            totals["committed"]+=n
        elif status in ("FAILED","NEEDS_REPREPARE"):
            totals["failed"]+=n
            j["failed"]=True
        elif status=="RUNNING":
            totals["running"]+=n
        else:
            totals["pending"]+=n
    totals["jobs"]=conn.execute("SELECT count(*) FROM oic_jobs WHERE prep=?",(prep,)).fetchone()[0]
    totals["jobs_done"]=sum(j["planned"]>0 and j["covered"]==j["planned"] and not j["failed"] for j in jobs.values())
    totals["jobs_failed"]=sum(j["failed"] for j in jobs.values())
    assignments={}
    for job,worker in conn.execute("SELECT id,json_extract(scope_json,'$.worker_id') FROM oic_jobs WHERE prep=?",(prep,)):
        if worker is None:
            continue
        item=assignments.setdefault(worker,{"assigned":0,"done":0,"failed":0})
        item["assigned"]+=1
        j=jobs.get(job,{})
        item["done"]+=int(bool(j.get("planned") and j.get("covered")==j.get("planned") and not j.get("failed")))
        item["failed"]+=int(bool(j.get("failed")))
    metrics=conn.execute("""SELECT coalesce(sum(CASE WHEN json_valid(metrics_json)
            THEN json_extract(metrics_json,'$.response_bytes') ELSE 0 END),0),
            coalesce(sum(CASE WHEN committed_at>=? THEN potential_cells ELSE 0 END),0)
        FROM oic_requests WHERE prep=? AND status IN ('COMMITTED','EMPTY')""",
        (run["created_at"],prep)).fetchone()
    totals["committed_bytes"]=metrics[0] or 0
    totals["run_cells"]=metrics[1] or 0
    row=conn.execute("SELECT summary_json FROM oic_preparations WHERE id=?",(prep,)).fetchone()
    export=json.loads(row[0] or "{}").get("excel_export",{}) if row else {}
    if export.get("run_id")!=rid:
        export={"status":"DEFERRED","requests":0,"total_requests":totals["committed"],"rows":0,"files":0}
    totals["excel_export"]=export
    workers=[dict(row) for row in conn.execute("SELECT * FROM oic_workers WHERE run=? ORDER BY worker",(rid,))]
    return {"run":run,"totals":totals,"cubes":cubes,"workers":workers,"assignments":assignments}


def worker_state_path(rid,worker):
    return OUTPUT/"runs"/rid/f"worker_{worker:02}.state.json"


def read_worker_state(rid,worker):
    try:
        state=json.loads(worker_state_path(rid,worker).read_text(encoding="utf-8"))
        return state if isinstance(state,dict) else {}
    except (OSError,ValueError):
        return {}


def rich_progress(done,total):
    from rich.progress_bar import ProgressBar
    return ProgressBar(total=max(1,total),completed=done,complete_style="green",finished_style="green")


class WorkerMonitor:
    """Display-only state. Publishing failure must never fail an extraction."""
    def __init__(self,rid,worker,owner,count):
        self.rid,self.worker,self.owner,self.count=rid,worker,owner,count
        self.lock=threading.RLock()
        self.started=time.time()
        self.last_publish=0
        self.state={"owner":owner,"worker":worker,"stage":"STARTING","started":self.started,
                    "stage_at":self.started,"request":{},"stats":{"committed":0,"failed":0,"splits":0,"retries":0,
                    "cells":0,"bytes":0,"seconds":0.0},"recent":[],"events":[]}
        self.live=None
        self.handler=None

    def publish(self,force=False):
        now=time.monotonic()
        if not force and now-self.last_publish<0.5:
            return
        self.last_publish=now
        self.state["updated_at"]=time.time()
        path=worker_state_path(self.rid,self.worker)
        tmp=path.with_name(f"w{self.worker:02}.{self.owner[:8]}.tmp")
        try:
            path.parent.mkdir(parents=True,exist_ok=True)
            tmp.write_text(json.dumps(self.state,ensure_ascii=False),encoding="utf-8")
            tmp.replace(path)
        except OSError:
            pass
        finally:
            try:
                tmp.unlink(missing_ok=True)
            except OSError:
                pass

    def observe(self,stage,**data):
        with self.lock:
            changed=stage!=self.state["stage"]
            if changed:
                self.state.update(stage=stage,stage_at=time.time())
            # Do not serialize credentials or arbitrary response content.
            for k in ("stage_timeout","bytes","http_status","remaining_seconds"):
                if k in data:
                    self.state[k]=data[k]
            if "remaining_seconds" in data:
                self.state["jwt_expires_at"]=time.time()+data["remaining_seconds"]
            if stage=="REQUEST IN FLIGHT":
                self.state.pop("http_status",None)
                self.state.pop("http_finished",None)
                self.state["http_started"]=time.time()
                self.state["http_attempts"]=self.state.get("http_attempts",0)+1
            if stage in ("SAVING RAW XML","PARSING","COMMITTING") and self.state.get("http_started") and not self.state.get("http_finished"):
                self.state["http_finished"]=time.time()
            if self.state["request"] and changed:
                self.state["request"].setdefault("stages",[]).append(stage)
            self.publish(force=changed)

    def event(self,message):
        with self.lock:
            self.state["events"]=(self.state["events"]+[{"time":time.time(),"text":str(message)[:240]}])[-6:]
            self.publish()

    def begin(self,r,caption):
        scope=json.loads(r["scope_json"])
        with self.lock:
            self.state["request"]={"id":r["id"],"cube":r["cube"],"hierarchy":scope["hierarchy"],
                "caption":caption,"x":len(scope["members"]) if scope["hierarchy"] else 0,
                "y":len(scope["years"]),"m":len(scope["measures"]),"cells":r["potential_cells"],
                "attempt":r["attempts"],"timeout":r["timeout"],"parent":r["parent"],
                "started":time.time(),"job_index":scope.get("job_index",0),
                "year_offset":scope.get("year_offset",0),"member_offset":scope.get("member_offset",0),
                "request_index":scope.get("request_index",0),"request_total":scope.get("request_total",0),
                "total_members":scope.get("total_members",0),"total_years":scope.get("total_years",0)}
            for k in ("http_status","http_started","http_finished","stage_timeout"):
                self.state.pop(k,None)
            self.state["bytes"]=0
            self.observe("LOADING PAYLOAD")

    def finish(self,result,status):
        with self.lock:
            q=self.state["request"]
            stats=self.state["stats"]
            duration=max(0,time.time()-q.get("started",time.time()))
            label={"COMMITTED":"SAVED","EMPTY":"EMPTY","REPLACED":"SPLIT","RETRY":"RETRY"}.get(status,status)
            if status in ("COMMITTED","EMPTY"):
                stats["committed"]+=1
                stats["cells"]+=q["cells"]
                stats["bytes"]+=0 if result.get("replayed") else int(result.get("response_bytes",0) or 0)
                stats["seconds"]+=duration
            elif status=="REPLACED":
                stats["splits"]+=1
            elif status=="RETRY":
                stats["retries"]+=1
            else:
                stats["failed"]+=1
            record={**q,"result":label,"duration":duration,"error":result.get("error",""),"time":time.time()}
            self.state["recent"]=(self.state["recent"]+[record])[-6:]
            self.event(f"{q['cube']} | {q['caption']} | {label} | {duration:.1f}s"+(f" | {record['error']}" if record["error"] else ""))
            self.state["request"]={}
            self.observe(label)
            self.publish(True)

    def start(self):
        from rich.console import Console
        from rich.live import Live
        console=Console()
        monitor=self
        class Activity(logging.Handler):
            def emit(self,record):
                if not getattr(record,"auth_event",False) and not record.getMessage().startswith("AUTH |"):
                    monitor.event(record.getMessage())
        self.handler=Activity()
        LOGGER.addHandler(self.handler)
        self.publish(True)
        if console.is_terminal:
            self.live=Live(console=console,get_renderable=self.render,refresh_per_second=2,
                           screen=True,vertical_overflow="crop")
            try:
                self.live.start(refresh=True)
            except Exception as exc:
                self.live=None
                LOGGER.warning("Rich worker display unavailable: %s",type(exc).__name__)
                console_handler=logging.StreamHandler(sys.stdout)
                console_handler.setFormatter(logging.Formatter("%(asctime)s | %(message)s",datefmt="%H:%M:%S"))
                LOGGER.addHandler(console_handler)

    def stop(self,final):
        self.observe(final)
        self.publish(True)
        if self.live is not None:
            try:
                self.live.stop()
                self.live.console.print(self.render())
            except Exception as exc:
                LOGGER.warning("Rich final display unavailable: %s",type(exc).__name__)
        if self.handler:
            LOGGER.removeHandler(self.handler)
            self.handler.close()

    def render(self):
        from rich.console import Group
        from rich.panel import Panel
        from rich.table import Table
        from rich.text import Text
        from copy import deepcopy
        with self.lock:
            s=deepcopy(self.state)
        q=s["request"]
        stats=s["stats"]
        header=Text(f"OIC COLLECTION  |  WORKER {self.worker:02}/{self.count:02}\n"
                    f"{s['stage']}  |  Elapsed {elapsed_label(time.time()-self.started)}  |  curl_cffi / XMLA",
                    style="bold cyan")
        table=Table.grid(padding=(0,2),expand=True)
        table.add_column(ratio=1);table.add_column(ratio=1)
        table.add_row(Text(f"Cube: {q.get('cube','-')}  |  {q.get('caption','Waiting for work')}"),
                      Text(f"Job {q.get('job_index',0)} | Page {q.get('request_index',0)}/{q.get('request_total',0)} | Attempt {q.get('attempt','-')}"))
        table.add_row(Text(q.get("hierarchy") or "Year totals / idle"),Text(f"Year offset {q.get('year_offset',0)} | Member offset {q.get('member_offset',0)}" + (" | SPLIT" if q.get("parent") else "")))
        table.add_row(Text(f"X={q.get('x',0):,}  Y={q.get('y',0):,}  M={q.get('m',0):,}"),
                      Text(f"Expected cells: {q.get('cells',0):,}  |  Timeout {q.get('timeout','-')}s"))
        stage=Text(f"{s['stage']}  |  Stage elapsed {elapsed_label(time.time()-s['stage_at'])}"
                   f"  |  HTTP {s.get('http_status','-')}  |  Received {s.get('bytes',0)/1048576:.2f} MiB")
        http_elapsed=s.get("http_finished",time.time())-s["http_started"] if q and s.get("http_started") else 0
        budget=q.get("timeout",0)
        visited=set(q.get("stages",[]))
        phases=[("Payload",{"LOADING PAYLOAD","CHECKING RAW RESPONSE"}),
                ("Auth",{"AUTH LOCK WAIT","AUTH REFRESH","REUSING VALID TOKEN","REUSING SHARED TOKEN"}),
                ("HTTP",{"REQUEST IN FLIGHT","HEADERS RECEIVED"}),
                ("XML",{"RECEIVING","SAVING RAW XML","REPLAYING SAVED XML"}),
                ("Parse",{"PARSING"}),("Commit",{"COMMITTING"})]
        flow=Text()
        for label,stages in phases:
            active=s["stage"] in stages
            flow.append(label+(" *" if active else " + " if visited.intersection(stages) else " - "),
                        style="bold yellow" if active else "green" if visited.intersection(stages) else "dim")
            flow.append("  ")
        jwt=f"JWT remaining {max(0,int(s['jwt_expires_at']-time.time()))}s" if s.get("jwt_expires_at") else "JWT expiry not yet observed"
        lifecycle=Group(stage,flow,Text(jwt+" | Auth activity shown in stage"),
                        Text(f"HTTP elapsed {elapsed_label(http_elapsed)} / configured {budget}s (time, not download progress)"),
                        rich_progress(min(http_elapsed,budget),budget))
        recent=Table(expand=True)
        for name in ("Cube / Dimension","Cells","Time","Result"):
            recent.add_column(name,overflow="fold")
        for item in reversed(s["recent"]):
            recent.add_row(Text(f"{item['cube']} / {item['caption']}"),f"{item['cells']:,}",
                           f"{item['duration']:.1f}s",Text(item["result"],style="green" if item["result"] in ("SAVED","EMPTY") else "yellow"))
        activity=Group(*(Text(time.strftime("%H:%M:%S",time.localtime(e["time"]))+"  "+e["text"]) for e in s["events"][-4:]))
        avg=stats["seconds"]/max(1,stats["committed"])
        return Group(Panel(header,border_style="cyan"),Panel(table,title="01 / CURRENT REQUEST"),
                     Panel(lifecycle,title="02 / REQUEST LIFECYCLE"),
                     Panel(Text(f"Saved {stats['committed']}  |  Failed {stats['failed']}  |  Splits {stats['splits']}  |  Retries {stats['retries']}\n"
                                f"Cells {stats['cells']:,}  |  Received {stats['bytes']/1048576:.2f} MiB  |  Avg saved request {avg:.1f}s"),title="03 / WORKER STATISTICS (this process)"),
                     Panel(recent,title="04 / RECENT REQUESTS"),Panel(activity,title="05 / ACTIVITY"))


class CollectionMonitor:
    def __init__(self,rid):
        self.rid=rid
        self.snapshot=None
        self.states={}
        self.error=""

    def render(self):
        from rich.console import Group
        from rich.panel import Panel
        from rich.table import Table
        from rich.text import Text
        if not self.snapshot:
            return Panel(Text(self.error or "Waiting for run snapshot"),title="OIC LIVE MONITOR")
        run=self.snapshot["run"]
        t=self.snapshot["totals"]
        duration=max(1,time.time()-run["created_at"])
        pct=100*t["covered"]/max(1,t["planned"])
        status="STOPPING" if run["state"]=="RUNNING" and run["stop"] else "STARTING" if run["state"]=="RUNNING" and not run["barrier"] else run["state"]
        export=t["excel_export"]
        if export["status"]=="EXPORTING":
            status="EXPORTING"
        header=Text(f"OIC COLLECTION  |  LIVE MONITOR\n{status}  |  Run {self.rid[:12]}  |  Elapsed {elapsed_label(duration)}",
                    style="bold cyan")
        global_panel=Group(Text(f"Jobs {t['jobs_done']}/{t['jobs']}  |  Failed jobs {t['jobs_failed']}  |  Splits {t['splits']}"),
            Text(f"Cells {t['covered']:,}/{t['planned']:,}  |  {pct:.2f}%"),rich_progress(t["covered"],t["planned"]),
            Text(f"Saved requests {t['committed']}  |  Pending {t['pending']}  |  Running {t['running']}  |  Failed {t['failed']}"),
            Text(f"Request attempts {t['attempts']} (plan lifetime)  |  {t['run_cells']/duration:,.0f} cells/s (this run)  |  Saved XML {t['committed_bytes']/1048576:.2f} MiB"),
            Text("Excel: deferred until collection stops" if export["status"]=="DEFERRED" else
                 f"Excel: {export['status']} | Requests {export['requests']}/{export['total_requests']} | Rows {export['rows']:,} | Files {export['files']}"+
                 (f" | {export['error']}" if export.get("error") else "")))
        workers=Table(expand=True)
        for col in ("Worker","Jobs","Cube / Dimension","Stage","Elapsed","Saved","Cells"):
            workers.add_column(col,overflow="fold")
        events=[]
        registered={w["worker"]:w for w in self.snapshot["workers"]}
        for n in range(1,run["workers"]+1):
            w=registered.get(n,{})
            s=self.states.get(n,{})
            q=s.get("request",{})
            valid=bool(w and s.get("owner")==w.get("owner"))
            if not valid and run["state"]=="RUNNING":
                s={};q={}
            stage=s.get("stage",w.get("status","STARTING"))
            if w and w.get("status")!="READY":
                stage=w["status"]
            elif w and time.time()-w["heartbeat"]>15:
                stage="HEARTBEAT STALE"
            stats=s.get("stats",{})
            assigned=self.snapshot["assignments"].get(n,{})
            workers.add_row(f"W{n:02}",f"{assigned.get('done',0)}/{assigned.get('assigned',0)}",Text(f"{q.get('cube','-')} / {q.get('caption','-')}"),
                            Text(stage),elapsed_label(time.time()-q.get("started",time.time())),
                            str(stats.get("committed",0)),f"{stats.get('cells',0):,}")
            events.extend((e["time"],n,e["text"]) for e in s.get("events",[]))
        grid=Table(expand=True)
        for col in ("Category","FY","UY","AY"):
            grid.add_column(col)
        for c in ("AH","FI","MR","MI","MV"):
            values=[]
            for y in ("FY","UY","AY"):
                v=self.snapshot["cubes"].get(c+"/"+y)
                values.append(f"{100*v['covered']/max(1,v['planned']):.1f}%" if v else "-")
            grid.add_row(c,*values)
        activity=Group(*(Text(f"{time.strftime('%H:%M:%S',time.localtime(ts))}  W{n:02}  {msg}") for ts,n,msg in sorted(events,reverse=True)[:5]))
        jwt_known=[s["jwt_expires_at"] for n,s in self.states.items() if s.get("jwt_expires_at") and n in registered and s.get("owner")==registered[n]["owner"]]
        auth_note=f"JWT min remaining {max(0,int(min(jwt_known)-time.time()))}s" if jwt_known else "JWT expiry not observed"
        footer=Text(f"Registered {len(registered)}/{run['workers']}  |  HTTP slots {sum(w['slot'] for w in registered.values())}/{run['allowed']}  |  {auth_note}  |  Read-only SQLite / 1s refresh"
                    +(f"  |  {self.error}" if self.error else ""))
        return Group(Panel(header,border_style="cyan"),Panel(global_panel,title="01 / GLOBAL PROGRESS"),
                     Panel(workers,title="02 / ACTIVE WORKERS"),Panel(grid,title="03 / CUBE COVERAGE"),
                     Panel(activity,title="04 / RECENT ACTIVITY"),footer)

    def run(self):
        from rich.console import Console
        from rich.live import Live
        console=Console()
        conn=sqlite3.connect(DB_PATH.resolve().as_uri()+"?mode=ro",uri=True,timeout=2)
        conn.row_factory=sqlite3.Row
        try:
            with Live(console=console,screen=console.is_terminal,auto_refresh=False,vertical_overflow="crop") as live:
                while True:
                    try:
                        try:
                            with conn:
                                conn.execute("BEGIN")
                                self.snapshot=collection_snapshot(conn,self.rid)
                            if self.snapshot is None:
                                raise RuntimeError("Run not found")
                            self.states={n:read_worker_state(self.rid,n) for n in range(1,self.snapshot["run"]["workers"]+1)}
                            self.error=""
                        except sqlite3.Error as exc:
                            self.error=f"Snapshot unavailable: {type(exc).__name__}"
                        live.update(self.render(),refresh=True)
                        if self.snapshot and self.snapshot["run"]["state"]!="RUNNING":
                            break
                        time.sleep(1)
                    except KeyboardInterrupt:
                        # Ctrl+C anywhere in the monitor loop requests stop,
                        # without interrupting the supervisor's final export.
                        control=OICDatabase(DB_PATH)
                        try:
                            control.stop(self.rid)
                        finally:
                            control.close()
            console.print(self.render())
            return 0
        finally:
            conn.close()


def run_worker(rid,worker):
    from rich.console import Console
    configure_log(OUTPUT/"runs"/rid/f"worker_{worker:02}.log",console_output=not Console().is_terminal)
    db=OICDatabase(DB_PATH)
    owner=uuid.uuid4().hex
    done=threading.Event()
    client=None
    thread=None
    registered=False
    monitor=None
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
        monitor=WorkerMonitor(rid,worker,owner,run["workers"])
        monitor.start()
        client.observer=monitor.observe
        captions={}
        for row in db.conn.execute("SELECT cube,field_json FROM oic_fields WHERE prep=?",(run["prep"],)):
            field=json.loads(row[1])
            if field.get("type")=="DIMENSION" and field.get("mdx"):
                captions.setdefault((row[0],field["mdx"]),field.get("caption") or field["mdx"])
        monitor.observe("WAITING FOR BARRIER")
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
                monitor.observe("WAITING FOR WORK")
                if db.all_work_finished(run["prep"]):
                    final="COMPLETED"
                    return 0
                time.sleep(0.5)
                continue
            scope=json.loads(r["scope_json"])
            monitor.begin(r,captions.get((r["cube"],scope["hierarchy"]),scope["hierarchy"] or "YEAR TOTALS"))
            result=execute_resilient(db,client,r,rid,owner)
            saved=db.get_request(r["id"])
            monitor.finish(result,saved["status"] if saved else "UNKNOWN")
            delay=json.loads(r["scope_json"]).get("delay",1.0)
            try:
                wait_seconds(delay,lambda:stopped(db,rid))
            except CollectionError:
                final="STOPPED"
                return 0
            if result.get("error")=="AUTH_COOKIE_EXPIRED":
                monitor.observe("WAITING FOR CREDENTIAL")
                log("WAITING FOR CREDENTIAL | update local .env")
                while not stopped(db,rid):
                    if client.reload_credentials():
                        break
                    time.sleep(2)
        final="STOPPED"
        return 0
    except KeyboardInterrupt:
        db.stop(rid)
        final="STOPPED"
        return 130
    finally:
        done.set()
        if thread:
            thread.join(timeout=5)
        if client:
            client.close()
        if monitor is not None:
            try:
                monitor.stop(final)
            except Exception as exc:
                LOGGER.warning("Worker monitor cleanup unavailable: %s",type(exc).__name__)
        if registered:
            # Supervisor fences and requeues remaining work if this process dies.
            db.finish_worker(owner,final)
        db.close()


def powershell_quote(value):
    return "'"+str(value).replace("'","''")+"'"


def worker_command(rid,worker=None):
    mode=" --monitor" if worker is None else " --worker-id "+str(worker)
    command=("& "+powershell_quote(sys.executable)+" "+powershell_quote(__file__)+
             mode+" --run-id "+powershell_quote(rid)+"; $code=$LASTEXITCODE; Read-Host 'Press Enter to close this tab'; exit $code")
    return base64.b64encode(command.encode("utf-16-le")).decode("ascii")


def launch_worker_tabs(rid,workers,include_monitor=False):
    wt=shutil.which("wt.exe") or shutil.which("wt")
    ps=shutil.which("pwsh.exe") or shutil.which("powershell.exe")
    if os.name!="nt" or not wt or not ps:
        raise RuntimeError("Visible Windows Terminal and PowerShell are required; no hidden fallback")
    args=[wt,"-w","OIC-"+rid[:12]]
    tabs=([None] if include_monitor else [])+list(workers)
    for index,worker in enumerate(tabs):
        if index:
            args.append(";")
        args.extend(["new-tab","--startingDirectory",str(PROJECT_ROOT),"--suppressApplicationTitle","--title",("LIVE MONITOR" if worker is None else f"OIC W{worker:02}"),
                     ps,"-NoLogo","-NoProfile","-EncodedCommand",worker_command(rid,worker)])
    if include_monitor:
        args.extend([";","focus-tab","--target","0"])
    result=subprocess.run(args,check=False,timeout=30)
    if result.returncode:
        raise RuntimeError(f"Windows Terminal launcher exited {result.returncode}")


def wait_workers_ready(db,rid,count,stop_event=None,on_progress=None):
    end=time.monotonic()+STARTUP_TIMEOUT
    next_progress=0
    while time.monotonic()<end:
        if stop_event is not None and stop_event.is_set():
            db.stop(rid)
            return
        if on_progress is not None and time.monotonic()>=next_progress:
            notify_collection_progress(db,rid,on_progress)
            next_progress=time.monotonic()+1
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


def supervise_workers(db,rid,count,stop_event=None,on_progress=None):
    run=db.run(rid)
    replacements={i:0 for i in range(1,count+1)}
    launch_deadlines={}
    next_log=0
    next_progress=0
    while not stopped(db,rid):
        if stop_event is not None and stop_event.is_set():
            db.stop(rid)
            return
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
        if on_progress is not None and time.monotonic()>=next_progress:
            notify_collection_progress(db,rid,on_progress)
            next_progress=time.monotonic()+1
        if time.monotonic()>=next_log:
            log(f"PROGRESS | coverage={db.coverage(run['prep'])} | {db.progress(run['prep'])}")
            next_log=time.monotonic()+10
        if db.all_work_finished(run["prep"]):
            return
        time.sleep(0.5)


def shutdown_workers(db,rid,on_progress=None):
    db.stop(rid)
    deadline=time.monotonic()+630
    next_progress=0
    while time.monotonic()<deadline:
        if on_progress is not None and time.monotonic()>=next_progress:
            notify_collection_progress(db,rid,on_progress)
            next_progress=time.monotonic()+1
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


def execute_discovery_request(collector_client,profile,payload,timeout,observe):
    """Send prepared SOAP using Main Collector curl_cffi session/auth."""
    code=profile["code"]
    started=time.monotonic()

    def ensure(force=False):
        if force or jwt_exp(collector_client.token)<=time.time()+120:
            observe("AUTH REFRESH",stage_timeout=45)
        else:
            observe("REUSING VALID TOKEN",remaining_seconds=int(jwt_exp(collector_client.token)-time.time()))
        try:
            collector_client.ensure_token(code,force=force)
        except CollectionError:
            raise
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
    ensure()
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
            with collector_client.session.stream("POST",XMLA_URL,data=payload,headers=headers,
                    timeout=(15,timeout),allow_redirects=False) as response:
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
        except RequestException as exc:
            raise classify_network(exc) from None
    raise CollectionError("ACCESS_DENIED","Authentication retry exhausted",endpoint="XMLA")


def load_member_inventory_summary(catalogs):
    """Read the matching saved snapshot without opening a writable DB."""
    from oic_all_dim_prepare import inventory_summary
    if not DB_PATH.is_file():
        return inventory_summary(None,catalogs)
    conn=sqlite3.connect(DB_PATH.resolve().as_uri()+"?mode=ro",uri=True,timeout=2)
    conn.row_factory=sqlite3.Row
    try:
        conn.execute("BEGIN")
        return inventory_summary(conn,catalogs)
    finally:
        conn.close()


def prepare_from_collector(catalogs,collector_client,on_progress,stop_event,mode="all"):
    """Use the existing supervisor and DB lifecycle without starting a new UI."""
    with supervisor_lock():
        db=OICDatabase(DB_PATH)
        try:
            db.initialize_schema()
            recover_old_runs(db)
            return run_prepare(db,catalogs=catalogs,collector_client=collector_client,
                               on_progress=on_progress,stop_event=stop_event,mode=mode)
        finally:
            db.close()


def run_prepare(db,catalogs=None,collector_client=None,on_progress=None,stop_event=None,mode="all",settings=None):
    monitor=PrepareMonitor(on_progress=on_progress)
    client=None
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
        if mode=="plan":
            ready=prepare_all(db,None,None,OUTPUT,log,emit=emit,catalogs=catalogs,mode=mode,settings=settings)
            status="READY" if ready else "INCOMPLETE"
            return 0 if ready else 1
        if collector_client is None:
            client=OICClient(periodic=False,stop=check_stop)
            if not client.cookie:
                raise RuntimeError("Set OIC_COOKIE in local .env")
        elif not collector_client.cookie_header():
            raise RuntimeError("Set OIC_COOKIE in Main Collector")

        health_checks={}
        unknown_streak=0
        log("AUTH | Cookie Jar active | renewal on demand at safe request boundaries",auth=True)
        log("TRANSPORT | "+"curl_cffi")
        initial_token=collector_client.token if collector_client is not None else client.token
        log(f"AUTH | Initial JWT remaining {max(0,int(jwt_exp(initial_token)-time.time()))}s")
        preflight_complete=False

        def monitored_call(profile,payload,timeout,phase,detail,attempt):
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
                    data,metrics=execute_discovery_request(collector_client,profile,payload,timeout,observer)
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
                    if exc.endpoint=="LOCAL_CREDENTIALS":
                        monitor.auth_status("ERROR",kind)
                        raise
                    if exc.endpoint=="GEN_TOKEN":
                        monitor.auth_status("RETRYING" if kind in ("RATE_LIMITED","SERVER_UNAVAILABLE","AUTH_ENDPOINT_ERROR","AUTH_TOKEN_EXPIRED") and attempt<MAX_ATTEMPTS else "ERROR",kind)
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

        ready=prepare_all(db,transport,None,OUTPUT,log,emit=emit,catalogs=catalogs,mode=mode,settings=settings)
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
        finally:
            monitor.stop(status)


def notify_collection_progress(db,rid,on_progress):
    if on_progress is not None:
        try:
            on_progress(collection_snapshot(db.conn,rid))
        except Exception as exc:
            log(f"Monitor update unavailable: {type(exc).__name__}")


def run_extraction(db,workers=None,stop_event=None,on_progress=None):
    from oic_all_dim_prepare import PLANNER_VERSION,PLAN_POLICY,inventory_fingerprint,load_all_catalogs,matching_preparation
    catalogs=load_all_catalogs()
    prep=matching_preparation(db.conn,catalogs)
    if not prep or prep["status"]!="READY":
        raise RuntimeError("Generate Job Plan before collection")
    summary=json.loads(prep["summary_json"] or "{}")
    if (summary.get("scope")!="WEB_ONLY" or summary.get("planner_version")!=PLANNER_VERSION
            or summary.get("plan_policy")!=PLAN_POLICY
            or prep["catalog_hash"]!=inventory_fingerprint(catalogs)):
        raise RuntimeError("Saved plan is not the current web-only scope; run --prepare first")
    settings=summary.get("plan_settings") or {}
    if workers is None:
        workers=settings.get("workers",DEFAULT_WORKERS)
    if workers!=settings.get("workers"):
        raise RuntimeError("Worker count differs from saved Job Plan; generate a matching plan")
    if db.all_work_finished(prep["id"]):
        failed_exports=export_consolidated_excel(db,prep["id"])
        progress=db.progress(prep["id"])
        log(f"No pending work | {progress}")
        return int(bool(failed_exports) or any(k in progress for k in ("FAILED","NEEDS_REPREPARE")))
    rid=db.start_run(prep["id"],workers,min(workers,prep["concurrency"]),os.getpid(),psutil.Process().create_time())
    configure_log(OUTPUT/"runs"/rid/"run.log",console_output=on_progress is None)
    code=0
    state="FAILED"
    try:
        notify_collection_progress(db,rid,on_progress)
        launch_worker_tabs(rid,list(range(1,workers+1)),include_monitor=True)
        wait_workers_ready(db,rid,workers,stop_event,on_progress)
        supervise_workers(db,rid,workers,stop_event,on_progress)
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
        if shutdown_workers(db,rid,on_progress):
            log("EXPORTING | consolidating all committed raw XML")
            failed_exports=export_consolidated_excel(db,prep["id"],rid,on_progress)
            if failed_exports:
                if state=="COMPLETED":
                    state="PARTIAL"
                code=1
            db.finish_run(rid,state)
        else:
            log("Workers still stopping; resume only after their processes exit")
        notify_collection_progress(db,rid,on_progress)
        log(f"RUN {state} | {db.progress(prep['id'])}")
    return code


def build_parser():
    parser=argparse.ArgumentParser(description=__doc__)
    group=parser.add_mutually_exclusive_group()
    group.add_argument("--prepare",action="store_true")
    group.add_argument("--refresh-members",action="store_true")
    group.add_argument("--job-plan",action="store_true")
    group.add_argument("--status",action="store_true")
    group.add_argument("--export-only",action="store_true")
    parser.add_argument("--workers",type=int,default=None)
    parser.add_argument("--years-per-request",type=int,default=1)
    parser.add_argument("--members-per-request",type=int,default=20)
    parser.add_argument("--timeout",type=int,default=60)
    parser.add_argument("--delay",type=float,default=1.0)
    group.add_argument("--monitor",action="store_true",help=argparse.SUPPRESS)
    parser.add_argument("--worker-id",type=int,help=argparse.SUPPRESS)
    parser.add_argument("--run-id",help=argparse.SUPPRESS)
    return parser


def main():
    args=build_parser().parse_args()
    configure_log()
    if args.workers is not None and not 1<=args.workers<=16:
        raise RuntimeError("Worker count must be between 1 and 16")
    if args.monitor:
        if not args.run_id or args.worker_id is not None:
            raise RuntimeError("Monitor requires --run-id")
        return CollectionMonitor(args.run_id).run()
    if args.worker_id is not None or args.run_id is not None:
        if args.worker_id is None or not args.run_id or args.prepare or args.status or args.refresh_members or args.job_plan:
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
            if args.export_only:
                from oic_all_dim_prepare import load_all_catalogs,matching_preparation
                prep=matching_preparation(db.conn,load_all_catalogs())
                if not prep:
                    raise RuntimeError("No matching preparation")
                return export_consolidated_excel(db,prep["id"])
            from oic_all_dim_prepare import plan_settings
            settings=plan_settings(args.workers or DEFAULT_WORKERS,args.years_per_request,
                                   args.members_per_request,args.timeout,args.delay)
            return run_prepare(db,mode="members" if args.refresh_members else "plan" if args.job_plan else "all",settings=settings) if (args.prepare or args.refresh_members or args.job_plan) else run_extraction(db,args.workers)
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
