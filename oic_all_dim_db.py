"""SQLite persistence for the prepared OIC request queue.

Network calls never run inside write transactions. Each process/thread owns
its connection. Request leases fence stale workers and commits are atomic.
"""
from __future__ import annotations

import hashlib
import json
import sqlite3
import time
import uuid
from contextlib import contextmanager
from pathlib import Path

SCHEMA_VERSION = 1
TERMINAL = ("COMMITTED", "EMPTY", "FAILED", "NEEDS_REPREPARE", "REPLACED")


def canonical(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def digest(value):
    data = value if isinstance(value, bytes) else canonical(value).encode("utf-8")
    return hashlib.sha256(data).hexdigest()


class OICDatabase:
    def __init__(self, path):
        self.path = Path(path).resolve()
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.conn = sqlite3.connect(self.path, timeout=30, isolation_level=None)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA busy_timeout=30000")
        self.conn.execute("PRAGMA foreign_keys=ON")

    def close(self):
        self.conn.close()

    @contextmanager
    def transaction(self):
        self.conn.execute("BEGIN IMMEDIATE")
        try:
            yield
            self.conn.commit()
        except BaseException:
            self.conn.rollback()
            raise

    def initialize_schema(self):
        version = self.conn.execute("PRAGMA user_version").fetchone()[0]
        if version not in (0, SCHEMA_VERSION):
            raise RuntimeError("Unsupported OIC database schema")
        self.conn.execute("PRAGMA journal_mode=WAL")
        self.conn.executescript("""
        CREATE TABLE IF NOT EXISTS oic_preparations(
            id TEXT PRIMARY KEY, catalog_hash TEXT NOT NULL, status TEXT NOT NULL,
            created_at REAL NOT NULL, summary_json TEXT NOT NULL DEFAULT '{}',
            concurrency INTEGER NOT NULL DEFAULT 1
        );
        CREATE TABLE IF NOT EXISTS oic_cubes(
            prep TEXT NOT NULL REFERENCES oic_preparations(id),
            cube TEXT NOT NULL, profile_json TEXT NOT NULL,
            PRIMARY KEY(prep,cube)
        );
        CREATE TABLE IF NOT EXISTS oic_fields(
            prep TEXT NOT NULL, cube TEXT NOT NULL, field_id TEXT NOT NULL,
            field_json TEXT NOT NULL, PRIMARY KEY(prep,cube,field_id),
            FOREIGN KEY(prep,cube) REFERENCES oic_cubes(prep,cube)
        );
        CREATE TABLE IF NOT EXISTS oic_dimensions(
            prep TEXT NOT NULL, cube TEXT NOT NULL, hierarchy TEXT NOT NULL,
            status TEXT NOT NULL, metadata_json TEXT NOT NULL,
            PRIMARY KEY(prep,cube,hierarchy),
            FOREIGN KEY(prep,cube) REFERENCES oic_cubes(prep,cube)
        );
        CREATE TABLE IF NOT EXISTS oic_levels(
            prep TEXT NOT NULL, cube TEXT NOT NULL, hierarchy TEXT NOT NULL,
            level TEXT NOT NULL, metadata_json TEXT NOT NULL,
            PRIMARY KEY(prep,cube,hierarchy,level),
            FOREIGN KEY(prep,cube,hierarchy) REFERENCES oic_dimensions(prep,cube,hierarchy)
        );
        CREATE TABLE IF NOT EXISTS oic_members(
            prep TEXT NOT NULL, cube TEXT NOT NULL, hierarchy TEXT NOT NULL,
            unique_name TEXT NOT NULL, caption TEXT, parent TEXT, level TEXT,
            metadata_json TEXT NOT NULL,
            PRIMARY KEY(prep,cube,hierarchy,unique_name),
            FOREIGN KEY(prep,cube,hierarchy) REFERENCES oic_dimensions(prep,cube,hierarchy)
        );
        CREATE TABLE IF NOT EXISTS oic_measures(
            prep TEXT NOT NULL, cube TEXT NOT NULL, measure TEXT NOT NULL,
            metadata_json TEXT NOT NULL, PRIMARY KEY(prep,cube,measure),
            FOREIGN KEY(prep,cube) REFERENCES oic_cubes(prep,cube)
        );
        CREATE TABLE IF NOT EXISTS oic_profiles(
            prep TEXT NOT NULL, cube TEXT NOT NULL, profile TEXT NOT NULL,
            metadata_json TEXT NOT NULL, PRIMARY KEY(prep,cube,profile)
        );
        CREATE TABLE IF NOT EXISTS oic_jobs(
            prep TEXT NOT NULL, id TEXT NOT NULL, cube TEXT NOT NULL,
            hierarchy TEXT NOT NULL, scope_json TEXT NOT NULL,
            PRIMARY KEY(prep,id)
        );
        CREATE TABLE IF NOT EXISTS oic_requests(
            id TEXT PRIMARY KEY, prep TEXT NOT NULL, job TEXT NOT NULL,
            cube TEXT NOT NULL, parent TEXT REFERENCES oic_requests(id),
            scope_json TEXT NOT NULL, mdx TEXT NOT NULL, payload BLOB NOT NULL,
            payload_hash TEXT NOT NULL, timeout INTEGER NOT NULL,
            potential_cells INTEGER NOT NULL, status TEXT NOT NULL,
            attempts INTEGER NOT NULL DEFAULT 0, available_at REAL NOT NULL DEFAULT 0,
            owner TEXT, lease TEXT, expires REAL, response_path TEXT, response_hash TEXT,
            error TEXT, metrics_json TEXT, committed_at REAL,
            FOREIGN KEY(prep,job) REFERENCES oic_jobs(prep,id)
        );
        CREATE INDEX IF NOT EXISTS request_queue
            ON oic_requests(prep,status,available_at,id);
        CREATE INDEX IF NOT EXISTS request_parent ON oic_requests(parent,status);
        CREATE TABLE IF NOT EXISTS oic_observations(
            id TEXT PRIMARY KEY, prep TEXT NOT NULL, cube TEXT NOT NULL,
            grain TEXT NOT NULL, is_total INTEGER NOT NULL,
            last_data_update TEXT, retrieved_at REAL NOT NULL
        );
        CREATE INDEX IF NOT EXISTS observation_version ON oic_observations(prep,cube,last_data_update);
        CREATE TABLE IF NOT EXISTS oic_coordinates(
            observation TEXT NOT NULL REFERENCES oic_observations(id),
            hierarchy TEXT NOT NULL, member TEXT NOT NULL,
            PRIMARY KEY(observation,hierarchy)
        );
        CREATE TABLE IF NOT EXISTS oic_values(
            observation TEXT NOT NULL REFERENCES oic_observations(id),
            measure TEXT NOT NULL, value_raw TEXT, value_numeric TEXT,
            value_formatted TEXT, value_state TEXT NOT NULL,
            request TEXT NOT NULL REFERENCES oic_requests(id),
            PRIMARY KEY(observation,measure)
        );
        CREATE TABLE IF NOT EXISTS oic_runs(
            id TEXT PRIMARY KEY, prep TEXT NOT NULL REFERENCES oic_preparations(id),
            supervisor_pid INTEGER NOT NULL, supervisor_created REAL NOT NULL,
            workers INTEGER NOT NULL, allowed INTEGER NOT NULL,
            state TEXT NOT NULL, barrier INTEGER NOT NULL DEFAULT 0,
            stop INTEGER NOT NULL DEFAULT 0, cooldown REAL NOT NULL DEFAULT 0,
            created_at REAL NOT NULL
        );
        CREATE TABLE IF NOT EXISTS oic_workers(
            run TEXT NOT NULL REFERENCES oic_runs(id), worker INTEGER NOT NULL,
            owner TEXT NOT NULL UNIQUE, pid INTEGER NOT NULL, created REAL NOT NULL,
            heartbeat REAL NOT NULL, status TEXT NOT NULL, slot INTEGER NOT NULL DEFAULT 0,
            PRIMARY KEY(run,worker)
        );
        """)
        self.conn.execute("PRAGMA user_version=1")

    def preparation(self, fingerprint):
        with self.transaction():
            row = self.conn.execute(
                "SELECT * FROM oic_preparations WHERE catalog_hash=? ORDER BY created_at DESC LIMIT 1",
                (fingerprint,)).fetchone()
            if row and row["status"] in ("PREPARING", "READY"):
                return dict(row)
            pid = uuid.uuid4().hex
            self.conn.execute("INSERT INTO oic_preparations(id,catalog_hash,status,created_at) VALUES(?,?,'PREPARING',?)",
                              (pid, fingerprint, time.time()))
            return dict(self.conn.execute("SELECT * FROM oic_preparations WHERE id=?", (pid,)).fetchone())

    def clone_inventory(self, source, fingerprint):
        """Copy only catalog/member inventory, leaving original extraction intact."""
        prep=uuid.uuid4().hex
        with self.transaction():
            self.conn.execute("INSERT INTO oic_preparations(id,catalog_hash,status,created_at) VALUES(?,?,'PREPARING',?)",
                              (prep,fingerprint,time.time()))
            for table,columns in (
                ("oic_cubes","cube,profile_json"),
                ("oic_fields","cube,field_id,field_json"),
                ("oic_dimensions","cube,hierarchy,status,metadata_json"),
                ("oic_levels","cube,hierarchy,level,metadata_json"),
                ("oic_members","cube,hierarchy,unique_name,caption,parent,level,metadata_json"),
                ("oic_measures","cube,measure,metadata_json"),
                ("oic_profiles","cube,profile,metadata_json"),
            ):
                where=" AND profile LIKE 'DISCOVERY_METADATA:%'" if table=="oic_profiles" else ""
                self.conn.execute(f"INSERT INTO {table}(prep,{columns}) SELECT ?,{columns} FROM {table} WHERE prep=?"+where,(prep,source))
            return dict(self.conn.execute("SELECT * FROM oic_preparations WHERE id=?",(prep,)).fetchone())

    def latest_preparation(self):
        row = self.conn.execute("SELECT * FROM oic_preparations ORDER BY created_at DESC LIMIT 1").fetchone()
        return dict(row) if row else None

    def save_catalog(self, prep, cube, profile, fields, measures):
        with self.transaction():
            self.conn.execute("""INSERT INTO oic_cubes VALUES(?,?,?)
                ON CONFLICT(prep,cube) DO UPDATE SET profile_json=excluded.profile_json""",
                              (prep,cube,canonical(profile)))
            self.conn.execute("DELETE FROM oic_fields WHERE prep=? AND cube=?",(prep,cube))
            self.conn.execute("DELETE FROM oic_measures WHERE prep=? AND cube=?",(prep,cube))
            self.conn.executemany("INSERT OR REPLACE INTO oic_fields VALUES(?,?,?,?)",
                [(prep,cube,str(f.get("id") or digest(f)),canonical(f)) for f in fields])
            self.conn.executemany("INSERT OR REPLACE INTO oic_measures VALUES(?,?,?,?)",
                [(prep,cube,m,canonical(meta)) for m,meta in measures.items()])

    def profile(self, prep, cube):
        row = self.conn.execute("SELECT profile_json FROM oic_cubes WHERE prep=? AND cube=?",
                                (prep,cube)).fetchone()
        if not row:
            raise RuntimeError("Unknown prepared cube")
        return json.loads(row[0])

    def inventory(self, prep, cube, hierarchy):
        row = self.conn.execute("SELECT status FROM oic_dimensions WHERE prep=? AND cube=? AND hierarchy=?",
                                (prep,cube,hierarchy)).fetchone()
        if not row or row[0] not in ("COMPLETE","EMPTY"):
            return None
        return [json.loads(r[0]) for r in self.conn.execute(
            "SELECT metadata_json FROM oic_members WHERE prep=? AND cube=? AND hierarchy=? ORDER BY unique_name",
            (prep,cube,hierarchy))]

    def save_inventory(self, prep, cube, hierarchy, status, metadata, levels, members):
        with self.transaction():
            self.conn.execute("""INSERT INTO oic_dimensions VALUES(?,?,?,?,?)
                ON CONFLICT(prep,cube,hierarchy) DO UPDATE SET status=excluded.status,metadata_json=excluded.metadata_json""",
                (prep,cube,hierarchy,status,canonical(metadata)))
            self.conn.execute("DELETE FROM oic_levels WHERE prep=? AND cube=? AND hierarchy=?", (prep,cube,hierarchy))
            self.conn.execute("DELETE FROM oic_members WHERE prep=? AND cube=? AND hierarchy=?", (prep,cube,hierarchy))
            self.conn.executemany("INSERT INTO oic_levels VALUES(?,?,?,?,?)",
                [(prep,cube,hierarchy,l["LEVEL_UNIQUE_NAME"],canonical(l)) for l in levels if l.get("LEVEL_UNIQUE_NAME")])
            self.conn.executemany("INSERT INTO oic_members VALUES(?,?,?,?,?,?,?,?)",
                [(prep,cube,hierarchy,m["unique_name"],m.get("caption"),m.get("parent"),
                  m.get("level"),canonical(m)) for m in members])

    def load_discovery_checkpoint(self, prep, cube, hierarchy):
        row=self.conn.execute("SELECT status,metadata_json FROM oic_dimensions WHERE prep=? AND cube=? AND hierarchy=?",
                              (prep,cube,hierarchy)).fetchone()
        if not row:
            return {"status":"PENDING","next_offset":0}
        return dict(json.loads(row["metadata_json"]),status=row["status"])

    def save_cardinality(self, prep, cube, hierarchy, metadata, levels, count=None):
        with self.transaction():
            state=self.load_discovery_checkpoint(prep,cube,hierarchy)
            if state["status"] in ("COMPLETE","EMPTY"):
                return
            status=state.pop("status")
            state.update(metadata=metadata,reported_cardinality=metadata.get("HIERARCHY_CARDINALITY"))
            if count is not None:
                previous=state.get("expected_count")
                if previous is not None and previous!=count and state.get("next_offset",0):
                    # Previous pages no longer belong to a consistent enumeration.
                    self.conn.execute("DELETE FROM oic_members WHERE prep=? AND cube=? AND hierarchy=?",
                                      (prep,cube,hierarchy))
                    state.update(next_offset=0,last_member=None,page_size=100)
                state["expected_count"]=count
            status="PARTIAL" if state.get("next_offset",0) else "COUNTED" if state.get("expected_count") is not None else "PENDING"
            self.conn.execute("""INSERT INTO oic_dimensions VALUES(?,?,?,?,?)
                ON CONFLICT(prep,cube,hierarchy) DO UPDATE SET status=excluded.status,metadata_json=excluded.metadata_json""",
                (prep,cube,hierarchy,status,canonical(state)))
            self.conn.executemany("INSERT OR REPLACE INTO oic_levels VALUES(?,?,?,?,?)",
                [(prep,cube,hierarchy,l["LEVEL_UNIQUE_NAME"],canonical(l)) for l in levels if l.get("LEVEL_UNIQUE_NAME")])

    def append_member_page(self, prep, cube, hierarchy, offset, members, page_size):
        if not members:
            raise ValueError("Empty page cannot advance discovery")
        names=[m["unique_name"] for m in members]
        if len(set(names))!=len(names):
            raise ValueError("Duplicate members inside discovery page")
        with self.transaction():
            state=self.load_discovery_checkpoint(prep,cube,hierarchy)
            if state["status"] in ("COMPLETE","EMPTY") or state.get("next_offset",0)!=offset:
                raise ValueError("Discovery checkpoint offset mismatch")
            existing=self.conn.execute("SELECT count(*) FROM oic_members WHERE prep=? AND cube=? AND hierarchy=?",
                                       (prep,cube,hierarchy)).fetchone()[0]
            if existing!=offset:
                raise ValueError("Stored member count does not match checkpoint")
            expected=state.get("expected_count")
            if expected is None or offset+len(names)>expected:
                raise ValueError("Page exceeds validated member count")
            try:
                self.conn.executemany("INSERT INTO oic_members VALUES(?,?,?,?,?,?,?,?)",
                    [(prep,cube,hierarchy,m["unique_name"],m.get("caption"),m.get("parent"),m.get("level"),canonical(m))
                     for m in members])
            except sqlite3.IntegrityError as exc:
                raise ValueError("Discovery page overlaps stored members") from exc
            self.conn.executemany("INSERT OR IGNORE INTO oic_levels VALUES(?,?,?,?,?)",
                [(prep,cube,hierarchy,m["level"],canonical({"LEVEL_UNIQUE_NAME":m["level"],
                   "LEVEL_NUMBER":m.get("level_number"),"source":"MEMBER_METADATA"})) for m in members if m.get("level")])
            state.pop("status",None)
            state.update(next_offset=offset+len(names),last_member=names[-1],page_size=page_size,method="MDX_PAGINATION")
            state.pop("error",None)
            self.conn.execute("UPDATE oic_dimensions SET status='PARTIAL',metadata_json=? WHERE prep=? AND cube=? AND hierarchy=?",
                              (canonical(state),prep,cube,hierarchy))

    def reset_partial_inventory(self, prep, cube, hierarchy):
        with self.transaction():
            state=self.load_discovery_checkpoint(prep,cube,hierarchy)
            if state["status"] in ("COMPLETE","EMPTY"):
                raise ValueError("Cannot reset completed inventory")
            state.pop("status",None)
            state.update(next_offset=0,last_member=None,page_size=100)
            self.conn.execute("DELETE FROM oic_members WHERE prep=? AND cube=? AND hierarchy=?",(prep,cube,hierarchy))
            self.conn.execute("UPDATE oic_dimensions SET status='COUNTED',metadata_json=? WHERE prep=? AND cube=? AND hierarchy=?",
                              (canonical(state),prep,cube,hierarchy))

    def finalize_inventory(self, prep, cube, hierarchy, count):
        with self.transaction():
            state=self.load_discovery_checkpoint(prep,cube,hierarchy)
            actual=self.conn.execute("SELECT count(*) FROM oic_members WHERE prep=? AND cube=? AND hierarchy=?",
                                     (prep,cube,hierarchy)).fetchone()[0]
            if state.get("expected_count")!=count or actual!=count or state.get("next_offset",0)!=count:
                raise ValueError("Member inventory coverage mismatch")
            state.pop("status",None)
            state.pop("error",None)
            state.update(discovered_cardinality=count,source="XMLA_COUNT_VALIDATED")
            self.conn.execute("UPDATE oic_dimensions SET status=?,metadata_json=? WHERE prep=? AND cube=? AND hierarchy=?",
                              ("COMPLETE" if count else "EMPTY",canonical(state),prep,cube,hierarchy))

    def mark_discovery_failure(self, prep, cube, hierarchy, error):
        with self.transaction():
            state=self.load_discovery_checkpoint(prep,cube,hierarchy)
            if state["status"] in ("COMPLETE","EMPTY"):
                return
            state.pop("status",None)
            state["error"]=str(error)[:300]
            state["failures"]=state.get("failures",0)+1
            status="PARTIAL" if state.get("next_offset",0) else "FAILED"
            self.conn.execute("""INSERT INTO oic_dimensions VALUES(?,?,?,?,?)
                ON CONFLICT(prep,cube,hierarchy) DO UPDATE SET status=excluded.status,metadata_json=excluded.metadata_json""",
                (prep,cube,hierarchy,status,canonical(state)))

    def save_profile(self, prep, cube, profile, data):
        self.conn.execute("INSERT OR REPLACE INTO oic_profiles VALUES(?,?,?,?)",
                          (prep,cube,profile,canonical(data)))

    def load_profile(self, prep,cube,profile):
        row = self.conn.execute("SELECT metadata_json FROM oic_profiles WHERE prep=? AND cube=? AND profile=?",
                                (prep,cube,profile)).fetchone()
        return json.loads(row[0]) if row else None

    def save_job(self, prep, job, cube, hierarchy, scope):
        self.conn.execute("INSERT OR IGNORE INTO oic_jobs VALUES(?,?,?,?,?)",
                          (prep,job,cube,hierarchy,canonical(scope)))

    def insert_requests(self, requests):
        with self.transaction():
            self.conn.executemany("""INSERT OR IGNORE INTO oic_requests
                (id,prep,job,cube,parent,scope_json,mdx,payload,payload_hash,timeout,potential_cells,status)
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?)""",
                [(r["id"],r["prep"],r["job"],r["cube"],r.get("parent"),canonical(r["scope"]),
                  r["mdx"],r["payload"],digest(r["payload"]),r["timeout"],r["potential_cells"],
                  r.get("status","PENDING")) for r in requests])

    def get_request(self, rid):
        row = self.conn.execute("SELECT * FROM oic_requests WHERE id=?", (rid,)).fetchone()
        return dict(row) if row else None

    def finalize_preparation(self, prep, status, summary, concurrency=1):
        self.conn.execute("UPDATE oic_preparations SET status=?,summary_json=?,concurrency=? WHERE id=?",
                          (status,canonical(summary),concurrency,prep))

    def start_run(self, prep, workers, allowed, pid, created):
        rid = uuid.uuid4().hex
        with self.transaction():
            if self.conn.execute("SELECT 1 FROM oic_runs WHERE state='RUNNING'").fetchone():
                raise RuntimeError("Another extraction run is active")
            self.conn.execute("""INSERT INTO oic_runs
                (id,prep,supervisor_pid,supervisor_created,workers,allowed,state,created_at)
                VALUES(?,?,?,?,?,?,'RUNNING',?)""", (rid,prep,pid,created,workers,allowed,time.time()))
        return rid

    def run(self, rid):
        row = self.conn.execute("SELECT * FROM oic_runs WHERE id=?", (rid,)).fetchone()
        return dict(row) if row else None

    def active_runs(self):
        return [dict(r) for r in self.conn.execute("SELECT * FROM oic_runs WHERE state='RUNNING'")]

    def workers(self,rid):
        return [dict(r) for r in self.conn.execute("SELECT * FROM oic_workers WHERE run=?", (rid,))]

    def register_worker(self, rid, worker, owner, pid, created):
        with self.transaction():
            row=self.run(rid)
            if not row or row["state"]!="RUNNING" or row["stop"] or not 1<=worker<=row["workers"]:
                raise RuntimeError("Run is not available")
            self.conn.execute("INSERT INTO oic_workers VALUES(?,?,?,?,?,?,'READY',0)",
                              (rid,worker,owner,pid,created,time.time()))

    def heartbeat(self, owner, lease=None):
        now=time.time()
        with self.transaction():
            self.conn.execute("UPDATE oic_workers SET heartbeat=? WHERE owner=?", (now,owner))
            self.conn.execute("UPDATE oic_requests SET expires=? WHERE owner=? AND status='RUNNING'",
                              (now+60,owner))

    def barrier(self,rid):
        self.conn.execute("UPDATE oic_runs SET barrier=1 WHERE id=?", (rid,))

    def stop(self,rid):
        self.conn.execute("UPDATE oic_runs SET stop=1 WHERE id=?", (rid,))

    def finish_run(self,rid,state):
        self.conn.execute("UPDATE oic_runs SET state=?,stop=1 WHERE id=?", (state,rid))

    def claim_next_request(self,rid,owner):
        with self.transaction():
            run=self.run(rid)
            if not run or run["stop"] or not run["barrier"]:
                return None
            worker=self.conn.execute("SELECT worker FROM oic_workers WHERE run=? AND owner=? AND status='READY'",(rid,owner)).fetchone()
            if not worker:
                return None
            # Select the earliest unfinished assigned job, even if its retry is cooling down.
            job=self.conn.execute("""SELECT j.id FROM oic_jobs j
                WHERE j.prep=? AND json_extract(j.scope_json,'$.worker_id')=?
                AND EXISTS (SELECT 1 FROM oic_requests r WHERE r.prep=j.prep AND r.job=j.id
                    AND r.status IN ('PENDING','RETRY','RUNNING'))
                ORDER BY json_extract(j.scope_json,'$.job_index') LIMIT 1""",
                (run["prep"],worker[0])).fetchone()
            if not job:
                return None
            row=self.conn.execute("""SELECT * FROM oic_requests WHERE prep=? AND job=?
                AND status IN ('PENDING','RETRY','RUNNING')
                ORDER BY json_extract(scope_json,'$.year_offset'),
                         json_extract(scope_json,'$.member_offset'),id LIMIT 1""",
                (run["prep"],job[0])).fetchone()
            if row and (row["status"]=="RUNNING" or row["available_at"]>time.time()):
                return None
            if not row:
                return None
            lease=uuid.uuid4().hex
            self.conn.execute("UPDATE oic_requests SET status='RUNNING',owner=?,lease=?,expires=?,attempts=attempts+1 WHERE id=?",
                              (owner,lease,time.time()+60,row["id"]))
            return self.get_request(row["id"])

    def claim_probe(self, rid):
        with self.transaction():
            row=self.get_request(rid)
            if row["status"] in ("COMMITTED","EMPTY"):
                return None
            if row["status"] not in ("PROBE","RETRY"):
                raise RuntimeError("Probe unavailable")
            self.conn.execute("UPDATE oic_requests SET status='RUNNING',owner='PREPARE',lease=?,attempts=attempts+1 WHERE id=?",
                              (uuid.uuid4().hex,rid))
            return self.get_request(rid)

    def _fence(self,r):
        row=self.get_request(r["id"])
        if not row or row["status"]!="RUNNING" or row["lease"]!=r["lease"]:
            raise RuntimeError("Request lease lost")

    def register_response(self,r,path,sha,metrics):
        with self.transaction():
            self._fence(r)
            self.conn.execute("UPDATE oic_requests SET response_path=?,response_hash=?,metrics_json=? WHERE id=?",
                              (str(path),sha,canonical(metrics),r["id"]))

    def commit_request(self,r,records,metrics):
        observations, coordinates, values = {}, [], []
        now=time.time()
        for record in records:
            coords=sorted(record["coordinates"])
            oid=digest([r["prep"],r["cube"],record["grain"],coords])
            if oid not in observations:
                observations[oid]=(oid,r["prep"],r["cube"],record["grain"],int(record["is_total"]),record.get("last_data_update"),now)
                coordinates.extend((oid,h,m) for h,m in coords)
            values.append((oid,record["measure"],record.get("raw"),record.get("numeric"),
                           record.get("formatted"),record["state"],r["id"]))
        with self.transaction():
            self._fence(r)
            version=metrics.get("source_update")
            if version:
                prior=self.conn.execute("SELECT last_data_update FROM oic_observations WHERE prep=? AND cube=? AND last_data_update IS NOT NULL LIMIT 1",
                                        (r["prep"],r["cube"])).fetchone()
                if prior and prior[0]!=version:
                    raise ValueError("SOURCE_VERSION_CHANGED")
            self.conn.executemany("INSERT OR IGNORE INTO oic_observations VALUES(?,?,?,?,?,?,?)",observations.values())
            self.conn.executemany("INSERT OR IGNORE INTO oic_coordinates VALUES(?,?,?)",coordinates)
            self.conn.executemany("""INSERT INTO oic_values VALUES(?,?,?,?,?,?,?)
                ON CONFLICT(observation,measure) DO UPDATE SET
                value_raw=excluded.value_raw,value_numeric=excluded.value_numeric,
                value_formatted=excluded.value_formatted,value_state=excluded.value_state,request=excluded.request""",values)
            self.conn.execute("""UPDATE oic_requests SET status=?,metrics_json=?,committed_at=?,
                owner=NULL,lease=NULL,expires=NULL,error=NULL WHERE id=?""",
                ("COMMITTED" if metrics.get("useful_cells") else "EMPTY",canonical(metrics),now,r["id"]))
        return len(values)

    def fail(self,r,error,status="FAILED",delay=0):
        with self.transaction():
            self._fence(r)
            self.conn.execute("""UPDATE oic_requests SET status=?,error=?,available_at=?,
                owner=NULL,lease=NULL,expires=NULL WHERE id=?""",
                (status,str(error)[:500],time.time()+delay,r["id"]))

    def activate_recovery_requests(self,r,error,children=None):
        with self.transaction():
            self._fence(r)
            if children:
                if len(children)!=2 or sum(c["potential_cells"] for c in children)!=r["potential_cells"]:
                    return False
                self.conn.executemany("""INSERT OR IGNORE INTO oic_requests
                    (id,prep,job,cube,parent,scope_json,mdx,payload,payload_hash,timeout,potential_cells,status)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?, 'STANDBY')""",
                    [(c["id"],c["prep"],c["job"],c["cube"],c["parent"],canonical(c["scope"]),
                      c["mdx"],c["payload"],digest(c["payload"]),c["timeout"],c["potential_cells"]) for c in children])
            rows=self.conn.execute("SELECT potential_cells FROM oic_requests WHERE parent=? AND status='STANDBY'",(r["id"],)).fetchall()
            if len(rows)!=2 or sum(row[0] for row in rows)!=r["potential_cells"]:
                return False
            self.conn.execute("UPDATE oic_requests SET status='PENDING' WHERE parent=? AND status='STANDBY'",(r["id"],))
            self.conn.execute("UPDATE oic_requests SET status='REPLACED',error=?,owner=NULL,lease=NULL,expires=NULL WHERE id=?",
                              (str(error)[:500],r["id"]))
            return True

    def recover_worker(self,owner):
        with self.transaction():
            self.conn.execute("""UPDATE oic_requests SET status='RETRY',owner=NULL,lease=NULL,expires=NULL
                WHERE owner=? AND status='RUNNING'""",(owner,))
            self.conn.execute("DELETE FROM oic_workers WHERE owner=?",(owner,))

    def recover_preparation(self,prep):
        self.conn.execute("""UPDATE oic_requests SET status='PROBE',owner=NULL,lease=NULL,expires=NULL
            WHERE prep=? AND owner='PREPARE' AND status='RUNNING'""",(prep,))

    def acquire_slot(self,rid,owner):
        with self.transaction():
            run=self.run(rid)
            if not run or run["stop"] or run["cooldown"]>time.time():
                return False
            n=self.conn.execute("SELECT count(*) FROM oic_workers WHERE run=? AND slot=1",(rid,)).fetchone()[0]
            if n>=run["allowed"]:
                return False
            return self.conn.execute("UPDATE oic_workers SET slot=1 WHERE run=? AND owner=? AND slot=0",(rid,owner)).rowcount==1

    def release_slot(self,owner):
        self.conn.execute("UPDATE oic_workers SET slot=0 WHERE owner=?",(owner,))

    def pressure(self,rid,seconds):
        self.conn.execute("UPDATE oic_runs SET cooldown=max(cooldown,?),allowed=max(1,allowed-1) WHERE id=?",
                          (time.time()+seconds,rid))

    def export_result(self,rid,path=None,error=None):
        self.conn.execute("""UPDATE oic_requests SET metrics_json=json_set(coalesce(metrics_json,'{}'),
            '$.excel_path',?,'$.excel_error',?,'$.excel_status',?) WHERE id=? AND status IN ('COMMITTED','EMPTY')""",
            (str(path) if path else None,error,"SAVED" if path and not error else "FAILED",rid))

    def reprepare_request(self, parent, children):
        with self.transaction():
            row=self.get_request(parent["id"])
            if not row or row["status"]!="NEEDS_REPREPARE":
                raise RuntimeError("Request is not eligible for reprepare")
            for r in children:
                self.conn.execute("""INSERT INTO oic_requests
                    (id,prep,job,cube,parent,scope_json,mdx,payload,payload_hash,timeout,potential_cells,status)
                    VALUES(?,?,?,?,?,?,?,?,?,?,?,'PENDING')
                    ON CONFLICT(id) DO UPDATE SET status='PENDING'
                    WHERE oic_requests.status='STANDBY'""",
                    (r["id"],r["prep"],r["job"],r["cube"],r["parent"],canonical(r["scope"]),
                     r["mdx"],r["payload"],digest(r["payload"]),r["timeout"],r["potential_cells"]))
            self.conn.execute("UPDATE oic_requests SET status='REPLACED' WHERE id=?",(parent["id"],))

    def retry_failed(self,prep):
        with self.transaction():
            for prefix in ("NETWORK_ERROR","DNS_ERROR","CONNECTION_RESET","AUTH_TOKEN_EXPIRED",
                           "INVALID_XML","DATABASE_BUSY","DATABASE_WRITE_ERROR","STORAGE_ERROR",
                           "DISK_FULL","RATE_LIMITED","SERVER_UNAVAILABLE"):
                self.conn.execute("""UPDATE oic_requests SET status='RETRY',attempts=0,available_at=0
                    WHERE prep=? AND status='FAILED' AND error LIKE ?""",(prep,prefix+"%"))

    def coverage(self,prep):
        # Replaced parents and standby children are alternative representations,
        # never additional logical coordinates. Successful probes are disjoint.
        total=self.conn.execute("""SELECT coalesce(sum(potential_cells),0) FROM oic_requests
            WHERE prep=? AND status NOT IN ('STANDBY','REPLACED','PROBE','PROBE_FAILED')""",(prep,)).fetchone()[0]
        done=self.conn.execute("""SELECT coalesce(sum(potential_cells),0) FROM oic_requests
            WHERE prep=? AND status IN ('COMMITTED','EMPTY')""",(prep,)).fetchone()[0]
        return {"covered":done,"planned":total,"percent":round(100*done/total,2) if total else 0}

    def progress(self,prep):
        rows=self.conn.execute("""SELECT status,count(*) n,sum(potential_cells) cells
            FROM oic_requests WHERE prep=? GROUP BY status""",(prep,)).fetchall()
        return {r["status"]:{"requests":r["n"],"cells":r["cells"] or 0} for r in rows}

    def all_work_finished(self,prep):
        return not self.conn.execute("""SELECT 1 FROM oic_requests WHERE prep=?
            AND status IN ('PENDING','RETRY','RUNNING') LIMIT 1""",(prep,)).fetchone()

    def finish_worker(self,owner,status):
        self.conn.execute("UPDATE oic_workers SET status=?,slot=0,heartbeat=? WHERE owner=?",
                          (status,time.time(),owner))
