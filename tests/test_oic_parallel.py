import json
import re
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch
from xml.etree import ElementTree as ET

import oic_parallel_collector as parallel
import oic_query as q
from test_oic_query import sample_xml


class FakeClient:
    calls = []

    def __init__(self, **kwargs):
        pass

    def execute(self, request_body, destination):
        text = ET.fromstring(request_body).find(
            f".//{{{q.XMLA_NS}}}Statement"
        ).text
        # X and Y have independent Subset offsets in the generated MDX.
        offsets = re.findall(r",(\d+),(\d+)\)", text.replace(" ", ""))
        x_offset, y_offset = int(offsets[0][0]), int(offsets[1][0])
        FakeClient.calls.append((x_offset, y_offset))
        x_count = 1 if x_offset < 2 else 0
        y_count = 1 if y_offset == 0 else 0
        destination.write_bytes(sample_xml(x_count=x_count, y_count=y_count))
        return {"http_status": 200, "response_bytes": destination.stat().st_size}

    def close(self):
        pass


class SelectedParallelTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.config = q.default_query("MV", workers=2)
        self.config["run_id"] = "test_run"
        self.config["delay"] = 0
        self.config["timeout"] = 10
        self.path = self.root / "query_config.json"
        self.path.write_text(json.dumps(self.config), encoding="utf-8")
        FakeClient.calls = []

    def worker(self, worker_id):
        return parallel.selected_worker(
            SimpleNamespace(query_config=self.path, worker_id=worker_id)
        )

    def test_two_workers_disjoint_chunks_resume_and_excel(self):
        with patch.object(parallel, "read_env", lambda: {"OIC_COOKIE": "test_cookie"}), \
             patch.object(parallel, "OICClient", FakeClient):
            self.assertEqual(self.worker(1), 0)
            self.assertEqual(self.worker(2), 0)
            initial = list(FakeClient.calls)
            self.assertTrue(initial)
            first = {x for x, _ in initial if x % 2 == 0}
            second = {x for x, _ in initial if x % 2 == 1}
            self.assertTrue(first)
            self.assertTrue(second)
            self.assertEqual(first & second, set())
            # Rerun must validate saved XML without sending the completed chunks again.
            self.assertEqual(self.worker(1), 0)
            self.assertEqual(self.worker(2), 0)
            self.assertEqual(FakeClient.calls, initial)
        workers, totals = parallel.selected_summary(self.config, self.root)
        self.assertEqual(totals["workers_completed"], 2)
        self.assertEqual(totals["workers_failed"], 0)
        self.assertEqual(len(workers), 2)
        path, count = parallel.selected_export(self.config, self.root)
        self.assertEqual(count, 2)
        self.assertTrue(Path(path).is_file())

    def test_http_503_is_recorded_and_bounded(self):
        class ErrorClient(FakeClient):
            calls = 0

            def execute(self, request_body, destination):
                ErrorClient.calls += 1
                destination.write_bytes(b"The service is unavailable.")
                return {"http_status": 503, "response_bytes": 27}

        with patch.object(parallel, "read_env", lambda: {"OIC_COOKIE": "test_cookie"}), \
             patch.object(parallel, "OICClient", ErrorClient), \
             patch.object(parallel.time, "sleep", lambda _: None):
            self.assertEqual(self.worker(1), 1)
        self.assertEqual(ErrorClient.calls, 2)
        state = json.loads((self.root / "worker_01" / "state.json").read_text())
        self.assertEqual(state["status"], "FAILED")
        self.assertEqual(state["last_http"], 503)
        manifest = json.loads((self.root / "worker_01" / "manifest.json").read_text())
        self.assertEqual([c["status"] for c in manifest["chunks"].values()], ["FAILED"])

    def test_reused_client_referer_covers_five_categories(self):
        for code, profile in q.PROFILES.items():
            client = parallel.OICClient(cookie="ASP.NET_SessionId=TEST",
                                         token="", timeout=5, max_response_mb=1,
                                         worker_id=1, code=code)
            try:
                self.assertTrue(client.referer().endswith("?Class=" + profile["cls"]))
            finally:
                client.close()

    def test_stop_flag_prevents_new_requests(self):
        (self.root / "stop.flag").touch()
        with patch.object(parallel, "read_env", lambda: {"OIC_COOKIE": "test_cookie"}), \
             patch.object(parallel, "OICClient", FakeClient):
            self.assertEqual(self.worker(1), 0)
        self.assertEqual(FakeClient.calls, [])
        state = json.loads((self.root / "worker_01" / "state.json").read_text())
        self.assertEqual(state["status"], "STOPPED")


if __name__ == "__main__":
    unittest.main()
