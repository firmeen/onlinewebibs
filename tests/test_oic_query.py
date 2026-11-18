import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from xml.etree import ElementTree as ET

import oic_query as q


def sample_xml(x_count=2, y_count=2):
    """Synthetic XMLA fixture: X-axis tuples, Y-axis dimensions+measure, CellOrdinal."""
    ns = q.MD_NS
    env = ET.Element("{http://schemas.xmlsoap.org/soap/envelope/}Envelope")
    body = ET.SubElement(env, "{http://schemas.xmlsoap.org/soap/envelope/}Body")
    root = ET.SubElement(body, f"{{{ns}}}root")
    info = ET.SubElement(root, f"{{{ns}}}OlapInfo")
    c = ET.SubElement(ET.SubElement(info, f"{{{ns}}}CubeInfo"), f"{{{ns}}}Cube")
    ET.SubElement(c, f"{{{ns}}}CubeName").text = "MV_AGGREGATE_FY"
    ET.SubElement(c, f"{{{ns}}}LastDataUpdate").text = "2026-10-18T04:39:34"
    axes = ET.SubElement(root, f"{{{ns}}}Axes")
    for name, count, dimension in [
        ("Axis0", x_count, "[VW PV MV AGGREGATE FY].[YEAR]"),
        ("Axis1", y_count, "[DIM CT SUBCLASS MV].[SUBCLASS DESC TH]"),
    ]:
        axis = ET.SubElement(axes, f"{{{ns}}}Axis", name=name)
        tuples = ET.SubElement(axis, f"{{{ns}}}Tuples")
        for i in range(count):
            tup = ET.SubElement(tuples, f"{{{ns}}}Tuple")
            for mdx, caption in (
                [(dimension, str(2565 + i))]
                if name == "Axis0"
                else [(dimension, f"Type {i}"), ("[Measures]", "จำนวนกรมธรรม์")]
            ):
                member = ET.SubElement(tup, f"{{{ns}}}Member", Hierarchy=mdx)
                ET.SubElement(member, f"{{{ns}}}UName").text = (
                    "[Measures].[ACT POLICY CNT]" if mdx == "[Measures]" else mdx + f".&[{i}]"
                )
                ET.SubElement(member, f"{{{ns}}}Caption").text = caption
                ET.SubElement(member, f"{{{ns}}}LNum").text = "1"
    cells = ET.SubElement(root, f"{{{ns}}}CellData")
    for ordinal in range(x_count * y_count):
        cell = ET.SubElement(cells, f"{{{ns}}}Cell", CellOrdinal=str(ordinal))
        ET.SubElement(cell, f"{{{ns}}}Value").text = str(10 + ordinal)
    return ET.tostring(env, encoding="utf-8")


class QueryCoreTests(unittest.TestCase):
    def test_profiles_all_five_and_no_guessed_mr_year(self):
        self.assertEqual(set(q.PROFILES), {"AH", "FI", "MR", "MI", "MV"})
        self.assertIsNone(q.PROFILES["MR"]["year"])
        for code in q.PROFILES:
            self.assertGreater(len(q.seeded_catalog(code)["fields"]), 1)

    def test_default_and_validation(self):
        config = q.default_query("MV", workers=1)
        q.validate_query(config)
        self.assertEqual(config["x"], [q.PROFILES["MV"]["year"]])
        self.assertEqual(config["y"], [q.PROFILES["MV"]["row1"]])
        self.assertTrue(q.build_mdx(config).startswith("SELECT"))
        self.assertIn(" ON COLUMNS", q.build_mdx(config))
        self.assertIn(" ON ROWS", q.build_mdx(config))
        self.assertIn("[Measures].[ACT POLICY CNT]", q.build_mdx(config))
        self.assertIn("Subset(", q.build_mdx(config))
        self.assertIn("Subset(", q.build_soap(config).decode())
        self.assertEqual(q.query_hash(config), q.query_hash(config | {"run_id": "another"}))

    def test_duplicates_invalid_fields_and_missing_metadata(self):
        config = q.default_query("MV")
        config["y"] = config["x"][:]
        with self.assertRaises(ValueError):
            q.validate_query(config)
        config = q.default_query("MV")
        config["x"] = ["[Invalid].[Guess]"]
        with self.assertRaises(ValueError):
            q.validate_query(config)
        self.assertEqual(q.default_query("MR")["x"], [])

    def test_catalog_cache_and_search_data(self):
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(q, "CATALOG_DIR", Path(directory)):
                cat = q.seeded_catalog("MV")
                cat["source"] = "LIVE"
                q.save_catalog(cat)
                got = q.load_catalog("MV")
                self.assertEqual(got["source"], "CACHE")
                self.assertTrue(any("กรมธรรม์" in x["display"] for x in got["fields"]))
                self.assertTrue(any("ACT POLICY CNT" in x["mdx"] for x in got["fields"]))

    def test_multidimensional_axis_construction(self):
        config = q.default_query("MV")
        other = "[VW PV MV AGGREGATE FY].[INS CL ID CT PROVINCE]"
        config["y"].append(other)
        cat = q.seeded_catalog("MV")
        cat["fields"].append(dict(mdx=other, type="DIMENSION", display="จังหวัด"))
        with patch.object(q, "load_catalog", lambda _: cat):
            q.validate_query(config)
            mdx = q.build_mdx(config, 1, 20)
            self.assertIn("CrossJoin(", mdx)
            self.assertIn(other, mdx)
            self.assertIn(",1,1)", mdx.replace(" ", ""))
            self.assertIn(",20,20)", mdx.replace(" ", ""))

    def test_real_cellordinal_semantics(self):
        response = q.parse_xmla(sample_xml())
        self.assertEqual(len(response["axes"]["Axis0"]), 2)
        self.assertEqual(len(response["axes"]["Axis1"]), 2)
        self.assertEqual(len(response["cells"]), 4)
        records = list(q.excel_records(q.default_query("MV"), response))
        self.assertEqual([r["VALUE"] for r in records], ["10", "11", "12", "13"])
        self.assertEqual(records[0]["LAST_DATA_UPDATE"], "2026-10-18T04:39:34")
        self.assertEqual(records[0]["Y"][0]["caption"], "Type 0")
        self.assertEqual(records[0]["MEASURE"], "[Measures].[ACT POLICY CNT]")

    def test_empty_and_invalid(self):
        response = q.parse_xmla(sample_xml(x_count=0))
        self.assertEqual(response["axes"]["Axis0"], [])
        self.assertEqual(list(q.excel_records(q.default_query("MV"), response)), [])
        with self.assertRaises(ET.ParseError):
            q.parse_xmla(b"not xml")
        with self.assertRaises(ValueError):
            q.parse_xmla(b"<Envelope><Fault>Denied</Fault></Envelope>")


if __name__ == "__main__":
    unittest.main()
