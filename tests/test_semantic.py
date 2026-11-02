import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "examples"))
from make_sample import create
from datadict.metadata import read_excel
from datadict.search import search
from datadict.relations import infer
from datadict.sql_history import extract


class SemanticTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.path = Path(self.temp.name) / "sample.xlsx"
        create(self.path)
        self.model = read_excel(self.path)

    def test_metadata_and_field_search(self):
        self.assertEqual(len(self.model["tables"]), 3)
        self.assertEqual(search(self.model, "采购数量")[0]["table"], "t_pm_purorderentry")
        self.assertEqual(search(self.model, "SUPPLIER")[0]["table"], "t_bd_supplier")
        self.assertEqual(search(self.model, "不存在xyz"), [])
        self.assertEqual(search(self.model, "vendor", {"synonyms": {"vendor": ["supplier"]}})[0]["table"], "t_bd_supplier")

    def test_relations_have_evidence_and_no_entry_pk_guess(self):
        edges = infer(self.model)
        self.assertTrue(any(e["source_column"] == "fsupplierid" and e["target_table"] == "t_bd_supplier" for e in edges))
        self.assertTrue(any(e["source_column"] == "fid" and e["source_table"].endswith("entry") for e in edges))
        self.assertFalse(any(e["source_column"] == "fentryid" for e in edges))
        self.assertTrue(all(e["evidence"] and 0 <= e["confidence"] <= 1 for e in edges))
        self.assertEqual(infer(self.model, {"min_confidence": .99}), [])

    def test_parent_uses_single_primary_key_not_nonunique_fid(self):
        parent = next(t for t in self.model['tables'] if t['name'].endswith('entry'))
        parent['columns'].append(dict(name='fid', type='bigint', primary_key=False))
        child = dict(name=parent['name']+'_r', parent=parent['name'], columns=[
            dict(name='fid', type='bigint', primary_key=False),
            dict(name='fentryid', type='bigint', primary_key=True)])
        self.model['tables'].append(child)
        matches=[r for r in infer(self.model) if r['source_table']==child['name']]
        self.assertTrue(any(r['source_column']=='fentryid' and r['target_column']=='fentryid' for r in matches))
        self.assertFalse(any(r['source_column']=='fid' and r['target_column']=='fid' for r in matches))

    def test_missing_columns_types_and_composite_key(self):
        supplier = self.model["tables"][2]
        supplier["columns"][0]["type"] = "varchar"
        self.assertFalse(any(e["source_column"] == "fsupplierid" for e in infer(self.model)))
        supplier["columns"][0]["type"] = "bigint"
        supplier["columns"][1]["primary_key"] = True
        self.assertFalse(any(e["source_column"] == "fsupplierid" for e in infer(self.model)))
        with self.assertRaises(ValueError):
            infer(self.model, {"relationships": [dict(source_table="missing", source_column="id", target_table="x", target_column="id")]})

    def test_sql_scopes_comments_or_and_functions(self):
        sql = "SELECT * FROM t_pm_purorder a LEFT JOIN t_bd_supplier b ON a.fsupplierid=b.fid AND b.fname='a.x=b.x' -- JOIN fake c ON c.x=a.x"
        obs, notes = extract(sql, "report.sql")
        self.assertEqual(len(obs), 1)
        self.assertEqual(obs[0]["target_table"], "t_bd_supplier")
        self.assertEqual(notes, [])
        for unsafe in ["SELECT * FROM a JOIN b ON a.id=b.id OR a.x=b.x", "WITH a AS (SELECT * FROM b) SELECT * FROM a JOIN c ON a.id=c.id", "SELECT * FROM a JOIN (SELECT * FROM b) c ON a.id=c.id", "SELECT * FROM a JOIN b ON coalesce(a.id=b.id, true)"]:
            self.assertEqual(extract(unsafe)[0], [])
        self.assertEqual(len(infer(self.model, observations=obs)), 2)
        with self.assertRaises(Exception):
            extract("SELECT * FROM (")

    def test_directory_and_failure_exit(self):
        create(self.path.parent / "second.xlsx")
        run = subprocess.run([sys.executable, str(ROOT / "scripts/main.py"), str(self.path.parent)], capture_output=True)
        self.assertEqual(run.returncode, 0, run.stderr)
        self.assertTrue((self.path.parent / "second_graph/index.json").exists())
        self.assertTrue((self.path.parent / "sample_graph/index.json").exists())
        self.path.write_bytes(b"not a workbook")
        run = subprocess.run([sys.executable, str(ROOT / "scripts/main.py"), str(self.path)], capture_output=True)
        self.assertNotEqual(run.returncode, 0)
        self.assertNotIn(b"Done", run.stdout)

    def test_all_columns_retained(self):
        import openpyxl
        wb = openpyxl.load_workbook(self.path)
        for i in range(3, 64):
            wb["7"].append([None, i, f"fextra{i}", f"Extra {i}", "bigint"])
        wb.save(self.path)
        wb.close()
        self.assertEqual(len(read_excel(self.path)["tables"][0]["columns"]), 63)
        run = subprocess.run([sys.executable, str(ROOT / "scripts/generate_search_graph.py"), str(self.path)], capture_output=True)
        self.assertEqual(run.returncode, 0, run.stderr)
        columns = json.loads((self.path.parent / "columns/t_pm_purorder.json").read_text(encoding="utf-8"))
        self.assertEqual(len(columns), 63)

    def test_cli_and_legacy_outputs(self):
        run = subprocess.run([sys.executable, str(ROOT / "scripts/generate_all.py"), str(self.path)], capture_output=True, text=True)
        self.assertEqual(run.returncode, 0, run.stderr)
        base = self.path.parent
        self.assertIn('comment on column "t_pm_purorder"."fsupplierid"', (base / "sample_comments.sql").read_text(encoding="utf-8"))
        self.assertIn("erDiagram", (base / "sample_datadict_card.md").read_text(encoding="utf-8"))
        self.assertTrue((base / "sample_graph/graph.html").exists())
        self.assertEqual(len(json.loads((base / "sample_graph/index.json").read_text(encoding="utf-8"))), 3)
        model = base / "sample_semantic.json"
        run = subprocess.run([sys.executable, str(ROOT / "scripts/semantic.py"), "context", str(model), "t_pm_purorder"], capture_output=True, text=True, encoding="utf-8")
        self.assertEqual(run.returncode, 0, run.stderr)
        result = json.loads(run.stdout)
        self.assertIn("t_bd_supplier", result["candidate_neighbors"])
        self.assertEqual(len(result["tables"]), 1)


if __name__ == "__main__":
    unittest.main()
