# -*- coding: utf-8 -*-
"""Job su file Excel (.xlsx) attraverso il Blueprint di csv_jobs, motore vero.

Stesso impianto di test_csv_jobs (Flask minima, rilevatore finto). Serve anche
openpyxl: in CI senza pip install si salta. Tutti i valori sono SINTETICI.
"""
import io
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tests"))

from test_csv_jobs import CSV_A, _Base, csv_jobs  # noqa: E402

try:
    from openpyxl import Workbook, load_workbook
    import csv_engine
    import xlsx_engine
except ImportError:
    xlsx_engine = None


def _book(sheets):
    wb = Workbook()
    wb.remove(wb.active)
    for title, rows in sheets.items():
        ws = wb.create_sheet(title)
        for r in rows:
            ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


BOOK = None if xlsx_engine is None else _book({
    "Clienti": [["nome", "email", "telefono"],
                ["Mario", "mario.rossi@example.com", "+39 333 1234567"],
                ["Anna", "anna.bianchi@example.com", None]],
    "Note": [["testo"], ["scrivere a mario.rossi@example.com"]],
})


@unittest.skipIf(csv_jobs is None or xlsx_engine is None, "Flask o openpyxl non installati")
class XlsxJobs(_Base):
    engine = csv_engine if xlsx_engine is not None else None

    def upload_bytes(self, name, data, chunk=1000):
        r = self.c.post("/csv/uploads", json={"name": name, "size": len(data)})
        self.assertEqual(r.status_code, 200, r.get_json())
        uid = r.get_json()["upload_id"]
        for off in range(0, len(data), chunk):
            r = self.c.put(f"/csv/uploads/{uid}?offset={off}", data=data[off:off + chunk],
                           content_type="application/octet-stream")
            self.assertEqual(r.status_code, 200, r.get_json())
        r = self.c.post(f"/csv/uploads/{uid}/finalize", json={})
        return r

    def test_finalize_fogli(self):
        d = self.upload_bytes("rubrica.xlsx", BOOK).get_json()
        self.assertEqual(d["format"], "xlsx")
        self.assertEqual([s["name"] for s in d["sheets"]], ["Clienti", "Note"])
        self.assertEqual(d["sheets"][0]["header"], ["nome", "email", "telefono"])
        self.assertEqual(len(d["plan"]), 2)
        self.assertNotIn("header", d)

    def test_job_misto_csv_e_xlsx(self):
        x = self.upload_bytes("rubrica.xlsx", BOOK).get_json()
        c = self.upload("clienti.csv", CSV_A)
        x["plan"][0][0] = {"mode": "preserve"}
        jid = self.start([x, c])
        st = self.wait(jid)
        self.assertEqual(st["state"], "done", st)
        self.assertEqual(st["files"][0]["out_name"], "rubrica.anon.xlsx")
        r = self.c.get(f"/csv/jobs/{jid}/files/0")
        self.assertIn("spreadsheetml", r.mimetype)
        wb = load_workbook(io.BytesIO(r.get_data()))
        r.close()
        rows = {ws.title: [list(v) for v in ws.iter_rows(values_only=True)] for ws in wb.worksheets}
        self.assertEqual(rows["Clienti"][1], ["Mario", "[EMAIL_1]", "[TELEFONO_1]"])
        self.assertEqual(rows["Note"][1], ["scrivere a [EMAIL_1]"])
        # stesso pseudonimo anche nel CSV dello stesso job
        r = self.c.get(f"/csv/jobs/{jid}/files/1")
        self.assertIn("[EMAIL_1]", r.get_data(as_text=True).splitlines()[1])
        r.close()
        p = self.c.get(f"/csv/jobs/{jid}/files/0/preview").get_json()
        self.assertEqual(p["sheets"][0]["preview"][0][1], "[EMAIL_1]")
        mp = json.loads(self.c.get(f"/csv/jobs/{jid}/dictionary").get_data(as_text=True))
        self.assertEqual(mp["[EMAIL_1]"], "mario.rossi@example.com")

    def test_piano_sbagliato(self):
        x = self.upload_bytes("rubrica.xlsx", BOOK).get_json()
        r = self.c.post("/csv/jobs", json={"files": [{"upload_id": x["upload_id"],
                                                      "plan": [{"mode": "anonymize"}] * 3}]})
        self.assertEqual(r.status_code, 400)

    def test_xlsx_finto(self):
        r = self.upload_bytes("finto.xlsx", b"nome,email\nMario,a@b.it\n")
        self.assertEqual(r.status_code, 422)
        self.assertNotIn("Mario", r.get_json()["error"])

    def test_senza_openpyxl(self):
        import flask
        a = flask.Flask("t2")
        bp = csv_jobs.create_blueprint(lambda: self.prefs, lambda: None, engine=csv_engine,
                                       root=self.root, xlsx_engine=False)
        a.register_blueprint(bp)
        r = a.test_client().post("/csv/uploads", json={"name": "a.xlsx", "size": 10})
        self.assertEqual(r.status_code, 501)
        bp.csv_state["instance"].close()


if __name__ == "__main__":
    unittest.main()
