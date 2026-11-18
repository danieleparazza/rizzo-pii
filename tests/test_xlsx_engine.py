# -*- coding: utf-8 -*-
"""Motore Excel (src/app/xlsx_engine.py) con un rilevatore finto a regex.

Serve openpyxl: in CI (nessun pip install) si salta. Tutti i dati sono SINTETICI.
"""
import datetime as dt
import os
import re
import sys
import tempfile
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src", "app"))
from registry import Registry  # noqa: E402

try:
    from openpyxl import Workbook, load_workbook
    import xlsx_engine
    from xlsx_engine import Cancelled, CsvError, cell_text, default_plan, process_file, sniff
except ImportError:                      # CI senza openpyxl
    xlsx_engine = None

A = {"mode": "anonymize"}
P = {"mode": "preserve"}

_FAKE = [
    ("EMAIL", re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")),
    ("DATE", re.compile(r"\d{2}/\d{2}/\d{4}")),
    ("TELEPHONENUM", re.compile(r"\b3\d{9}\b")),
    ("FULLNAME", re.compile(r"Mario Rossi|Anna Bianchi")),
]
CALLS = []


def fake_detect(texts, excluded):
    CALLS.append(list(texts))
    out = []
    for t in texts:
        spans = []
        for lab, rx in _FAKE:
            if lab not in excluded:
                spans += [(m.start(), m.end(), lab) for m in rx.finditer(t)]
        out.append(sorted(spans))
    return out


@unittest.skipIf(xlsx_engine is None, "openpyxl non installato")
class XlsxEngine(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = self._tmp.name
        CALLS.clear()

    def tearDown(self):
        self._tmp.cleanup()

    def book(self, sheets, name="in.xlsx"):
        """sheets: {titolo: [righe]} -> path di un .xlsx."""
        wb = Workbook()
        wb.remove(wb.active)
        for title, rows in sheets.items():
            ws = wb.create_sheet(title)
            for r in rows:
                ws.append(r)
        p = os.path.join(self.dir, name)
        wb.save(p)
        return p

    def read(self, p):
        wb = load_workbook(p)
        try:
            return {ws.title: [list(r) for r in ws.iter_rows(values_only=True)]
                    for ws in wb.worksheets}
        finally:
            wb.close()

    def run_(self, src, plan=None, **kw):
        out = os.path.join(self.dir, "out.xlsx")
        d = sniff(src)
        reg = kw.pop("registry", None) or Registry()
        rep = process_file(src, out, plan or default_plan(d), registry=reg,
                           detect_fn=kw.pop("detect_fn", fake_detect), dialect=d, **kw)
        return out, rep, reg

    # ------------------------------------------------------------------ sniff
    def test_sniff_fogli_intestazione_anteprima(self):
        src = self.book({"Clienti": [["Nome", "Email", None, "Note"],
                                     ["Mario Rossi", "m@example.com", 3, None]],
                         "Vuoto": []})
        d = sniff(src, n_preview=5)
        self.assertEqual(d["format"], "xlsx")
        c, v = d["sheets"]
        self.assertEqual(c["name"], "Clienti")
        self.assertEqual(c["header"], ["Nome", "Email", "C", "Note"])
        self.assertEqual(c["preview"], [["Mario Rossi", "m@example.com", "3", ""]])
        self.assertEqual(v["header"], [])
        self.assertEqual(default_plan(d), [[A, A, A, A], []])

    def test_file_non_excel(self):
        p = os.path.join(self.dir, "x.xlsx")
        with open(p, "wb") as f:
            f.write(b"Nome,Email\nMario,a@b.it\n")
        with self.assertRaises(CsvError) as cm:
            sniff(p)
        self.assertNotIn("Mario", str(cm.exception))

    # ---------------------------------------------------------------- process
    def test_anonimizza_e_conserva_tipi(self):
        when = dt.datetime(2024, 1, 5)
        src = self.book({"Foglio": [
            ["Nome", "Telefono", "Quando", "Importo", "Attivo", "Codice"],
            ["Mario Rossi", 3331234567, when, 12.5, True, "Anna Bianchi"],
            ["Anna Bianchi", 42, dt.date(2023, 2, 1), 7, False, "x"],
        ]})
        out, rep, reg = self.run_(src, [[A, A, A, A, A, P]])
        rows = self.read(out)["Foglio"]
        self.assertEqual(rows[0], ["Nome", "Telefono", "Quando", "Importo", "Attivo", "Codice"])
        self.assertEqual(rows[1][:2], ["[FULLNAME_1]", "[TELEPHONENUM_1]"])
        self.assertEqual(rows[1][2], "[DATE_1]")
        self.assertEqual(rows[1][3:], [12.5, True, "Anna Bianchi"])     # preserve intatta
        self.assertEqual(rows[2][:2], ["[FULLNAME_2]", 42])              # numero senza PII: resta int
        self.assertEqual(rows[2][3:5], [7, False])
        self.assertEqual(reg.mapping["[DATE_1]"], "05/01/2024")
        self.assertEqual(rep["rows_in"], 2)
        self.assertEqual(rep["by_label"]["FULLNAME"], 2)

    def test_pseudonimi_coerenti_tra_fogli_e_file(self):
        reg = Registry()
        a = self.book({"A": [["n"], ["Mario Rossi"]], "B": [["n"], ["Mario Rossi"]]}, "a.xlsx")
        b = self.book({"C": [["n"], ["Mario Rossi scrive"]]}, "b.xlsx")
        out_a, _, _ = self.run_(a, registry=reg)
        ra = self.read(out_a)
        out_b, _, _ = self.run_(b, registry=reg)
        self.assertEqual(ra["A"][1], ["[FULLNAME_1]"])
        self.assertEqual(ra["B"][1], ["[FULLNAME_1]"])
        self.assertEqual(self.read(out_b)["C"][1], ["[FULLNAME_1] scrive"])

    def test_tag_esclusi(self):
        src = self.book({"S": [["a"], ["Mario Rossi, m@example.com"]]})
        out, _, _ = self.run_(src, excluded={"EMAIL"})
        self.assertEqual(self.read(out)["S"][1], ["[FULLNAME_1], m@example.com"])

    def test_testo_che_inizia_con_uguale_resta_testo(self):
        src = self.book({"S": [["a"], ["x"]]})
        wb = load_workbook(src)
        wb["S"]["A2"].value = "=1+1"
        wb["S"]["A2"].data_type = "s"
        wb.save(src)
        out, _, _ = self.run_(src, [[P]])
        wb = load_workbook(out)
        self.assertEqual(wb["S"]["A2"].data_type, "s")
        self.assertEqual(wb["S"]["A2"].value, "=1+1")

    def test_formula_diventa_valore(self):
        src = self.book({"S": [["a", "b"], ["Mario", '=A2&" Rossi"']]})
        out, _, _ = self.run_(src, [[P, P]])
        # openpyxl non calcola: nessun valore in cache -> cella vuota, la formula sparisce
        self.assertEqual(self.read(out)["S"][1], ["Mario", None])

    def test_nome_foglio_con_pii(self):
        src = self.book({"Mario Rossi": [["a"], ["ciao"]], "Dati": [["a"], ["x"]]})
        out, rep, _ = self.run_(src)
        self.assertEqual(list(self.read(out)), ["Foglio 1", "Dati"])
        self.assertEqual(rep["sheets_renamed"], 1)
        # foglio tutto preserve: nome lasciato com'e'
        out, _, _ = self.run_(src, [[P], [P]], detect_fn=None)
        self.assertEqual(list(self.read(out)), ["Mario Rossi", "Dati"])

    def test_celle_oltre_intestazione_anonimizzate(self):
        # oltre le righe d'anteprima: sniff non le vede, il piano copre solo "a"
        src = self.book({"S": [["a"]] + [["x"]] * 30 + [["x", "Mario Rossi"]]})
        out, _, _ = self.run_(src, [[P]])
        self.assertEqual(self.read(out)["S"][31], ["x", "[FULLNAME_1]"])

    def test_piano_non_valido(self):
        src = self.book({"S": [["a", "b"], ["x", "y"]]})
        for plan in ([[A]], [[A, A], [A]], [[A, {"mode": "regex"}]]):
            with self.assertRaises(CsvError):
                self.run_(src, plan)
        self.assertFalse(os.path.exists(os.path.join(self.dir, "out.xlsx.partial")))

    def test_annullato_niente_output(self):
        src = self.book({"S": [["a"]] + [["Mario Rossi %d" % i] for i in range(50)]})
        with self.assertRaises(Cancelled):
            self.run_(src, cancel=lambda: True, batch_size=10)
        self.assertEqual(os.listdir(self.dir), ["in.xlsx"])

    def test_batch_e_avanzamento(self):
        src = self.book({"S": [["a"]] + [["Mario Rossi %d" % i] for i in range(25)]})
        seen = []
        _, rep, _ = self.run_(src, batch_size=10, progress=lambda r, d, t: seen.append((r, d, t)))
        self.assertEqual(rep["rows_out"], 25)
        self.assertEqual(len(CALLS), 4)            # 3 flush + il nome del foglio
        self.assertEqual(seen[-1][0], 25)
        self.assertEqual(seen[-1][1], seen[-1][2])
        self.assertTrue(all(d <= t for _, d, t in seen))

    def test_niente_file_nella_temp_di_sistema(self):
        """openpyxl write_only scriverebbe i fogli in tempfile.gettempdir(): mai, nemmeno
        se il job si interrompe a meta'."""
        import tempfile as tf
        sysdir = os.path.join(self.dir, "systemp")
        os.mkdir(sysdir)
        old, tf.tempdir = tf.tempdir, sysdir
        try:
            src = self.book({"S": [["a"]] + [["Mario Rossi %d" % i] for i in range(30)]})
            self.run_(src, batch_size=10)
            calls = []
            with self.assertRaises(Cancelled):
                self.run_(src, batch_size=10, cancel=lambda: calls.append(1) or len(calls) > 2)
        finally:
            tf.tempdir = old
        self.assertEqual(os.listdir(sysdir), [])
        self.assertEqual(sorted(os.listdir(self.dir)), ["in.xlsx", "out.xlsx", "systemp"])

    def test_cell_text(self):
        self.assertEqual(cell_text(3331234567.0), "3331234567")
        self.assertEqual(cell_text(1.5), "1.5")
        self.assertEqual(cell_text(dt.datetime(2024, 3, 2, 14, 5)), "02/03/2024 14:05")
        self.assertEqual(cell_text(dt.time(9, 30)), "09:30")
        self.assertIsNone(cell_text(True))
        self.assertIsNone(cell_text(None))


if __name__ == "__main__":
    unittest.main()
