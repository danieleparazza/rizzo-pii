"""Anonimizzazione di cartelle di lavoro Excel (.xlsx) in streaming.

Stessa API di csv_engine (sniff / default_plan / process_file / MODES / CsvError /
Cancelled), cosi' csv_jobs tratta i due formati allo stesso modo, e lo stesso cuore:
`csv_engine.RowBatcher` (batch al rilevatore, cache, registro degli pseudonimi). Il
rilevatore entra solo da `detect_fn`, come per i CSV.

Differenze rispetto al CSV:
- Piu' fogli: il dialetto ha `sheets` = [{"name", "header", "preview", "rows"}] e il piano
  e' una lista di piani, uno per foglio ([[{"mode": ...}, ...], ...]). La prima riga di
  ogni foglio e' l'intestazione e resta com'e', come nei CSV.
- Celle tipizzate: si analizzano anche numeri e date, convertiti in testo (date come
  gg/mm/aaaa). Se non vi si trova nulla la cella resta identica, col suo tipo; se si trova
  una PII diventa testo con lo pseudonimo. I booleani non si analizzano.
- Formule: si legge il valore calcolato che Excel ha salvato nel file (data_only) e si
  scrive quello. La formula sparisce: potrebbe contenere essa stessa la PII
  (="Mario "&B2). Un file mai ricalcolato (generato da script) puo' non averne: la
  cella esce vuota.
- Output = soli valori: niente formattazione, larghezze, celle unite, commenti,
  immagini, grafici, nomi definiti, intestazioni/pie' di pagina o proprieta' del
  documento (autore...). E' il prezzo dello streaming (openpyxl write_only) ed e' anche
  piu' sicuro: in quei posti le PII sopravviverebbero.
- Nomi dei fogli: se il motore vi trova una PII (in un foglio con almeno una colonna
  anonymize) il foglio diventa "Foglio N". Un placeholder `[TAG_N]` non e' un nome di
  foglio valido (le parentesi quadre sono vietate).
- Il testo di una cella che inizia con "=" resta testo: non diventa una formula.

File temporanei: in write_only openpyxl scrive ogni foglio in un file temporaneo prima di
comprimerlo, e di suo lo mette nella cartella temp di sistema, dove resterebbe se il job
viene annullato o il processo muore. Qui lo si mette accanto all'output
(`<out>.sheets/`, dentro la cartella del job, spazzata all'avvio) e lo si cancella
comunque a fine run.

Memoria: openpyxl in read_only legge i fogli in streaming, ma la tabella delle stringhe
condivise del file sta tutta in RAM (un .xlsx ha al massimo ~1M righe per foglio).

Nessun valore delle celle in log, eccezioni o report.
"""
import datetime as dt
import decimal
import os
import shutil
import time
import zipfile

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.worksheet._writer import WorksheetWriter

from csv_engine import BATCH_ROWS, MODES, Cancelled, CsvError, RowBatcher, _clean  # noqa: F401

SHEET_TITLE_MAX = 31


# --------------------------------------------------------------------------- celle

def cell_text(value):
    """Testo da analizzare di una cella Excel, o None se non va analizzata."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, str):
        return value
    if isinstance(value, dt.datetime):
        if (value.hour, value.minute, value.second) == (0, 0, 0):
            return value.strftime("%d/%m/%Y")
        return value.strftime("%d/%m/%Y %H:%M" + (":%S" if value.second else ""))
    if isinstance(value, dt.date):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, dt.time):
        return value.strftime("%H:%M" + (":%S" if value.second else ""))
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() and abs(value) < 1e16 else repr(value)
    if isinstance(value, (int, decimal.Decimal)):
        return str(value)
    return None          # timedelta, errori (#N/A sono gia' stringhe) ecc.


def _display(value):
    t = cell_text(value)
    if t is not None:
        return t
    return "" if value is None else str(value)


# --------------------------------------------------------------------------- lettura

def _open(path):
    try:
        return load_workbook(path, read_only=True, data_only=True, keep_links=False)
    except (InvalidFileException, zipfile.BadZipFile, KeyError, OSError, ValueError):
        raise CsvError("file Excel non leggibile: serve un .xlsx (non .xls), "
                       "senza password") from None


def _header(first, width):
    """Etichette dell'intestazione per l'UI: valore della cella o lettera di colonna."""
    out = []
    for i in range(width):
        v = first[i] if i < len(first) else None
        t = _display(v).strip()
        out.append(t or get_column_letter(i + 1))
    return out


def _trim(row):
    """Toglie le celle vuote in coda (read_only allinea le righe a max_column)."""
    n = len(row)
    while n and row[n - 1] is None:
        n -= 1
    return row[:n]


def sniff(path, n_preview=20):
    """Fogli, intestazioni e prime righe senza leggere tutto il file.

    Ritorna {"format": "xlsx", "sheets": [{"name", "header": [str], "preview": [[str]],
             "rows": int | None (righe dichiarate dal file, intestazione compresa)}]}."""
    if os.path.getsize(path) == 0:
        raise CsvError("il file e' vuoto")
    wb = _open(path)
    try:
        sheets = []
        for ws in wb.worksheets:
            it = ws.iter_rows(values_only=True)
            first = _trim(list(next(it, ()) or ()))
            preview = []
            for row in it:
                if len(preview) >= n_preview:
                    break
                preview.append(list(row))
            width = max([len(first)] + [len(_trim(r)) for r in preview])
            sheets.append({
                "name": ws.title,
                "header": _header(first, width),
                "preview": [[_display(v) for v in r[:width]] + [""] * (width - len(r[:width]))
                            for r in preview],
                "rows": ws.max_row,
            })
    finally:
        wb.close()
    if not sheets:
        raise CsvError("la cartella di lavoro non contiene fogli")
    return {"format": "xlsx", "sheets": sheets}


# --------------------------------------------------------------------------- piano

def default_plan(dialect):
    """Un piano per foglio, tutte le colonne anonymize (come per i CSV)."""
    return [[{"mode": "anonymize"} for _ in s["header"]] for s in dialect["sheets"]]


def check_plan(plan, dialect):
    """Messaggio d'errore (senza valori) o None."""
    sheets = dialect["sheets"]
    if not isinstance(plan, list) or len(plan) != len(sheets):
        return "Il piano deve avere una voce per foglio."
    for p, s in zip(plan, sheets):
        if not isinstance(p, list) or len(p) != len(s["header"]):
            return "Il piano deve avere una voce per colonna."
        for c in p:
            if not isinstance(c, dict) or c.get("mode") not in MODES:
                return "Modalità di colonna non valida."
    return None


# --------------------------------------------------------------------------- scrittura

def _out_cell(ws, value):
    if isinstance(value, str) and value.startswith("="):
        c = WriteOnlyCell(ws, value=value)
        c.data_type = "s"           # testo, non formula
        return c
    return value


def _safe_title(title, idx, detect_fn, excluded, used):
    """Titolo del foglio in uscita: quello originale, oppure "Foglio N" se contiene PII."""
    if detect_fn is not None and title.strip():
        spans = _clean(detect_fn([title], set(excluded))[0], len(title), set(excluded))
        if spans:
            title = f"Foglio {idx + 1}"
    base, n = title[:SHEET_TITLE_MAX], 2
    while base.lower() in used:
        suffix = f" ({n})"
        base, n = title[:SHEET_TITLE_MAX - len(suffix)] + suffix, n + 1
    used.add(base.lower())
    return base


class _SheetWriter(WorksheetWriter):
    """Writer di un foglio write_only su un file scelto da noi (non nella temp di sistema)."""

    def cleanup(self):
        try:
            os.remove(self.out)
        except OSError:
            pass


def _new_sheet(book, title, tmpdir, idx):
    ws = book.create_sheet(title)
    # come WriteOnlyWorksheet._get_writer, ma col file temporaneo in tmpdir
    ws._writer = _SheetWriter(ws, out=os.path.join(tmpdir, f"sheet{idx}.xml"))
    ws._writer.write_top()
    return ws


def _abort_sheets(book):
    """Chiude i fogli lasciati a meta' (altrimenti i generatori di openpyxl scrivono su
    file gia' chiusi quando vengono raccolti)."""
    for ws in book.worksheets:
        try:
            ws.close()
        except Exception:  # noqa: BLE001 - si sta gia' gestendo un errore
            pass


# --------------------------------------------------------------------------- process

def process_file(in_path, out_path, plan, *, registry, detect_fn, excluded=(),
                 dialect=None, progress=None, cancel=None, strict=False,
                 batch_size=BATCH_ROWS):
    """Anonimizza `in_path` (.xlsx) -> `out_path` (.xlsx, via .partial + rename).

    plan: [[{"mode": "anonymize"|"preserve"}, ...], ...] un piano per foglio.
    Gli altri argomenti come csv_engine.process_file (strict non ha effetto: in un
    .xlsx non ci sono righe malformate). progress(rows, done, total) riporta l'avanzamento
    in byte stimati dalle righe, nella stessa unita' dei CSV.

    Ritorna {"rows_in", "rows_out", "rows_skipped": 0, "skipped_lines": [], "sheets",
             "sheets_renamed", "by_label", "seconds"}."""
    t0 = time.monotonic()
    excluded = set(excluded or ())
    if dialect is None:
        dialect = sniff(in_path)
    msg = check_plan(plan, dialect)
    if msg:
        raise CsvError(msg)
    if detect_fn is None and any(c["mode"] == "anonymize" for p in plan for c in p):
        raise CsvError("le colonne da anonimizzare richiedono il motore di rilevamento "
                       "(modello non disponibile)")
    size = os.path.getsize(in_path)
    rows_total = sum(max(0, (s.get("rows") or 0) - 1) for s in dialect["sheets"])
    partial = out_path + ".partial"
    tmpdir = out_path + ".sheets"
    rep = {"rows_in": 0, "rows_out": 0, "rows_skipped": 0, "skipped_lines": [],
           "sheets": 0, "sheets_renamed": 0}
    counts = {}

    def report_progress():
        if progress is not None:
            done = int(size * rep["rows_in"] / rows_total) if rows_total else 0
            progress(rep["rows_in"], min(done, size - 1), size)

    wb = _open(in_path)
    out = Workbook(write_only=True)
    saved = False
    try:
        os.makedirs(tmpdir, exist_ok=True)
        used = set()
        for idx, (ws, sheet_plan) in enumerate(zip(wb.worksheets, plan)):
            modes = [c["mode"] for c in sheet_plan]
            anon = "anonymize" in modes
            title = _safe_title(ws.title, idx, detect_fn if anon else None, excluded, used)
            rep["sheets_renamed"] += title != ws.title[:SHEET_TITLE_MAX]
            ws_out = _new_sheet(out, title, tmpdir, idx)
            rep["sheets"] += 1

            def emit(row, ws_out=ws_out):
                ws_out.append([_out_cell(ws_out, v) for v in row])
                rep["rows_out"] += 1

            rows = RowBatcher(registry, detect_fn, excluded, batch_size, emit,
                              cancel=cancel, on_flush=report_progress, to_text=cell_text)
            it = ws.iter_rows(values_only=True)
            first = next(it, None)
            if first is not None:
                ws_out.append([_out_cell(ws_out, v) for v in _trim(list(first))])
            # celle oltre l'intestazione: si anonimizzano (nel dubbio, si protegge)
            extra = "anonymize" if detect_fn is not None else "preserve"
            for row in it:
                row = _trim(list(row))
                rep["rows_in"] += 1
                rows.add(row, modes + [extra] * (len(row) - len(modes)))
            rows.flush()
            for k, v in rows.counts.items():
                counts[k] = counts.get(k, 0) + v
        if cancel is not None and cancel():
            raise Cancelled()
        out.save(partial)
        saved = True
        os.replace(partial, out_path)
    except BaseException:
        if not saved:
            _abort_sheets(out)
        try:
            os.remove(partial)
        except OSError:
            pass
        raise
    finally:
        wb.close()
        shutil.rmtree(tmpdir, ignore_errors=True)

    if progress is not None:
        progress(rep["rows_in"], size, size)
    rep["by_label"] = counts
    rep["seconds"] = round(time.monotonic() - t0, 3)
    return rep
