"""Where the Universal Sheet comes from: Google Sheets (main) or a local .xlsx (offline fallback and tests).

Both give the same thing - the worksheet's title and its rows as lists of cell values - so the parsing and checks in
crest/universal.py are identical for either source.

    GoogleSheet  reads the whole worksheet in one request and writes AI results back in one request, through a
                 Google Cloud service account (.env: UNIVERSAL_SHEET_ID, UNIVERSAL_WORKSHEET,
                 GOOGLE_CREDENTIALS_FILE). GOOGLEFINANCE cells still showing "Loading..." are read once more after
                 a short wait.
    LocalSheet   reads templates/universal-sheet.xlsx (values saved in the file). It is never written.

UniversalSheet wraps a source for one batch of runs: it reads once, saves a snapshot of the data under
runs/<period>/snapshots/, and when Google cannot be reached it uses the latest snapshot instead (the run is marked
WARN).
"""
from __future__ import annotations

import gzip
import hashlib
import json
import re
import time
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

from openpyxl.utils import column_index_from_string

from .util import ROOT, PipelineError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
API = "https://sheets.googleapis.com/v4/spreadsheets"
DEFAULT_WORKSHEET = "Universal"
DEFAULT_CREDENTIALS = "secrets/google-service-account.json"
LOADING = "loading..."


class SheetUnreachable(Exception):
    """Google could not be reached (network, timeout, Google-side error): the latest snapshot can stand in."""


class ReadOnlySheet(Exception):
    """The source cannot be written (local file, or a snapshot standing in for Google)."""


def _cell(v):
    """JSON-safe cell value (a local .xlsx can hold dates)."""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d/%m/%Y")
    return v


# --------------------------------------------------------------------------- sources
class LocalSheet:
    """templates/universal-sheet.xlsx - read only."""
    live = False

    def __init__(self, path):
        self.path = Path(path)
        self.label = f"local file {self.path.name}"

    def read(self):
        from .universal import read_xlsx_rows
        if not self.path.exists():
            raise PipelineError(f"Universal Sheet not found: {self.path}")
        return read_xlsx_rows(self.path)

    def write(self, updates, expect=None):
        raise ReadOnlySheet(f"the {self.label} is not written (Google Sheets is the source of truth)")


class GoogleSheet:
    """One worksheet of a Google Sheet, through the Sheets API v4 and a service account."""
    live = True

    def __init__(self, sheet_id: str, worksheet: str = DEFAULT_WORKSHEET, credentials_file=None, *,
                 session=None, sleep=time.sleep, loading_wait: float = 15.0, timeout: float = 60.0):
        self.sheet_id = sheet_id.strip()
        self.worksheet = worksheet or DEFAULT_WORKSHEET
        self.credentials_file = Path(credentials_file or ROOT / DEFAULT_CREDENTIALS)
        if not self.credentials_file.is_absolute():
            self.credentials_file = ROOT / self.credentials_file
        self._session = session
        self.sleep = sleep
        self.loading_wait = loading_wait
        self.timeout = timeout
        self.label = f"Google Sheet {self.sheet_id[:8]}... / {self.worksheet}"
        self.notes: list[str] = []

    # ---- HTTP
    def session(self):
        if self._session is None:
            if not self.credentials_file.exists():
                raise PipelineError(
                    f"Google credentials file not found: {self.credentials_file}. Save the service-account key "
                    "there (see README: 'Google Sheets setup'), or set GOOGLE_CREDENTIALS_FILE in .env.")
            try:
                from google.auth.transport.requests import AuthorizedSession
                from google.oauth2.service_account import Credentials
            except ImportError as e:
                import sys
                raise PipelineError(
                    f"A package needed to read the Google Sheet is missing for the Python running this tool "
                    f"({sys.executable}): {e}. Install the packages into that Python with this command, then run "
                    f"again:   {Path(sys.executable).name} -m pip install -r requirements.txt") from e
            try:
                creds = Credentials.from_service_account_file(str(self.credentials_file), scopes=SCOPES)
            except (ValueError, KeyError) as e:
                raise PipelineError(f"{self.credentials_file.name} is not a service-account key file ({e}).") from e
            self.service_email = getattr(creds, "service_account_email", "")
            self._session = AuthorizedSession(creds)
        return self._session

    def _call(self, method: str, url: str, **kw):
        try:
            r = getattr(self.session(), method)(url, timeout=self.timeout, **kw)
        except PipelineError:
            raise
        except Exception as e:  # noqa: BLE001 - connection refused, DNS, timeout, token refresh failure
            raise SheetUnreachable(f"{type(e).__name__}: {e}") from e
        if r.status_code in (401, 403, 404):
            who = getattr(self, "service_email", "") or "the service account's email"
            raise PipelineError(
                f"Google refused access to the Universal Sheet (HTTP {r.status_code}). Check that UNIVERSAL_SHEET_ID "
                f"is right and that the sheet is shared with {who} as Editor. Google said: {_err_text(r)}")
        if r.status_code == 400:
            raise PipelineError(f"Google Sheets rejected the request (HTTP 400): {_err_text(r)}. Check "
                                f"UNIVERSAL_WORKSHEET ('{self.worksheet}') names a tab of the sheet.")
        if r.status_code >= 300:
            raise SheetUnreachable(f"HTTP {r.status_code}: {_err_text(r)}")
        return r.json()

    def _a1(self, rng: str = "") -> str:
        ws = "'" + self.worksheet.replace("'", "''") + "'"
        return f"{ws}!{rng}" if rng else ws

    # ---- read
    def _get_rows(self) -> list[list]:
        url = (f"{API}/{self.sheet_id}/values/{quote(self._a1(), safe='')}"
               "?valueRenderOption=UNFORMATTED_VALUE&dateTimeRenderOption=FORMATTED_STRING&majorDimension=ROWS")
        # Google sends a blank cell as "" where a local .xlsx gives None: make them the same
        return [[None if c == "" else c for c in r] for r in self._call("get", url).get("values", [])]

    def read(self):
        rows = self._get_rows()
        n = _loading_cells(rows)
        if n:
            self.notes.append(f"{n} cell(s) still showed 'Loading...'; read the sheet again after {self.loading_wait:g}s")
            self.sleep(self.loading_wait)
            rows = self._get_rows()
            n = _loading_cells(rows)
            if n:
                self.notes.append(f"WARN: {n} cell(s) still 'Loading...' after the second read; those prices count "
                                  "as missing")
        return self.worksheet, rows

    # ---- write
    def write(self, updates: dict, expect: dict | None = None) -> list[int]:
        """updates: {row number: {column letter: value}}, written in one request (values as typed, RAW).
        expect: {row number: (column letter, value)} - before writing, these cells are read again and a row whose
        cell no longer holds that value (rows inserted or sorted in the sheet since it was read) is skipped.
        Returns the row numbers written."""
        if not updates:
            return []
        if expect:
            rows = sorted(expect)
            rngs = "&".join("ranges=" + quote(self._a1(f"{expect[r][0]}{r}"), safe="") for r in rows)
            got = self._call("get", f"{API}/{self.sheet_id}/values:batchGet?{rngs}&valueRenderOption=UNFORMATTED_VALUE")
            vr = got.get("valueRanges", [])
            moved = []
            for r, v in zip(rows, vr):
                cur = (v.get("values") or [[None]])[0]
                cur = cur[0] if cur else None
                if _same(cur, expect[r][1]):
                    continue
                moved.append(r)
                updates.pop(r, None)
            if moved:
                self.notes.append(f"WARN: rows {moved} of the Universal Sheet changed since it was read (rows inserted "
                                  "or sorted?), so they were not written; the next run writes them")
            if not updates:
                return []
        self._ensure_columns(max(column_index_from_string(c) for u in updates.values() for c in u))
        data = [{"range": self._a1(f"{c}{r}"), "values": [[v]]} for r in sorted(updates) for c, v in updates[r].items()]
        self._call("post", f"{API}/{self.sheet_id}/values:batchUpdate",
                   json={"valueInputOption": "RAW", "data": data})
        return sorted(updates)

    def _ensure_columns(self, need: int):
        """A new column (e.g. Rating Source the first time) may lie beyond the sheet's grid: widen it first."""
        meta = self._call("get", f"{API}/{self.sheet_id}?fields=sheets.properties")
        props = next((s["properties"] for s in meta.get("sheets", [])
                      if s.get("properties", {}).get("title") == self.worksheet), None)
        if props is None:
            raise PipelineError(f"No tab named '{self.worksheet}' in the Google Sheet (UNIVERSAL_WORKSHEET).")
        have = props.get("gridProperties", {}).get("columnCount", need)
        if need > have:
            self._call("post", f"{API}/{self.sheet_id}:batchUpdate", json={"requests": [{"appendDimension": {
                "sheetId": props["sheetId"], "dimension": "COLUMNS", "length": need - have}}]})


def _err_text(r) -> str:
    try:
        return r.json().get("error", {}).get("message", "")[:300]
    except Exception:  # noqa: BLE001
        return (getattr(r, "text", "") or "")[:300]


def _loading_cells(rows) -> int:
    return sum(1 for r in rows for c in r if isinstance(c, str) and c.strip().lower() == LOADING)


def _same(a, b) -> bool:
    norm = lambda x: re.sub(r"\s+", " ", str("" if x is None else x)).strip().lower()  # noqa: E731
    return norm(a) == norm(b)


# --------------------------------------------------------------------------- one batch of runs
@dataclass
class SheetData:
    title: str
    rows: list
    source: str            # label of where the rows came from
    live: bool             # True: read from Google now; False: local file or snapshot
    read_at: str
    snapshot: Path | None = None
    warning: str | None = None


class UniversalSheet:
    """The Universal Sheet for a batch of runs: read once (every client and both steps see the same data),
    snapshot every run's data, fall back to the latest snapshot when Google is unreachable."""

    def __init__(self, source, runs_dir=None):
        self.source = source
        self.runs_dir = Path(runs_dir) if runs_dir else ROOT / "runs"
        self._data: SheetData | None = None

    @property
    def label(self) -> str:
        return self.source.label

    def data(self) -> SheetData:
        if self._data is None:
            now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            try:
                title, rows = self.source.read()
                self._data = SheetData(title, [[_cell(c) for c in r] for r in rows], self.source.label,
                                       self.source.live, now)
            except SheetUnreachable as e:
                snap = latest_snapshot(self.runs_dir)
                if snap is None:
                    raise PipelineError(f"Could not reach the Google Sheet ({e}) and there is no saved snapshot under "
                                        f"{self.runs_dir} to use instead. Check the internet connection and run again.")
                d = json.loads(gzip.decompress(snap.read_bytes()).decode("utf-8"))
                self._data = SheetData(d["title"], d["rows"], f"snapshot {snap.name} (of {d['source']}, read "
                                       f"{d['read_at']})", False, d["read_at"], snap,
                                       f"Could not reach the Google Sheet ({e}); used the latest snapshot {snap.name}, "
                                       f"read {d['read_at']}. Prices and results may be out of date.")
        return self._data

    def notes(self) -> list[str]:
        return list(getattr(self.source, "notes", []))

    def save_snapshot(self, period_compact: str) -> Path:
        """Save the data this run used under runs/<period>/snapshots/ (once per distinct content) and return it."""
        d = self.data()
        if d.snapshot is not None:
            return d.snapshot
        body = json.dumps({"title": d.title, "source": d.source, "read_at": d.read_at, "rows": d.rows},
                          ensure_ascii=False, default=str)
        digest = hashlib.sha1(json.dumps(d.rows, default=str).encode("utf-8")).hexdigest()[:10]
        folder = self.runs_dir / period_compact / "snapshots"
        folder.mkdir(parents=True, exist_ok=True)
        same = sorted(folder.glob(f"universal-*-{digest}.json.gz"))
        if same:
            d.snapshot = same[-1]
        else:
            d.snapshot = folder / f"universal-{datetime.now():%Y%m%d-%H%M%S}-{digest}.json.gz"
            d.snapshot.write_bytes(gzip.compress(body.encode("utf-8")))
        return d.snapshot

    def write(self, updates: dict, expect: dict | None = None) -> list[int]:
        if not self.data().live:
            raise ReadOnlySheet(f"this run used {self.data().source}, which is not written")
        return self.source.write(updates, expect)


def latest_snapshot(runs_dir: Path) -> Path | None:
    snaps = sorted(Path(runs_dir).glob("*/snapshots/universal-*.json.gz"), key=lambda p: p.name)
    return snaps[-1] if snaps else None


def open_universal(env: dict, local_path=None, runs_dir=None, **google_kw) -> UniversalSheet:
    """The Google Sheet when .env has UNIVERSAL_SHEET_ID (and no local file was asked for), else the local file."""
    sid = (env.get("UNIVERSAL_SHEET_ID") or "").strip()
    if local_path is None and sid:
        src = GoogleSheet(sid, env.get("UNIVERSAL_WORKSHEET") or DEFAULT_WORKSHEET,
                          env.get("GOOGLE_CREDENTIALS_FILE") or None, **google_kw)
    else:
        src = LocalSheet(local_path or ROOT / "templates" / "universal-sheet.xlsx")
    return UniversalSheet(src, runs_dir)

