"""Website API (Phase 5): a thin FastAPI layer over the pipeline, for the Vue front end in web/.

    GET  /api/status                  Universal Sheet source, AI / search keys, LibreOffice, queue length
    GET  /api/template                the blank holdings template
    POST /api/runs                    upload a file (form fields: file, step = portfolio | deck | both)
    GET  /api/runs                    run history, newest first
    GET  /api/runs/{id}               one run: status, live stage, and when finished its checks, exceptions,
                                      AI-sourced values and files
    GET  /api/runs/{id}/events        the same, pushed as server-sent events while the run works
    GET  /api/runs/{id}/files/{name}  download a result file
    /                                 the built front end (web/dist), when present

Runs are queued and processed one at a time: LibreOffice and the Universal Sheet write-back must not run twice at
once. Every run lives in runs/_web/<id>/ (the upload, job.json and a copy of its result files), so history
downloads never change when the same client is run again.
"""
from __future__ import annotations

import asyncio
import json
import queue
import re
import shutil
import threading
import traceback
import uuid
from collections import Counter
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from .narrative_api import resolve_ai
from .office import find_soffice
from .util import ROOT, PipelineError, load_env, load_json, CONFIG_DIR

RUNS_DIR = ROOT / "runs"
JOBS_DIR = RUNS_DIR / "_web"
OUTPUT_DIR = ROOT / "output"
HOLDINGS_TEMPLATE = ROOT / "templates" / "Holdings_Template.xlsx"
DECK_TEMPLATE = ROOT / "templates" / "Deck_-_template.pptx"
WEB_DIST = ROOT / "web" / "dist"
MAX_UPLOAD = 20 * 1024 * 1024

STAGES = {"sheet": "Reading the Universal Sheet", "match": "Matching holdings", "research": "AI research",
          "excel": "Building the Portfolio Excel", "check": "Checking the Excel formulas",
          "ai": "Ratings and narrative", "deck": "Building the deck", "qa": "Quality checks and PDF"}
STEP_STAGES = {"portfolio": ["sheet", "match", "research", "excel", "check"],
               "deck": ["sheet", "match", "research", "ai", "deck", "qa"],
               "both": ["sheet", "match", "research", "excel", "check", "ai", "deck", "qa"]}
STEP_LABELS = {"portfolio": "Portfolio Excel", "deck": "Deck", "both": "Excel and deck"}


def _now() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _safe_name(name: str) -> str:
    name = Path(name or "upload.xlsx").name
    return re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", name).strip() or "upload.xlsx"


# --------------------------------------------------------------------------- jobs
class Jobs:
    """Job records (in memory, mirrored to runs/_web/<id>/job.json) and the single worker thread."""

    def __init__(self, jobs_dir: Path, runner):
        self.dir = Path(jobs_dir)
        self.runner = runner              # runner(job, progress) -> result dict; raises PipelineError
        self.lock = threading.Lock()
        self.changed = threading.Condition(self.lock)
        self.jobs: dict[str, dict] = {}
        self.q: queue.Queue = queue.Queue()
        self._load()
        self.worker = threading.Thread(target=self._work, daemon=True, name="crest-run-worker")
        self.worker.start()

    def _load(self):
        if not self.dir.exists():
            return
        for f in self.dir.glob("*/job.json"):
            try:
                j = json.loads(f.read_text(encoding="utf-8"))
            except (OSError, json.JSONDecodeError):
                continue
            if j.get("status") in ("queued", "running"):        # the server stopped mid-run
                j.update(status="failed", error="The server was restarted before this run finished. Run it again.",
                         finished=j.get("finished") or _now())
            self.jobs[j["id"]] = j

    def _save(self, j: dict):
        d = self.dir / j["id"]
        d.mkdir(parents=True, exist_ok=True)
        (d / "job.json").write_text(json.dumps(j, indent=2, ensure_ascii=False, default=str), encoding="utf-8")

    def _update(self, jid: str, **kw):
        with self.changed:
            j = self.jobs[jid]
            j.update(kw)
            j["version"] = j.get("version", 0) + 1
            self._save(j)
            self.changed.notify_all()

    def add(self, filename: str, data: bytes, step: str) -> dict:
        jid = datetime.now().strftime("%Y%m%d-%H%M%S-") + uuid.uuid4().hex[:6]
        d = self.dir / jid
        d.mkdir(parents=True, exist_ok=True)
        (d / filename).write_bytes(data)
        j = {"id": jid, "step": step, "step_label": STEP_LABELS[step], "filename": filename, "status": "queued",
             "stages": [{"key": k, "label": STAGES[k]} for k in STEP_STAGES[step]], "stage": None,
             "created": _now(), "started": None, "finished": None, "client": None, "period": None,
             "result": None, "error": None, "version": 0}
        with self.changed:
            self.jobs[jid] = j
            self._save(j)
        self.q.put(jid)
        return self.get(jid)

    def get(self, jid: str) -> dict | None:
        with self.lock:
            j = self.jobs.get(jid)
            if j is None:
                return None
            out = json.loads(json.dumps(j, default=str))
            if j["status"] == "queued":
                waiting = [k for k, v in self.jobs.items() if v["status"] == "queued"]
                out["queue_position"] = sorted(waiting).index(jid) + 1 + self.running()
            return out

    def running(self) -> int:
        return sum(1 for v in self.jobs.values() if v["status"] == "running")

    def list(self, limit: int = 100) -> list[dict]:
        with self.lock:
            rows = sorted(self.jobs.values(), key=lambda j: j["id"], reverse=True)[:limit]
            return [{k: j.get(k) for k in ("id", "step", "step_label", "filename", "status", "client", "period",
                                            "created", "finished", "error")} |
                    {"run_status": (j.get("result") or {}).get("status"),
                     "files": [f["name"] for f in (j.get("result") or {}).get("files", [])]} for j in rows]

    def wait_change(self, jid: str, version: int, timeout: float) -> bool:
        with self.changed:
            return self.changed.wait_for(lambda: self.jobs[jid].get("version", 0) != version, timeout=timeout)

    def queued(self) -> int:
        with self.lock:
            return sum(1 for v in self.jobs.values() if v["status"] in ("queued", "running"))

    def _work(self):
        while True:
            jid = self.q.get()
            self._update(jid, status="running", started=_now(), stage=None)
            try:
                result = self.runner(self.jobs[jid], self.dir / jid,
                                     lambda key, jid=jid: self._update(jid, stage=key))
                self._update(jid, status="done", finished=_now(), stage=None, result=result,
                             client=result.get("client"), period=result.get("period"))
            except PipelineError as e:
                self._update(jid, status="failed", finished=_now(), error=str(e))
            except Exception as e:  # noqa: BLE001 - a crash must not stop the worker
                traceback.print_exc()
                self._update(jid, status="failed", finished=_now(),
                             error=f"Unexpected error: {type(e).__name__}: {e}")


# --------------------------------------------------------------------------- one pipeline run
def detect_period(universal) -> str:
    from .universal import result_periods
    c = Counter(str(v or "").replace(" ", "").upper() for v in result_periods(universal))
    c.pop("", None)
    if not c:
        raise PipelineError("The Universal Sheet has no 'Q Result Period' values, so the period is unknown. "
                            "Fill that column (e.g. Q2 FY27).")
    return c.most_common(1)[0][0]


def run_pipeline(job: dict, folder: Path, progress) -> dict:
    """The real runner: one pipeline run for an uploaded file. Returns the summary the front end shows."""
    from .pipeline import run
    from .sheet_source import open_universal
    progress("sheet")
    us = open_universal(load_env(), runs_dir=RUNS_DIR)
    period = detect_period(us)
    res = run(period=period, universal=us, holdings=folder / job["filename"], template=DECK_TEMPLATE,
              runs_dir=RUNS_DIR, publish_dir=OUTPUT_DIR, step=job["step"], progress=progress)
    return summarise(res, folder)


def summarise(res, folder: Path) -> dict:
    """RunResult -> JSON for the front end; result files are copied into the job folder."""
    files_dir = folder / "files"
    files_dir.mkdir(exist_ok=True)
    files = []
    kinds = {"excel": "Portfolio Excel", "deck": "Deck (PowerPoint)", "pdf": "Deck (PDF)", "ai_ratings": "AI ratings"}
    for key, label in kinds.items():
        p = res.outputs.get(key)
        if p and Path(p).exists():
            shutil.copy2(p, files_dir / Path(p).name)
            files.append({"kind": key, "label": label, "name": Path(p).name, "size": Path(p).stat().st_size})
    if res.run_dir and (res.run_dir / "report.md").exists():
        shutil.copy2(res.run_dir / "report.md", files_dir / "report.md")
        files.append({"kind": "report", "label": "Run report", "name": "report.md",
                      "size": (files_dir / "report.md").stat().st_size})
    i = res.info
    checks = [{"name": c.name, "status": c.status, "detail": c.detail, "step": c.step,
               "items": [str(x) if not isinstance(x, dict) else ", ".join(f"{k}: {v}" for k, v in x.items())
                         for x in (c.items or [])][:50]} for c in res.checks]
    ar = i.get("ai_research") or {}
    return {
        "status": res.status, "client": i.get("client"), "period": i.get("period"), "asof": i.get("asof"),
        "holdings_kind": i.get("holdings_kind"), "universal": i.get("universal"),
        "universal_read_at": i.get("universal_read_at"),
        "matched": sum((i.get("matched_by") or {}).values()), "holdings": i.get("holdings_parsed"),
        "checks": checks,
        "exceptions": [{"severity": e.severity, "kind": e.kind, "symbol": e.symbol, "detail": e.detail}
                       for e in res.exceptions],
        "notes": list(res.notes),
        "ai_sourced": i.get("ai_sourced") or [],
        "ai_research": {"requests": ar.get("requests", 0), "searches": ar.get("searches", 0),
                        "filled": ar.get("filled", 0),
                        "skipped": [{"symbol": s, "reason": w} for s, w in ar.get("skipped", [])]},
        "narrative_pending": list((res.narrative or {}).get("pending_fields") or []),
        "files": files,
    }


# --------------------------------------------------------------------------- the app
def create_app(runner=run_pipeline, jobs_dir: Path = JOBS_DIR, serve_web: bool = True) -> FastAPI:
    app = FastAPI(title="CREST Portfolio Review", docs_url="/api/docs", openapi_url="/api/openapi.json")
    jobs = Jobs(jobs_dir, runner)
    app.state.jobs = jobs

    @app.get("/api/status")
    def status():
        env = load_env()
        sid = (env.get("UNIVERSAL_SHEET_ID") or "").strip()
        settings = load_json(CONFIG_DIR / "settings.json")
        return {"universal": "Google Sheet" if sid else "Local file (templates/universal-sheet.xlsx)",
                "google": bool(sid), "ai": resolve_ai(env), "web_search": bool(env.get("TAVILY_API_KEY")),
                "ai_research": settings.get("ai_research", "auto"),
                "libreoffice": bool(find_soffice(settings.get("libreoffice_path", ""))),
                "deck_template": DECK_TEMPLATE.exists(), "queue": jobs.queued()}

    @app.get("/api/template")
    def template():
        if not HOLDINGS_TEMPLATE.exists():
            raise HTTPException(404, "The holdings template is missing: run python tools/build_holdings_template.py")
        return FileResponse(HOLDINGS_TEMPLATE, filename="CREST_Holdings_Template.xlsx")

    @app.post("/api/runs", status_code=201)
    async def start(file: UploadFile = File(...), step: str = Form("both")):
        if step not in STEP_STAGES:
            raise HTTPException(422, "Choose Portfolio, Deck or Both.")
        name = _safe_name(file.filename)
        if Path(name).suffix.lower() not in (".xlsx", ".xlsm"):
            raise HTTPException(422, "Upload an Excel file (.xlsx): a filled holdings template or a Portfolio Excel.")
        data = await file.read(MAX_UPLOAD + 1)
        if len(data) > MAX_UPLOAD:
            raise HTTPException(413, "That file is larger than 20 MB, which is too big for a holdings file.")
        if not data:
            raise HTTPException(422, "That file is empty.")
        return jobs.add(name, data, step)

    @app.get("/api/runs")
    def history(limit: int = 100):
        return jobs.list(limit)

    @app.get("/api/runs/{jid}")
    def one(jid: str):
        j = jobs.get(jid)
        if j is None:
            raise HTTPException(404, "No such run.")
        return j

    @app.get("/api/runs/{jid}/events")
    async def events(jid: str):
        if jobs.get(jid) is None:
            raise HTTPException(404, "No such run.")

        async def stream():
            version = -1
            while True:
                j = jobs.get(jid)
                if j["version"] != version:
                    version = j["version"]
                    yield f"data: {json.dumps(j, ensure_ascii=False)}\n\n"
                    if j["status"] in ("done", "failed"):
                        return
                changed = await asyncio.to_thread(jobs.wait_change, jid, version, 15.0)
                if not changed:
                    yield ": keep-alive\n\n"
        return StreamingResponse(stream(), media_type="text/event-stream",
                                 headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

    @app.get("/api/runs/{jid}/files/{name}")
    def download(jid: str, name: str):
        j = jobs.get(jid)
        if j is None or not j.get("result"):
            raise HTTPException(404, "No such run, or it has no files yet.")
        if name not in {f["name"] for f in j["result"]["files"]}:
            raise HTTPException(404, "No such file in this run.")
        return FileResponse(jobs.dir / jid / "files" / name, filename=name)

    if serve_web and WEB_DIST.exists():
        app.mount("/assets", StaticFiles(directory=WEB_DIST / "assets"), name="assets")

        @app.get("/{path:path}", include_in_schema=False)
        def spa(path: str):
            f = WEB_DIST / path
            if path and f.is_file() and WEB_DIST in f.resolve().parents:
                return FileResponse(f)
            return FileResponse(WEB_DIST / "index.html")

    return app
