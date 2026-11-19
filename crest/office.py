"""LibreOffice (headless) helpers: recalculate a workbook, convert PPTX -> PDF. Local only, no network."""
from __future__ import annotations

import os
import shutil
import subprocess
import tempfile
from pathlib import Path

from .util import PipelineError

_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/usr/bin/soffice", "/usr/local/bin/soffice",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]

_MACRO = """<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">
Sub RecalcAndSave()
  ThisComponent.calculateAll()
  ThisComponent.store()
  ThisComponent.close(True)
End Sub
</script:module>"""


def find_soffice(configured: str = "") -> str | None:
    for c in [configured, shutil.which("soffice"), shutil.which("soffice.exe"), *_CANDIDATES]:
        if c and Path(c).exists():
            return str(c)
    return None


def _run(cmd, timeout):
    env = dict(os.environ)
    if os.name != "nt":
        env.setdefault("SAL_USE_VCLPLUGIN", "svp")
    return subprocess.run(cmd, capture_output=True, text=True, timeout=timeout, env=env)


def _profile(work: Path, soffice: str, timeout: int) -> str:
    prof = work / "profile"
    prof.mkdir()
    url = prof.as_uri()
    _run([soffice, "--headless", "--terminate_after_init", f"-env:UserInstallation={url}"], timeout)
    macro_dir = prof / "user" / "basic" / "Standard"
    if not macro_dir.exists():
        raise PipelineError("LibreOffice did not create a usable profile; cannot recalculate.")
    (macro_dir / "Module1.xba").write_text(_MACRO, encoding="utf-8")
    return url


def recalculate_copy(xlsx: Path, out_copy: Path, soffice: str, timeout: int = 120) -> Path:
    """Write a LibreOffice-recalculated copy of `xlsx` to `out_copy` (formulas kept, values cached)."""
    with tempfile.TemporaryDirectory(prefix="crest-lo-") as td:
        work = Path(td)
        url = _profile(work, soffice, timeout)
        tmp = work / xlsx.name
        shutil.copyfile(xlsx, tmp)
        before = tmp.stat().st_mtime_ns
        res = _run([soffice, "--headless", "--norestore", f"-env:UserInstallation={url}",
                    "vnd.sun.star.script:Standard.Module1.RecalcAndSave?language=Basic&location=application",
                    str(tmp)], timeout)
        if tmp.stat().st_mtime_ns == before:
            raise PipelineError(f"LibreOffice did not save the recalculated workbook. stderr: {res.stderr[-500:]}")
        out_copy.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(tmp, out_copy)
    return out_copy


def convert(src: Path, fmt: str, outdir: Path, soffice: str, timeout: int = 180) -> Path:
    with tempfile.TemporaryDirectory(prefix="crest-lo-") as td:
        work = Path(td)
        url = (work / "profile").as_uri()
        outdir.mkdir(parents=True, exist_ok=True)
        res = _run([soffice, "--headless", "--norestore", f"-env:UserInstallation={url}",
                    "--convert-to", fmt, "--outdir", str(outdir), str(src)], timeout)
    out = outdir / (src.stem + "." + fmt.split(":")[0])
    if not out.exists():
        raise PipelineError(f"LibreOffice could not convert {src.name} to {fmt}. stderr: {res.stderr[-500:]}")
    return out
