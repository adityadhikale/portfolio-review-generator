r"""CREST Portfolio Review website.

    python serve.py                 then open http://localhost:8000
    python serve.py --port 8080     another port
    python serve.py --host 0.0.0.0  reachable from other machines on the office network

The front end must be built once (and after every change to web/):  cd web  then  npm install  and  npm run build
"""
from __future__ import annotations

import argparse
import os
import sys

import uvicorn

from crest.api import WEB_DIST, create_app
from crest.util import hosted


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(errors="replace")
    ap = argparse.ArgumentParser(description="CREST Portfolio Review website")
    # a host sets PORT (and HOST, or CREST_HOSTED=1 to listen on every address)
    ap.add_argument("--host", default=os.environ.get("HOST") or ("0.0.0.0" if hosted() else "127.0.0.1"))
    ap.add_argument("--port", type=int, default=int(os.environ.get("PORT") or 8000))
    a = ap.parse_args(argv)
    if not WEB_DIST.exists():
        print("The front end is not built yet: in the web folder run  npm install  then  npm run build. "
              "The API still runs (http://localhost:%d/api/docs)." % a.port)
    print(f"CREST Portfolio Review: http://localhost:{a.port}")
    uvicorn.run(create_app(), host=a.host, port=a.port, log_level="warning")
    return 0


if __name__ == "__main__":
    sys.exit(main())
