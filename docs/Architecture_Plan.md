# CREST Portfolio Review Generator: Architecture Plan

Version 1, 20/11/2026. Status: **approved, not started**.

## 1. Why the architecture is changing

Today the Universal Sheet is a local Excel file (`templates\universal-sheet.xlsx`). Its GOOGLEFINANCE prices only
refresh when someone opens the file in Google Sheets and saves it again. Between saves, every Portfolio Excel,
deck and PDF the tool builds uses the prices from that last save. Results, commentary and sector go out of date
too, and staff have to research and type them by hand.

The goal is a **core tool, used by CREST staff through a website**, where:

- The **Universal Sheet lives in Google Sheets** and is the single source of truth. GOOGLEFINANCE keeps price and
  market cap live. The app reads the sheet, and writes AI results back to it, through the Google Sheets API.
- Staff convert each client's holdings into **one standard holdings template** before uploading it.
- The work is **two separate steps**, so a deck can be made without making a new Excel:
  - **Step A, Portfolio:** holdings template in, Portfolio Excel out.
  - **Step B, Deck:** a Portfolio Excel (or a holdings template) in, PPT and PDF out.
- **AI fills Sector, Result Line and Management Commentary** where they are blank or out of date, and Rating and
  Sentiment as today. **Views stay CREST's own** and are never generated.
- Login and cloud deployment come **later**. This plan builds the core tool first.

## 2. The staff process this supports

1. The client sends holdings as an Excel file or a Google Sheet.
2. Staff copy the holdings into the standard holdings template.
3. Staff upload the template on the website and download the Portfolio Excel (Step A).
4. Staff upload that Portfolio Excel and download the deck as PPT and PDF (Step B).
5. The deck and the Portfolio Excel are shared with the client (the Excel through Google Drive, where
   GOOGLEFINANCE goes live).

## 3. Decisions

| Topic | Decision |
|---|---|
| Source of truth | Universal Sheet on Google Sheets, read and written through a Google Cloud service account (CREST supplies the credentials) |
| Holdings input | **Only** the standard template. Raw broker files are no longer read directly |
| Client name and as-on date | Two labelled cells at the top of the template |
| Template columns | ISIN, Symbol, Company Name, Quantity, Average Price. No price column: prices come live from the Universal Sheet |
| Steps | A and B are separate actions on the website. A third action, "Both", runs them together |
| AI | Small batches, saved and reused, with a limit per run. This replaces the old "at most one request per run" rule. Web search: a free Tavily key, read by the free narrative-API key (the free key cannot search itself); the narrative API's own search stays available for a paid key (decided 21/11/2026) |
| Website | A Python API (FastAPI) around the pipeline, and a Vue 3 front end with shadcn-vue components, Motion for Vue animation and Tailwind, in CREST's colours. Not Streamlit (decided 21/11/2026: too plain for staff use) |
| Typed values | Anything CREST types in the Universal Sheet always wins over AI |
| GOOGLEFINANCE lag and gaps | Acceptable for client output |
| Users | CREST staff only. Login is added later |
| Unchanged rules | No AI attribution anywhere. No client data in git. Staff commit and push themselves |

## 4. How the new tool fits together

| Part | Role |
|---|---|
| Standard holdings template | The one input format for client holdings |
| Universal Sheet (Google Sheets) | Live prices and market cap; CREST's Views; sector, results, commentary, ratings |
| Holdings template reader | Reads and checks the template; stops with the exact cell if something is wrong |
| Universal Sheet connector | Reads the whole sheet in one request; writes AI results back; saves a snapshot of every run's data |
| AI research | Fills blank or outdated Sector, Result Line and Commentary for the stocks being run |
| AI review | Rating and Sentiment for unrated stocks, and the deck narrative (as today) |
| Step A, Portfolio | Matches holdings to the Universal Sheet, runs the checks, builds the Portfolio Excel |
| Step B, Deck | Reads holdings from a Portfolio Excel or a template, takes fresh data from the Universal Sheet, builds PPT and PDF |
| Website API (FastAPI) | Takes uploads, runs Step A / Step B / both one at a time, streams progress, serves the results |
| Website front end (Vue) | Upload, live progress, results and review screens, run history |
| Command line | `python run.py` still works: every template in `input\` goes through Step A and Step B |

Flow: **template → Step A → Portfolio Excel → Step B → PPT + PDF**, with both steps reading the live Universal
Sheet. A "Both" run uses one snapshot of the sheet for A and B, so the Excel and the deck show identical numbers.

## 5. Data formats

### 5.1 Standard holdings template

A new file, `templates\Holdings_Template.xlsx`, kept in git (it holds no client data).

- Cell `B2` reads "Client Name"; staff type the name in `C2`.
- Cell `B3` reads "As On Date"; staff type the date in `C3` (dd/mm/yyyy, or a real Excel date).
- Row 5 holds the headers: `ISIN | Symbol | Company Name | Quantity | Average Price`. Holdings start on row 6 and
  end at the first blank row.
- Every row needs Quantity above 0, Average Price above 0, and an ISIN or a Symbol. Company Name is optional and
  is used only as a last resort for matching.
- Rows for the same stock are merged with a weighted average price.
- Any problem (missing column, text where a number should be, blank client name or date) stops the run and names
  the exact cell.
- A second sheet in the file explains how to fill it.

### 5.2 Portfolio Excel

The layout stays as it is today:

- Portfolio sheet: summary at the top, headers on row 8, holdings from row 9, hidden Key column A.
- Market Cap and CMP are GOOGLEFINANCE formulas that go live in Google Sheets; Ticker and Views use INDEX/MATCH.
- Other sheets: Raw_Holdings, Holdings_Normalized, Universal_Snapshot, Exceptions, and References (added
  23/11/2026: AI-written values with their source links, and AI ratings with their reasons).

Two additions so Step B can read it back reliably:

- A hidden marker cell ("CREST Portfolio v1") so Step B can tell a Portfolio Excel from a holdings template.
- Step B finds the Portfolio columns **by header name**, not by position. Staff edits still read correctly:
  deleted or added rows, re-ordered columns, or a copy downloaded again from Google Drive.

### 5.3 Universal Sheet on Google Sheets

The same columns as today, plus:

- **Data Source** and **Data Updated**, written by the tool when AI fills Sector, Result Line or Commentary. They
  record the model, period, date and source links, in the same way Rating Source and Rating Reason do today.
- Cells CREST types have a blank Data Source and are never overwritten.
- **AI Searched** (added 23/11/2026): "quarter · date · why" when a stock's results were searched but are not out
  yet, so the stock is not searched again for a few days.

## 6. Phases

Each phase ends with all tests passing, the README updated, and commit commands for CREST to run.

### Phase 1: Standard holdings template

Needs no credentials.

- New template reader. It produces the same holding records the rest of the tool already uses, so matching, the
  Excel build and the deck keep working unchanged.
- New `templates\Holdings_Template.xlsx`, built by a small script so it can be re-created exactly.
- Removed: the broker auto-detect code and its configuration, the `--broker` option, taking the client name from
  the file name, and taking the as-on date from the file's saved date.
- Removed: the "broker price" option. The template has no price column, so prices always come from the Universal
  Sheet.
- Optional: a command-line helper that turns a raw broker export into the template, to help staff during the
  switch-over. It is not on the website and can be dropped if not wanted.

### Phase 2: Split into Step A and Step B

Needs no credentials.

- The main run file (about 600 lines) is split into three parts:
  - **Shared setup:** settings, period, loading the Universal Sheet, matching, applying ratings, deck groups,
    the exceptions list, the run archive.
  - **Step A:** checks, the Portfolio Excel, the LibreOffice cross-check copy, the validation gate.
  - **Step B:** reads holdings from a Portfolio Excel or a template; takes fresh price, market cap, sector, views,
    results, commentary, rating and sentiment from the Universal Sheet; builds the deck data; narrative; fills
    the deck; quality checks; PDF.
- **Prices in Step B** are always the live Universal Sheet values at deck time.
- **Hand-edited Excel:** quantities and average prices in the uploaded Portfolio sheet win. A row with no Key is
  matched by Ticker, then by Name. Rows that cannot be matched are listed in the report.
- `python run.py` still works. A new option, `--step portfolio|deck|both`, defaults to both.
- The run report gets a section for each step.

### Phase 3: Universal Sheet on Google Sheets

Needs the Google credentials.

- A new connector with two sources behind one interface:
  - **Google Sheet** (main): reads the whole sheet in one request and writes back in one request.
  - **Local file** (fallback and tests): today's Excel reader.
- The existing sheet parsing and checks (column names, tickers, ETF detection, duplicates, ISIN lookup) are
  reused unchanged for both sources.
- **Live values:** cells still showing "Loading..." are read again once after a short wait. Remaining `#N/A` or
  blank prices go to Exceptions as "Missing price", as today.
- **Snapshot:** every run saves the data it used, with a timestamp, under `runs\<period>\snapshots\`, and the
  report names it. If Google cannot be reached, the latest snapshot is used and the run is marked WARN.
- **Write-back** of AI ratings goes through the Google Sheets API. The current workaround that edits the Excel
  file's XML directly (with local backups and the "file is open in Excel" handling) is removed.
- Period detection reads the same data.
- Settings: `.env` gets the sheet ID, the worksheet name and the path to the credentials file. The credentials
  file lives in a `secrets\` folder that git ignores.
- **CREST does:** create the Google Cloud project and service account, turn on the Sheets API, upload the
  Universal Sheet to Google Sheets, share it with the service account's email as Editor, and save the key file
  locally. Step-by-step instructions are given at the start of this phase.

### Phase 4: AI fills Sector, Result Line and Commentary

Needs the paid narrative-API key.

- Runs before ratings, and only for the **stocks held in the portfolio being run** where:
  - Sector is blank, or
  - Result Line or Commentary is blank, or
  - the result is from an older quarter than the run's period.
- Cells CREST typed are never touched.
- Batches of about 10 stocks per request, with a maximum number of requests per run (a setting, default 3). The
  API searches the web; every answer must give the quarter, the figures and the source links.
- Sector must be one of the sectors the deck already knows, so it maps to a deck group. Result Line must follow
  the house format the rating rules read, or it is rejected.
- Answers are saved under `runs\<period>\` and written back to the Universal Sheet with Data Source and Data
  Updated. The same stock and quarter is never asked twice.
- Without web search (free key), only Sector is filled; results stay "Results awaited".
- Rating, Sentiment and the narrative keep the current flow, counted inside the same per-run limit.
- The run report lists every AI-sourced result under "AI-sourced: review before sending".

### Phase 5: Internal website

Runs locally first.

Two parts, in one repository: the Python pipeline stays exactly as it is, with a thin API in front of it, and a
separate Vue front end talks to that API.

- **API (FastAPI, `crest/api.py`):**
  - Download the blank holdings template.
  - Start a run: upload a file and choose Portfolio, Deck or Both. Runs go into a queue and run one at a time,
    because LibreOffice and the Universal Sheet write-back must not run twice at once; a waiting run shows its
    place in the queue.
  - Live progress while a run works (reading the sheet, AI research, ratings, Excel, deck, PDF), sent to the
    page as it happens.
  - The finished run: status, checks by step, exceptions, the AI-sourced values to review, notes, and download
    links for the Excel, PPT, PDF and report.
  - Run history, and the Universal Sheet connection status (Google live, or the snapshot in use).
  - The pipeline gains one small hook to report which stage it is on; nothing else in it changes.
- **Front end (Vue 3 + Vite + TypeScript, `web/`):** shadcn-vue components, Motion for Vue animation, Tailwind
  styling, in CREST's colours and fonts.
  - **New run:** three cards (Portfolio, Deck, Both), a drag-and-drop upload area, and the blank template download.
  - **Run in progress:** an animated step-by-step timeline with the current stage and elapsed time.
  - **Results:** a status banner, download buttons, checks grouped by step, the exceptions table, and the
    AI-sourced review table with source links.
  - **History:** past runs by client and period, with their status and files.
- Each upload runs in its own folder under `runs\`. No client data leaves the server except the trimmed,
  anonymous narrative input already sent today.
- **Starting it:** one command, `python serve.py`, serves the API and the built front end at
  `http://localhost:8000`. For front-end work, `npm run dev` gives live reload against the same API.
- Node.js is needed only to build the front end; the built files are plain HTML, CSS and JavaScript.

### Phase 6: Free hosting

Planned 23/11/2026. Goal: the website online for CREST staff, on a host that **never asks for a card**.
Progress: steps 1, 2 and 6 built 23/11/2026; step 0 (Render sign-up), step 5 and step 7 to do.

**Decisions (23/11/2026)**

| Topic | Decision |
|---|---|
| Host | **Render, Free web service** (one instance, 512 MB memory, 0.1 CPU, 750 free hours a month). Region: Singapore |
| Not used | Google Cloud Run, Oracle, Fly.io, Koyeb (ask for a card); Vercel (free plan is non-commercial only, and requests are time-limited); Netlify (no Python); a VM |
| LibreOffice | **Not on the host.** The hosted site delivers the Portfolio Excel and the PPT; staff make the PDF in PowerPoint (File, Save as PDF). The two checks LibreOffice gave are dropped on the host and accepted: the Excel-vs-Python cross-check, and the PNG render that finds text running off a slide. All numbers already come from Python when LibreOffice is missing |
| Local use | Unchanged. `python run.py` and `python serve.py` on a PC with LibreOffice still make the PDF and run both checks |
| Sign-in | **Not now** (decided 23/11/2026). Anyone with the site's link can use it, so the link is shared only inside CREST. "Sign in with Google" limited to `@crest-group.co` can be added later |
| Storage | Nothing that matters is kept on the host's disk (Render wipes it on every restart and after 15 idle minutes). Result files are downloaded straight after the run |
| AI caches | **Not moved** (decided 23/11/2026). On the host, `ai_research.json` and `ai_ratings.json` last until the server restarts or sleeps. Answers written into the Universal Sheet (Data Source, Rating Source) are still reused, so little is asked twice. Added 23/11/2026 so nothing depends on those files: an **AI Searched** column remembers "results not out yet, searched on date X", and sheet writes retry brief Google errors and follow rows that moved |
| Client data | Uploaded holdings pass through Render while a run works and are wiped when the server sleeps. **CREST to agree this before go-live** |

**What Render's free plan means in practice**

- The site sleeps after 15 minutes with no visitors; the first visit after that shows a loading page for about a
  minute.
- Render may restart the server at any time. A run in progress when that happens fails and is run again.
- Run history covers only the runs since the server last woke. Staff download the Excel and PPT when the run
  finishes; the page says so.
- 0.1 CPU is slow. A full run is timed during the trial (step 7); if it is too slow, Hugging Face Spaces (also
  free, no card, more CPU and memory) is the fallback, with the same code.

**Step 0: confirm before building**

Render's docs (read 23/11/2026) do not say whether sign-up needs a card, whether the free plan allows Docker, or
the free bandwidth and build-minute figures. Sign up for Render with a CREST account first and confirm: no card
asked, a free web service can be created, and the Singapore region is offered. If a card is asked for, stop and
use Hugging Face Spaces instead.

**Step 1: hosted mode without LibreOffice**

- A setting (environment variable `CREST_HOSTED=1`) for the hosted site.
- On the host, "LibreOffice not found" is not a warning: the report and website say "PDF: make it in
  PowerPoint" instead.
- The website shows "PowerPoint" instead of "PowerPoint and PDF", and leaves out the "Checking the Excel
  formulas" stage when hosted. `pymupdf` is already loaded only when a PDF is rendered.

**Step 2: settings from environment variables**

- The host has no `.env`, so `load_env()` also reads every setting it knows (sheet, keys, hosted mode) straight
  from the environment; the host sets them in Render's dashboard.
- New: the Google service-account key can be given as the **contents** of the key file
  (`GOOGLE_CREDENTIALS_JSON`), not only a path, so no key file is needed on the host.
- The deck template (`Deck_-_template.pptx`, not in git) is uploaded to the host as a private file, and its
  path is given in `CREST_DECK_TEMPLATE`.
- Lost on the host, and accepted: a saved `narrative.json` for a client (a re-run asks the narrative API again,
  one request) and the "use the latest snapshot if Google is down" fallback (the run stops with a clear message
  instead).

Steps 3 (AI caches in the Universal Sheet) and 4 (Google sign-in) were dropped on 23/11/2026; see Decisions.

**Step 5: deployment files**

- `render.yaml` (or Dockerfile, per step 0): install Python requirements, build the front end
  (`npm ci`, `npm run build`), start with `uvicorn` on Render's port.
- `serve.py` reads the port and host from the environment.

**Step 6: tests**

- Hosted mode: no LibreOffice warning, no PDF in the results, the stage list without the two LibreOffice stages.
- `GOOGLE_CREDENTIALS_JSON` read in place of a file; settings read from the environment with no `.env`.
- All existing tests still pass.

**Step 7: trial on Render**

- Deploy and run a sample client: Portfolio, Deck and Both.
- Time a full run and watch memory in Render's dashboard.
- Confirm the PPT opens in PowerPoint and saves as PDF with the right fonts.

**Before go-live (CREST)**

- Agree that client holdings may pass through Render.
- Fix the narrative-API key (the free-tier 403), or the hosted site runs without AI research and narrative.

## 7. What is kept and what is removed

**Kept as is:** matching (ISIN, then symbol, then alternate symbols, then name), the Excel build, deck filling,
deck data, narrative checks and house style, the narrative API client (model switching, tolerant reply reading),
ratings and the rating cache, deck quality checks, LibreOffice handling, validation checks, the run report, the
sector-to-deck-group map and the deck settings.

**Removed:** the broker auto-detect code and its configuration, the XML write-back to the local Excel, the
"broker price" option, client name from the file name, as-on date from the file's saved date, and the local
Universal Sheet as the source of truth (it stays only as an offline fallback and for tests).

## 8. How it will be tested

- **Automatic tests** (never call Google or the AI API):
  - Template reader: a good file, a missing column, a blank client name or date, bad numbers, merged duplicates.
  - Step B reading a Portfolio Excel: as built, after edits (row deleted, column moved, Key blank), and after a
    Google Drive round trip.
  - Universal Sheet parsing from in-memory data, the "Loading..." retry, and falling back to a snapshot.
  - AI research choices: typed cells untouched, an older quarter triggers a request, a saved answer means no
    request.
  - The existing 41 tests moved to the template format.
- **End to end, locally:** one past client converted to the template and run against a saved snapshot of the
  current sheet. Excel numbers and deck tables must match the earlier Q2 FY27 run for the same prices.
- **Google:** read the shared sheet (about 5,334 rows, with the count of `#N/A` prices reported); test the
  write-back on a copy of the sheet first, then run a full client against the live sheet.
- **Website:** API tests for upload, queueing, progress and downloads (no network); then in the browser: upload a
  template and download the Excel; upload that Excel and download the PPT and PDF; run "Both" and confirm the
  Excel and the deck show the same numbers; check the screens at desktop and laptop widths.
- Code lint check clean.

## 9. What CREST provides, and when

| When | What |
|---|---|
| Phase 3 | The service-account key saved as a file (never pasted in a chat); the Universal Sheet uploaded to Google Sheets and shared with the service account's email; the sheet ID |
| Phase 4 | The paid narrative-API key, added to `.env` |
| Phase 6 | A Render account (no card); agreement that client holdings may pass through Render |
| End of every phase | Run the commit and push commands provided |
