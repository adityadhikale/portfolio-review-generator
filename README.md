# CREST Portfolio Review Generator

Turns a client's holdings, entered in the standard CREST holdings template, into two deliverables:

* **`<Client> Portfolio.xlsx`**: holdings, values, P&L and weights. Every calculated cell is a formula, and Market Cap and
  CMP are Google Finance formulas.
* **`<Client>_CREST_Portfolio_<Period>_Review.pptx`** (plus a PDF): the CREST quarterly review deck, built from the CREST
  template.

Numbers and structure come from plain, deterministic Python. The written commentary slides are filled from a
`narrative.json`, which you supply or which the optional narrative API (set up in `.env`) writes (see
[Narrative](#narrative-slides)). The Universal Sheet is read from Google Sheets (see
[Google Sheets setup](#google-sheets-setup)), or from a local file until that is set up.

---

## Quick start

Once, in the project folder: create the virtual environment and install the packages into it.

```powershell
python -m venv .venv
```

```powershell
.\.venv\Scripts\Activate.ps1
```

```powershell
python -m pip install -r requirements.txt
```

Every new terminal: activate it first (`.\.venv\Scripts\Activate.ps1`; the prompt shows `(.venv)`). If Windows
refuses ("running scripts is disabled"), run `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once, or skip
activating and use `.\.venv\Scripts\python.exe` in place of `python`.

**Website** (see [Website](#website)): `python serve.py`, then open http://localhost:8000.

**Command line:**

1. Copy the client's holdings into a copy of **`templates\Holdings_Template.xlsx`** (see
   [Holdings template](#holdings-template)) and save it in **`input\`**.
2. Run:

   ```powershell
   python run.py
   ```

3. Collect the results from **`output\<Client>\<Period>\`**:

| File | What it is |
|---|---|
| `<Client> Portfolio.xlsx` | Portfolio workbook |
| `<Client>_CREST_Portfolio_<Period>_Review.pptx` / `.pdf` | Review deck |
| `report.md` | Checks passed or failed, exceptions, warnings, what still needs input |
| `rating_preview.md` | Ratings the draft rules would give (see [Rating rules](#rating-and-sentiment-rules)) |

Every file in `input\` is processed in one go, and the tool works the rest out:

| Setting | Where it comes from |
|---|---|
| Period (e.g. Q2FY27) | the Universal Sheet's `Q Result Period` column |
| Client name | cell C2 of the holdings template |
| As-on date | cell C3 of the holdings template |
| Universal Sheet | the Google Sheet named in `.env` (`UNIVERSAL_SHEET_ID`); without it, `templates\universal-sheet.xlsx` |
| Deck template | `templates\Deck_-_template.pptx` |

Optional overrides: `--step`, `--holdings` (one file), `--period`, `--universal`, `--template`, `--narrative`,
`--output-dir`, `--runs-dir`, `--settings`, `--no-render`.

### Step A and Step B

A run has two steps, which can also be run on their own, so a deck can be made without making a new Excel:

| Command | Input | Output |
|---|---|---|
| `python run.py` (or `--step both`) | holdings template | Portfolio Excel, then PPT + PDF |
| `python run.py --step portfolio` | holdings template | Portfolio Excel (Step A) |
| `python run.py --step deck` | Portfolio Excel, or a holdings template | PPT + PDF (Step B) |

* **Step B always takes fresh data from the Universal Sheet**: price, market cap, sector, views, results,
  commentary, rating and sentiment. Only the holdings come from the uploaded file.
* **A Portfolio Excel edited by hand still works.** Quantity and Average Price in its Portfolio sheet win. Rows can
  be deleted or added, columns moved, and the file can be a copy downloaded again from Google Drive (columns are
  found by header name). A row with no Key (added by hand) is matched by its Ticker, then by its Name. Holdings that
  Step A had but the Portfolio sheet no longer has are listed in the report.
* A Portfolio Excel is recognised by a hidden marker in cell A1 ("CREST Portfolio v1"). It can only be used for
  Step B; Step A always needs a holdings template.
* A `both` run reads the Universal Sheet once, so the Excel and the deck show identical numbers.
* A single-step run replaces only its own results in `output\<client>\<period>\`: a deck run leaves the
  Portfolio Excel there, and a Portfolio run leaves the deck.

---

## Website

Staff can do everything from a browser: choose **Portfolio Excel**, **Review deck** or **Excel and deck**, drop in the
file, watch the run step by step, then download the Excel, PPT, PDF and report and review the checks, exceptions
and AI-written values. Every run stays in **History** with its files.

Build the front end once (and again after any change in `web\`), then start the server:

```powershell
cd web; npm install; npm run build; cd ..
```

```powershell
python serve.py
```

Open **http://localhost:8000**. `python serve.py --host 0.0.0.0` makes it reachable from other machines on the
office network (there is no login yet; that is Phase 6).

| Part | What it is |
|---|---|
| `crest/api.py` | FastAPI: upload, run queue (one run at a time), live progress (server-sent events), results, downloads, history; API docs at `/api/docs` |
| `serve.py` | Starts the API and serves the built front end |
| `web/` | Vue 3 + Vite + TypeScript, shadcn-vue components, Motion for Vue, Tailwind; CREST colours and Maharlika headings from the deck |
| `runs\_web\<run>\` | Each website run: the upload, `job.json`, and a copy of its result files (history downloads never change) |

**Front-end development:** with `python serve.py` running, `cd web` then `npm run dev` serves the front end at
http://localhost:5173 with live reload, forwarding `/api` to the Python server. Node.js 20+ is needed only to build
the front end.

## Requirements

* Python 3.11+ and `pip install -r requirements.txt` (openpyxl, python-pptx, pymupdf, google-auth, requests).
* **LibreOffice** (for formula checks and the PDF / slide render): `winget install TheDocumentFoundation.LibreOffice`.
  Without it the run still completes; report.md says which checks were skipped. A non-standard install path can be
  set in `config/settings.json` → `libreoffice_path`.

---

## Private files (not in git)

Client data and CREST research never go to GitHub; `.gitignore` excludes them. After cloning, put these in place yourself:

| File | Purpose |
|---|---|
| `secrets/google-service-account.json` | The Google service-account key (see [Google Sheets setup](#google-sheets-setup)) |
| `.env` | Google Sheet ID, narrative API key and models, Tavily search key (copy `.env.example`) |
| `templates/universal-sheet.xlsx` | Local copy of the Universal Sheet: offline fallback until Google Sheets is set up, and for tests. Never written by the tool |
| `templates/Deck_-_template.pptx` | The CREST deck template |
| `input/*` | Filled holdings templates (client data) |
| `samples/holdings.xlsx`, `samples/reference_portfolio.xlsx`, `samples/reference_deck_Q1FY27.pptx` | Reference files, only needed for the end-to-end tests (the tests copy the broker sample into the holdings template) |

`output/`, `runs/` and the generated test fixtures are also excluded.

---

## Folder layout

```
run.py                     command line (python run.py)
serve.py                   website (python serve.py)
crest/                     the pipeline
  holdings.py              reads and checks the holdings template (duplicate merge, exact-cell errors)
  portfolio_reader.py      reads holdings back from a Portfolio Excel for Step B (columns found by header name)
  universal.py             Universal Sheet loader and schema checks
  match.py                 holdings -> Universal rows (ISIN, symbol, alt symbols, exact name)
  excel_build.py           the Portfolio workbook
  office.py                LibreOffice: recalculation and PDF conversion
  validate.py              validation gate
  ratings.py               draft Rating / Sentiment rules
  narrative_api.py         narrative API client (model list, switching on busy / bad replies)
  ai_review.py             the single API request per client: ratings + narrative
  ai_research.py           AI fills blank Sector / Result Line / Commentary (web search, checks, write-back)
  ai_ratings.py            per-quarter backup of API ratings (runs/<period>/ai_ratings.json)
  sheet_source.py          Universal Sheet source: Google Sheets API (read, write-back) or local file; snapshots
  deck_data.py             numbers for the deck, read back from the Excel
  narrative.py             narrative template, saved narrative.json, number and style checks
  deck_fill.py             fills the PowerPoint template
  qa.py                    deck QA (placeholders, style, overflow, render)
  api.py                   website API (FastAPI): uploads, run queue, live progress, downloads
  pipeline.py              runs shared setup, then Step A and/or Step B; archives and publishes each run
  shared.py                shared setup: Universal Sheet, matching, ratings, deck groups, data checks, exceptions
  step_portfolio.py        Step A: Portfolio Excel, LibreOffice cross-check, validation gate
  step_deck.py             Step B: deck data, narrative and API request, deck, QA, PDF
  report.py                writes report.md
config/                    settings, sector groups, sanity rules, rating rules
templates/                 Holdings_Template.xlsx and Universal_Sheet_template.xlsx (blank, in git);
                           live sheet and deck template (private)
input/  output/  runs/     client files in, results out, full run archive
web/                       website front end (Vue); web/dist is the build
tests/                     unit + end-to-end tests
tools/                     developer helpers (holdings template, test fixtures, comparison with reference files)
```

---

## How a run works

1. **Read the holdings** (`crest/holdings.py`, or `crest/portfolio_reader.py` for a Portfolio Excel in Step B):
   client name, as-on date and one row per holding. Rows for the same stock are merged at a weighted average price.
   Any problem stops the run and names the cell to fix.
2. **Match** each holding to the Universal Sheet: by ISIN, then symbol, then Alt Symbols (BSE codes, former symbols),
   then exact company name. Anything else goes to Exceptions; the tool never guesses.
3. **Tickers** are written as `NSE:SYMBOL` / `BSE:CODE`. `BOM:` is rewritten as `BSE:`, and a ticker with no exchange
   defaults to NSE and is flagged.
4. **Build the Excel** (Step A): sheets `Portfolio`, `Raw_Holdings` (the template as uploaded, unchanged), `Holdings_Normalized`,
   `Universal_Snapshot` (this client's stocks only) and `Exceptions`. The Portfolio columns are:

   | Column | Content |
   |---|---|
   | Name, Sector, Quantity Available, Average Price | values |
   | Ticker, Views | `IFERROR(INDEX/MATCH)` to Universal_Snapshot |
   | Market Cap (Cr) | `=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE(C9,""marketcap"")/10^7"),<value>)` |
   | CMP | `=IFERROR(__xludf.DUMMYFUNCTION("GOOGLEFINANCE(C9,""price"")"),<value>)` |
   | Invested / Current Value / Unrealized P&L | `=F*G`, `=F*I`, `=J-H` |
   | Unrealized P&L (%), % Allocation | `=K/H`, `=J/SUM(J)` |

   Amounts have no currency symbol. Uploaded to Google Drive, the Google Finance formulas go live; Excel shows the run's
   value. BSE tickers are passed to Google as `BOM:`. LibreOffice cannot evaluate those formulas, so it recalculates a
   check copy with the run's values in their place. That copy is scanned for `#N/A`, `#REF!`, `#DIV/0!` and other errors,
   and checked against an independent Python calculation.
5. **Validation gate**:
   * every holding is in the Portfolio table or in Exceptions, and the invested amounts add up;
   * unmatched symbols and missing prices (listed in Exceptions);
   * % Allocation = 100%;
   * sector sanity and missing sectors;
   * Universal `Q Result Period` = run period;
   * ETF market cap and ticker format.

   Any FAIL stops the run with a readable report. A deck-only run checks the same numbers, computed in Python.
6. **Deck** (Step B, `crest/deck_fill.py`):
   * cover, the Positive / Average / Weak counts, and results tables grouped by sector (`config/sector_groups.json`),
     sorted by weight;
   * groups over 8 rows split into (1/2), (2/2) slides; unused template slides are removed;
   * the template's footer is rewritten for the client and period.
7. **QA**: leftover placeholders fail the run. It also checks that the brand is written CREST and that no heading
   starts with a number, estimates text overflow, and renders PDF + PNG.
8. **Archive**: `runs/<period>/<client>/` keeps inputs, outputs, `narrative_input.json`, slide images and `report.md`
   (with a checks section for each step).
   Earlier runs move to `_history/`. Results are copied to `output\<client>\<period>\`; a file there that is open in
   Excel / PowerPoint is not replaced - the new version is saved beside it as `<name> (new).<ext>`.

**Pricing**: every price is the Universal Sheet CMP at the time of the run. Neither the holdings template nor the
Portfolio Excel supplies prices.

---

## Universal Sheet

Blank layout with a column guide: `templates/Universal_Sheet_template.xlsx` (Legend sheet). Header row 1, one row per stock.

| Column | Status | Notes |
|---|---|---|
| Name, Ticker, Sector, Market Cap (Cr), CMP, Views | required columns | blank cells allowed (flagged) |
| Result Line (or `Q1 Result Summary` / `Result Summary`), Rating, Sentiment, Management Commentary | required columns | quarterly fields for the results tables |
| Symbol, Alt Symbols, ISIN | optional | used for matching holdings |
| Q Result Period | optional | e.g. `Q2 FY27`; sets the run period, and a mismatch stops the run |
| Short Name, Deck Group, Exchange, Instrument Type | optional | display name, group override, ETF flag (`Sector = ETF` also works) |
| Data Source, Data Updated | written by the tool | added when the AI fills a blank Sector / Result Line / Commentary: model, quarter, which fields, source links; leave blank for values you type. The sheet's own Source / Last Updated columns are never written |
| Rating Source, Rating Reason | written by the tool | added when the narrative API fills a blank Rating / Sentiment (source, period and a change code); leave blank for ratings you type |

**Where it is read from.** The Universal Sheet lives in Google Sheets, where GOOGLEFINANCE keeps price and market cap
live. Each batch of runs reads the whole tab in one request, so every client in the batch, and both steps of a run,
use the same data. Cells still showing `Loading...` are read once more after 15 seconds; a price still `#N/A`,
`Loading...` or blank sends that stock to Exceptions as "Missing price".

**Snapshots.** Every run saves the data it used in `runs\<period>\snapshots\` (compressed, about 250 KB; a new file
only when the data changed) and `report.md` names it. If Google cannot be reached, the latest snapshot is used
instead and the run shows a WARN check, "Universal Sheet read live from Google". If Google refuses access (sheet not
shared, wrong ID or tab name), the run stops with the reason.

**Writing back.** Ratings from the narrative API are written to the Google Sheet in one request: Rating, Sentiment,
Rating Source and Rating Reason, only for stocks whose Rating and Sentiment were blank. Just before writing, the
tool checks each target row still holds the same company (in case rows were inserted or sorted meanwhile) and skips
any row that moved. Nothing is written when the run used a snapshot or the local file; the ratings stay in
`runs\<period>\ai_ratings.json` and are written on the next run that reads the Google Sheet.

**Local file.** With no `UNIVERSAL_SHEET_ID` in `.env`, or with `--universal <file.xlsx>`, the tool reads
`templates\universal-sheet.xlsx` (values saved in the file) and never writes to it.

## Google Sheets setup

Done once by CREST. The key file is a password: never paste it into a chat or an email, and never commit it.

1. Go to [console.cloud.google.com](https://console.cloud.google.com), sign in with the CREST Google account, and
   create a project (e.g. "CREST Portfolio Review").
2. In that project, open **APIs & Services → Library**, search for **Google Sheets API**, and click **Enable**.
3. Open **IAM & Admin → Service Accounts → Create service account**. Name it (e.g. `crest-portfolio`), skip the
   optional role and user steps, and click **Done**. Copy its email address (it ends in `.iam.gserviceaccount.com`).
4. Click the service account, open **Keys → Add key → Create new key → JSON**. A `.json` file downloads. Save it in
   this project as `secrets\google-service-account.json` (create the `secrets` folder; git ignores it).
5. Upload the Universal Sheet to Google Drive and open it with Google Sheets (**File → Save as Google Sheets** if it
   opened as an .xlsx). Check that the stocks are on a tab named `Universal` and that GOOGLEFINANCE prices show.
6. Click **Share**, paste the service account's email, choose **Editor**, untick "Notify people", and share.
7. Copy the sheet ID from its address, `docs.google.com/spreadsheets/d/<SHEET ID>/edit`, and put it in `.env`:

   ```
   UNIVERSAL_SHEET_ID=<SHEET ID>
   UNIVERSAL_WORKSHEET=Universal
   GOOGLE_CREDENTIALS_FILE=secrets/google-service-account.json
   ```

8. Check it: `python run.py --step portfolio` prints `Universal Sheet: Google Sheet ...`, and `report.md` shows
   "Universal Sheet read live from Google: PASS".

Tip: test the write-back on a **copy** of the sheet first (share the copy the same way and put the copy's ID in
`.env`), then switch the ID to the real sheet.

## AI research: Sector, Result Line, Commentary

For the stocks held in the portfolio being run, the AI fills, before ratings:

* **Sector** when it is blank, choosing only a sector the deck knows (`config/sector_groups.json`);
* **Result Line** and **Management Commentary** when they are blank, or were written by the AI earlier for an
  older quarter.

**Keeping requests down.** A Portfolio-only run fills blank Sectors only, with no web search (the Excel shows no
result line or commentary); results and commentary are researched only when a deck is made. Stocks that only need
a Sector ride along in the first research request. When a researched stock has no Rating / Sentiment, the same
request rates it, so the later ratings + narrative request covers only what is left. Everything found is written
to the Universal Sheet and saved per quarter, so the next client holding the same stock costs nothing.

A value CREST typed is never changed. A row whose typed result is for another quarter is left alone (the period
check flags it). Views are never touched.

**How it searches** (`ai_research` in `config/settings.json`):

| Setting | What happens | Needs |
|---|---|---|
| `auto` (default) | `tavily` when `.env` has `TAVILY_API_KEY`, else `sector_only` | |
| `tavily` | a Tavily web search per stock, then one narrative-API request per batch reads the search extracts | narrative API key (free tier is enough) + free Tavily key (1,000 searches a month, app.tavily.com) |
| `api_search` | the narrative API searches the web itself | a narrative API key with billing turned on |
| `sector_only` | blank Sectors only, no web search; results stay "Results awaited" | narrative API key |
| `off` | nothing | |

Only company names and tickers are sent, never client data.

**Checks before a value is used.** The answer must be for the run's quarter (otherwise "results awaited"), and
must give source links. The Result Line is built by the tool in the house format from the returned figures, e.g.
`Rev +45.9% YoY (Rs 610cr); PAT +153.6% YoY (Rs 71cr)`, and must be readable by the rating rules. With Tavily,
the revenue or profit figure must appear in the cited search extracts, and every figure in the commentary too;
otherwise that value is rejected and listed in the report.

**Limits.** Batches of `ai_research_batch_size` stocks (20) per request, largest holdings first, and at most
`ai_max_requests_per_run` narrative-API requests per client run (3), with one always kept for the ratings + narrative
request. Stocks over the limit are researched on the next run.

**Saved answers.** `runs\<period>\ai_research.json` keeps every answer: a stock's result for a quarter is never
asked for twice, and "not published yet" is asked again after `ai_research_retry_days` (3). New values are
written to the Google Sheet with **Data Source** (`AI · model · quarter · fields · source links`) and **Data
Updated**.

**Review.** `report.md` lists every AI-written value among the client's holdings under **"AI-sourced: review
before sending"**, with its sources, and what could not be filled and why. To correct a value, type over it in
the Universal Sheet and clear that row's Data Source.

## Rating and Sentiment rules

A value typed in the Universal Sheet always wins. With the narrative API set up (below), a blank Rating / Sentiment is
filled by it and saved into the sheet. Without it, a blank Rating or Sentiment is filled from
`config/rating_rules.json` **only after `"approved": true`** is set there. The rules read revenue and PAT YoY from the
result line. Every run writes `rating_preview.md` with the rating each stock would get, so the rules can be reviewed first.

## Narrative slides

Slides 2-5 and 15 hold written analysis. Each run writes `runs/<period>/<client>/outputs/narrative_input.json`: every
computed fact, plus the list of fields with length limits. To fill them:

1. Draft the text from that file, in the shape of `narrative_template.json`.
2. Save it as `runs/<period>/<client>/narrative.json` (or pass `--narrative`).
3. Run again.

Until then those slides show **PENDING NARRATIVE**. Every number in narrative.json must exist in
narrative_input.json, otherwise the run fails and lists the figure.

**Narrative API: narrative, Rating and Sentiment.** Copy `.env.example` to `.env` (ignored by git) and fill in
`NARRATIVE_PROVIDER`, `NARRATIVE_API_KEY` and, optionally, `NARRATIVE_MODEL`: one model or a comma-separated list in
order of preference (blank = the built-in list in `crest/narrative_api.py`). After AI research (which already rates
the stocks it researches), each deck run makes **one request** that rates the stocks still without a Rating /
Sentiment and writes the narrative (client name and amounts removed). Views always come from the Universal Sheet only. The API's ratings are written into the
Universal Sheet (Rating, Sentiment, and
the Rating Source / Rating Reason columns) and kept as a backup in `runs/<period>/ai_ratings.json`; every later run,
for any client, reads them from there and sends no request for those stocks. A stock is rated again only when its
result line or commentary changes. If the run could not write to the Google Sheet, the next one that can does. The narrative is
saved in `runs/<period>/<client>/narrative.json` and reused. Edit the sheet or that file to correct anything;
delete the narrative (or a stock's Rating Source) to have it written again. There are no automatic
correction requests: if the narrative fails a check, the run names the field to edit. If a model is busy, gives
no answer within 90 seconds, is over quota or not available for the key, the same request goes straight to the
next model in the list; if all fail, the list is tried once more after 30 seconds. If no model answers, the Excel
and deck are still built (narrative slides pending) and the next run asks again. Set `narrative_provider` and
`rating_provider` to `"manual"` / `"rules"` in `config/settings.json` to never call the API.

## Troubleshooting

| What you see | What to do |
|---|---|
| `No module named ...`, or "A package needed to read the Google Sheet is missing" | Activate `.venv`, then `python -m pip install -r requirements.txt` |
| The virtual environment does not start | Create it (Quick start); if scripts are refused, `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` once |
| Narrative slides say **PENDING NARRATIVE** and the report shows an API error 403 ("free tier is not available for this project") | The key's project needs billing, or use a key from a project where the free tier works; put it in `.env` and run again |
| A run is slow at "AI research" or "Ratings and narrative" | The narrative API is busy; the tool switches models by itself. Nothing to do |
| "Google refused access" | Share the Universal Sheet with the service account's email as Editor (see Google Sheets setup) |

## Holdings template

`templates\Holdings_Template.xlsx` is the only holdings format the tool reads (re-create it with
`python tools/build_holdings_template.py`). Its second sheet, "How to fill", explains it for staff.

| Cell / row | Content |
|---|---|
| `B2` / `C2` | "Client Name" / the client's name, as it should appear on the deck |
| `B3` / `C3` | "As On Date" / the statement date, `dd/mm/yyyy` or a real Excel date |
| Row 5 | Headers: `ISIN`, `Symbol`, `Company Name`, `Quantity`, `Average Price` (found by name, any order) |
| Row 6 onwards | One holding per row, ending at the first blank row |

Every row needs Quantity above 0, Average Price above 0, and an ISIN or a Symbol (NSE symbol or BSE code). Company
Name is optional and is only used for matching when neither ISIN nor Symbol is found. Rows for the same stock (same
ISIN or symbol) are merged at a weighted average price. A missing column, text where a number should be, a malformed
ISIN, or a blank client name or date stops the run and lists every cell to fix. Rows below the first blank row are
not read, and the run warns about them.

## Configuration

| File | Purpose |
|---|---|
| `config/settings.json` | AI research and request limits, narrative checks, tolerances, table split sizes |
| `config/sector_groups.json` | sector (Screener industry) → deck group, group order, per-stock overrides |
| `config/sector_rules.json` | name keywords → plausible sectors (sanity check) |
| `config/rating_rules.json` | draft rating rules (inactive until approved) |
| `config/name_overrides.json` | display names by symbol |

## Tests

```powershell
python -m unittest discover -s tests -v
```

Unit tests always run. End-to-end tests need the private sample files (see [Private files](#private-files-not-in-git)).
Rebuild their fixtures with these commands:

```powershell
python tools/build_fixture.py
```

```powershell
python tools/extract_narrative.py samples/reference_deck_Q1FY27.pptx tests/fixtures/narrative_reference.json
```

`python tools/compare_reference.py runs/<period>/<client>` compares a run with the reference Excel and deck.
