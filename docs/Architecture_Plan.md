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
| AI | Paid narrative-API key with web search. Small batches, saved and reused, with a limit per run. This replaces the old "at most one request per run" rule |
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
| Website | Three tabs: Portfolio, Deck, Both |
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
- Other sheets: Raw_Holdings, Holdings_Normalized, Universal_Snapshot, Exceptions.

Two additions so Step B can read it back reliably:

- A hidden marker cell ("CREST Portfolio v1") so Step B can tell a Portfolio Excel from a holdings template.
- Step B finds the Portfolio columns **by header name**, not by position. Staff edits still read correctly:
  deleted or added rows, re-ordered columns, or a copy downloaded again from Google Drive.

### 5.3 Universal Sheet on Google Sheets

The same columns as today, plus:

- **Data Source** and **Data Updated**, written by the tool when AI fills Sector, Result Line or Commentary. They
  record the model, period, date and source links, in the same way Rating Source and Rating Reason do today.
- Cells CREST types have a blank Data Source and are never overwritten.

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

- A Streamlit app with three tabs:
  - **Portfolio:** download the blank template; upload a filled template; download the Portfolio Excel.
  - **Deck:** upload a Portfolio Excel or a template; download the PPT and PDF.
  - **Both:** upload a template; download the Excel, PPT and PDF.
- Every run shows its status (Passed, Passed with warnings, Failed), the checks, the exceptions and the
  AI-sourced fields to review.
- Each upload runs in its own folder under `runs\`. No client data leaves the server except the trimmed,
  anonymous narrative input already sent today.
- Started with `streamlit run app.py`.

### Phase 6: Later (not in this build)

- Google sign-in limited to `@crest-group.co` accounts.
- A container image with LibreOffice, deployed to Google Cloud Run.
- The run archive kept in Google Cloud Storage instead of on the server's disk.

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
- **Website:** upload a template and download the Excel; upload that Excel and download the PPT and PDF; run
  "Both" and confirm the Excel and the deck show the same numbers.
- Code lint check clean.

## 9. What CREST provides, and when

| When | What |
|---|---|
| Phase 3 | The service-account key saved as a file (never pasted in a chat); the Universal Sheet uploaded to Google Sheets and shared with the service account's email; the sheet ID |
| Phase 4 | The paid narrative-API key, added to `.env` |
| End of every phase | Run the commit and push commands provided |
