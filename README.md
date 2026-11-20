# CREST Portfolio Review Generator

Turns a client's broker holdings file into two deliverables:

* **`<Client> Portfolio.xlsx`**: holdings, values, P&L and weights. Every calculated cell is a formula, and Market Cap and
  CMP are Google Finance formulas.
* **`<Client>_CREST_Portfolio_<Period>_Review.pptx`** (plus a PDF): the CREST quarterly review deck, built from the CREST
  template.

Numbers and structure come from plain, deterministic Python. No AI is called and nothing is sent over the network.
The written commentary slides are filled from a `narrative.json` that you supply (see [Narrative](#narrative-slides)).

---

## Quick start

```powershell
pip install -r requirements.txt
```

1. Put the client's broker file (.xlsx or .csv, exactly as downloaded) in **`input\`**.
   The file name becomes the client name: `input\Sharma Family.xlsx` → client "Sharma Family".
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
| As-on date | the holdings file's title ("as on dd/mm/yyyy"), else the file's saved date (flagged) |
| Universal Sheet | `templates\universal-sheet.xlsx` |
| Deck template | `templates\Deck_-_template.pptx` |

Optional overrides: `--client`, `--holdings`, `--period`, `--asof`, `--universal`, `--template`, `--narrative`,
`--price-source broker_ltp`, `--broker`, `--output-dir`, `--runs-dir`, `--settings`, `--no-render`.

---

## Requirements

* Python 3.11+ and `pip install -r requirements.txt` (openpyxl, python-pptx, pymupdf).
* **LibreOffice** (for formula checks and the PDF / slide render): `winget install TheDocumentFoundation.LibreOffice`.
  Without it the run still completes; report.md says which checks were skipped. A non-standard install path can be
  set in `config/settings.json` → `libreoffice_path`.

---

## Private files (not in git)

Client data and CREST research never go to GitHub; `.gitignore` excludes them. After cloning, put these in place yourself:

| File | Purpose |
|---|---|
| `templates/universal-sheet.xlsx` | The live Universal Sheet (all stocks, results, views); backups the tool makes before writing to it go to `templates/backups/` |
| `.env` | Gemini API key and models (copy `.env.example`) |
| `templates/Deck_-_template.pptx` | The CREST deck template |
| `input/*` | Client holdings files |
| `samples/holdings.xlsx`, `samples/reference_portfolio.xlsx`, `samples/reference_deck_Q1FY27.pptx` | Reference files, only needed for the end-to-end tests |

`output/`, `runs/` and the generated test fixtures are also excluded.

---

## Folder layout

```
run.py                     entry point (python run.py)
crest/                     the pipeline
  brokers.py               reads any broker layout (header detection, duplicate merge)
  universal.py             Universal Sheet loader and schema checks
  match.py                 holdings -> Universal rows (ISIN, symbol, alt symbols, exact name)
  excel_build.py           the Portfolio workbook
  office.py                LibreOffice: recalculation and PDF conversion
  validate.py              validation gate
  ratings.py               draft Rating / Sentiment rules
  gemini.py                Gemini client (model list, switching on busy / bad replies)
  ai_review.py             the single Gemini request per client: ratings + narrative
  ai_ratings.py            per-quarter backup of Gemini ratings (runs/<period>/ai_ratings.json)
  sheet_writeback.py       writes Gemini ratings into the Universal Sheet without touching other cells
  deck_data.py             numbers for the deck, read back from the Excel
  narrative.py             narrative template, saved narrative.json, number and style checks
  deck_fill.py             fills the PowerPoint template
  qa.py                    deck QA (placeholders, style, overflow, render)
  pipeline.py              runs the steps in order and archives each run
  report.py                writes report.md
config/                    settings, broker column names, sector groups, sanity rules, rating rules
templates/                 Universal_Sheet_template.xlsx (blank, in git); live sheet and deck template (private)
input/  output/  runs/     client files in, results out, full run archive
tests/                     unit + end-to-end tests
tools/                     developer helpers (test fixtures, comparison with reference files)
```

---

## How a run works

1. **Read the holdings** (`crest/brokers.py`): finds the header row in any layout. It needs a quantity column, an
   average/cost price column, and a symbol, ISIN or company-name column. Title, blank and total rows are skipped,
   duplicate symbols are merged at a weighted average price, and zero-quantity rows are dropped.
2. **Match** each holding to the Universal Sheet: by ISIN, then symbol, then Alt Symbols (BSE codes, former symbols),
   then exact company name. Anything else goes to Exceptions; the tool never guesses.
3. **Tickers** are written as `NSE:SYMBOL` / `BSE:CODE`. `BOM:` is rewritten as `BSE:`, and a ticker with no exchange
   defaults to NSE and is flagged.
4. **Build the Excel**: sheets `Portfolio`, `Raw_Holdings` (the upload, unchanged), `Holdings_Normalized`,
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
   * broker tie-out (within Σ qty × ₹0.005 for 2-dp rounding of average prices);
   * unmatched symbols and missing prices (listed in Exceptions);
   * % Allocation = 100%;
   * sector sanity and missing sectors;
   * Universal `Q Result Period` = run period;
   * ETF market cap and ticker format.

   Any FAIL stops the run with a readable report.
6. **Deck** (`crest/deck_fill.py`):
   * cover, the Positive / Average / Weak counts, and results tables grouped by sector (`config/sector_groups.json`),
     sorted by weight;
   * groups over 8 rows split into (1/2), (2/2) slides; unused template slides are removed;
   * the template's footer is rewritten for the client and period.
7. **QA**: leftover placeholders fail the run. It also checks that the brand is written CREST and that no heading
   starts with a number, estimates text overflow, and renders PDF + PNG.
8. **Archive**: `runs/<period>/<client>/` keeps inputs, outputs, `narrative_input.json`, slide images and `report.md`.
   Earlier runs move to `_history/`. Results are copied to `output\<client>\<period>\`; a file there that is open in
   Excel / PowerPoint is not replaced - the new version is saved beside it as `<name> (new).<ext>`.

**Pricing**: `price_source` in `config/settings.json` (or `--price-source`) is `universal_cmp` (default, the
Universal Sheet CMP) or `broker_ltp`. One run never mixes the two.

---

## Universal Sheet

Blank layout with a column guide: `templates/Universal_Sheet_template.xlsx` (Legend sheet). Header row 1, one row per stock.

| Column | Status | Notes |
|---|---|---|
| Name, Ticker, Sector, Market Cap (Cr), CMP, Views | required columns | blank cells allowed (flagged) |
| Result Line (or `Q1 Result Summary` / `Result Summary`), Rating, Sentiment, Management Commentary | required columns | quarterly fields for the results tables |
| Symbol, Alt Symbols, ISIN | optional | used for matching broker files |
| Q Result Period | optional | e.g. `Q2 FY27`; sets the run period, and a mismatch stops the run |
| Short Name, Deck Group, Exchange, Instrument Type | optional | display name, group override, ETF flag (`Sector = ETF` also works) |
| Rating Source, Rating Reason | written by the tool | added when Gemini fills a blank Rating / Sentiment (`AI · model · period · code`); leave blank for ratings you type |

Formula cells (e.g. GOOGLEFINANCE) are read using the value saved in the file. A `#N/A` price sends that stock to
Exceptions.

## Rating and Sentiment rules

A value typed in the Universal Sheet always wins. With Gemini set up (below), a blank Rating / Sentiment is filled by
Gemini and saved into the sheet. Without Gemini, a blank Rating or Sentiment is filled from
`config/rating_rules.json` **only after `"approved": true`** is set there. The rules read revenue and PAT YoY from the
result line. Every run writes `rating_preview.md` with the rating each stock would get, so the rules can be reviewed first.

## Narrative slides

Slides 2-5 and 15 hold written analysis. Each run writes `runs/<period>/<client>/outputs/narrative_input.json`: every
computed fact, plus the list of fields with length limits. To fill them:

1. Draft the text from that file (for example with an AI chat), in the shape of `narrative_template.json`.
2. Save it as `runs/<period>/<client>/narrative.json` (or pass `--narrative`).
3. Run again.

Until then those slides show **PENDING NARRATIVE**. Every number in narrative.json must exist in
narrative_input.json, otherwise the run fails and lists the figure.

**Gemini: narrative, Rating and Sentiment.** Copy `.env.example` to `.env` (ignored by git), set
`NARRATIVE_PROVIDER=gemini` and fill in `NARRATIVE_API_KEY` (a Google AI Studio key). `NARRATIVE_MODEL` is optional:
one model or a comma-separated list in order of preference, e.g.
`gemini-3.5-flash, gemini-flash-latest, gemini-flash-lite-latest` (blank = a built-in list starting with
`gemini-pro-latest`). Each client run makes **one request**, which both rates the stocks whose
Rating / Sentiment is blank in the Universal Sheet and writes the narrative (client name removed). Views always
come from the Universal Sheet only. Gemini's ratings are written into the Universal Sheet (Rating, Sentiment, and
the Rating Source / Rating Reason columns) and kept as a backup in `runs/<period>/ai_ratings.json`; every later run,
for any client, reads them from there and sends no request for those stocks. A stock is rated again only when its
result line or commentary changes. If the sheet is open in Excel the write waits for a later run. The narrative is
saved in `runs/<period>/<client>/narrative.json` and reused. Edit the sheet or that file to correct anything;
delete the narrative (or a stock's Rating Source) to have it written again. There are no automatic
correction requests: if the narrative fails a check, the run names the field to edit. If a model is busy, gives
no answer within 2 minutes, is over quota or not available for the key, the same request goes straight to the
next model in the list; if all fail, the list is tried once more after 30 seconds. If no model answers, the Excel
and deck are still built (narrative slides pending) and the next run asks again. Set `narrative_provider` and
`rating_provider` to `"manual"` / `"rules"` in `config/settings.json` to never call the API.

## Broker layouts

Most exports work with no setup. If a file is not recognised, the error shows the first rows it saw. Add its column
names to `config/brokers/generic_equity_holdings.json`:

| Key | Meaning |
|---|---|
| `columns` | exact header names per field |
| `keywords` | whole words looked for inside a header |
| `exclude` | words that rule a header out (e.g. `Long Term Qty` is not the quantity) |
| `asof_regex` | reads the as-on date from title rows (`dd/mm/yyyy`, `dd-mm-yyyy` or `yyyy-mm-dd`) |

**Workbooks with several sheets** (e.g. `Equity`, `Mutual Funds`, `Combined`): the first sheet with a holdings
table is used, skipping sheets whose name contains fund / MF / combined / consolidated / summary / bond / debt / ETF.
The sheet used is shown in `report.md`. If a skipped funds sheet has holdings, the run warns that they are not in the review.

## Configuration

| File | Purpose |
|---|---|
| `config/settings.json` | price source, narrative checks, tolerances, table split sizes |
| `config/brokers/*.json` | broker column names |
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
