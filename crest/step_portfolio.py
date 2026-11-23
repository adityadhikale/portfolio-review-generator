"""Step A: the Portfolio Excel, its LibreOffice cross-check copy and the validation gate."""
from __future__ import annotations

from . import deck_data as dd
from . import validate as V
from .excel_build import build_workbook
from .office import find_soffice, recalculate_copy
from .shared import Ctx, _py_as_excel, gate
from .util import hosted


def build_portfolio(ctx: Ctx) -> dict:
    """Write '<client> Portfolio.xlsx'; return its recalculated numbers (the deck is built from these)."""
    res, settings, per, out = ctx.res, ctx.settings, ctx.per, ctx.out
    client, asof, bp = ctx.client, ctx.asof, ctx.hf
    priced, matched, exc_rows, deck_group_of = ctx.priced, ctx.matched, ctx.exc_rows, ctx.deck_group_of
    # 4. Excel
    ctx.stage("excel")
    xlsx = out / f"{client} Portfolio.xlsx"
    layout = build_workbook(xlsx, client=client, asof=asof, period_label=per.label, parse=bp,
                            priced=priced, snapshot_rows=matched, exception_rows=exc_rows, deck_group_of=deck_group_of)
    res.outputs["excel"] = xlsx

    soffice = find_soffice(settings.get("libreoffice_path", ""))
    res.info["libreoffice"] = soffice or "not found"
    py = dd.python_values(priced)
    if soffice:
        ctx.stage("check")
        # The delivered workbook carries Google-Sheets GOOGLEFINANCE formulas (Market Cap, CMP) that LibreOffice
        # cannot evaluate (and would rewrite). So a check copy with this run's values in those two columns is
        # recalculated instead; every other formula is identical. The delivered file is left untouched.
        check = out / "_recalc" / f"check_{xlsx.name}"
        check.parent.mkdir(parents=True, exist_ok=True)
        build_workbook(check, client=client, asof=asof, period_label=per.label, parse=bp,
                       priced=priced, snapshot_rows=matched, exception_rows=exc_rows, deck_group_of=deck_group_of,
                       google_finance=False)
        rc = recalculate_copy(check, check, soffice)
        errors = dd.scan_errors(rc)
        excel_vals = dd.read_portfolio(rc, layout)
        res.checks.append(V.excel_errors(errors))
        res.checks.append(V.python_vs_excel(dd.compare(excel_vals, py)))
        res.info["excel_cached_values"] = ("no - Excel / Google Sheets calculate on open; Market Cap and CMP are "
                                           "GOOGLEFINANCE formulas (live in Google Sheets, this run's values elsewhere)")
    elif hosted():
        res.notes.append("Hosted website: no LibreOffice here, so the Excel was not recalculated as a cross-check; "
                         "every number was computed in Python from the same inputs. Excel calculates on open.")
        excel_vals = _py_as_excel(py, matched, priced, deck_group_of)
    else:
        res.checks.append(V.Check("No Excel error values after recalculation", "WARN",
                                  "LibreOffice not found - workbook not recalculated; values computed in Python "
                                  "from the same inputs. Excel will calculate on open (fullCalcOnLoad)."))
        excel_vals = _py_as_excel(py, matched, priced, deck_group_of)

    ctx.xlsx, ctx.layout, ctx.soffice = xlsx, layout, soffice
    gate(ctx, excel_vals, "Validation gate failed - see the checks table. Excel was written for inspection; "
                          "deck not built.")
    return excel_vals
