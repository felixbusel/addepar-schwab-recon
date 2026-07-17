# Addepar ↔ Schwab Securities Reconciliation

A utility that reconciles a **securities master exported from Addepar** against
**position files delivered by Charles Schwab**, and — the point of the whole
thing — **flags every security that doesn't line up and explains why.**

It's a daily headache in wealth-management operations. The portfolio-management
platform and the custodian describe the *same* instruments differently, and the
real job isn't "join the two lists" — it's *finding the breaks*: securities
tracked in Addepar that the custodian has no record of, ones that match on an
ISIN but disagree on the ticker, ones held at another custodian entirely.
Multiply that across a couple of thousand holdings and eyeballing it stops being
viable.

So for every security the tool doesn't just match-or-not. It assigns a **status**,
records **how confidently** it matched (which identifiers agreed), and for anything
that isn't a clean match, writes down the **underlying reason**. The exceptions are
split into their own sheet — the worklist an operator actually acts on.

> **At a glance.** [`docs/reconciliation_logic.pdf`](docs/reconciliation_logic.pdf)
> is a visual walkthrough — the end-to-end data flow, the decision logic that
> classifies each security into matched / flagged / unmatched with a reason, and
> worked examples from the sample data.

> **Note on data.** Everything in `sample_data/` is synthetic. The securities are
> well-known public instruments and all internal identifiers (Entity IDs, Schwab
> security numbers) are fabricated. No real firm, client, account, or position
> data is included anywhere in this repository.

---

## What it produces

An Excel workbook with three sheets (plus a CSV of the cleaned Schwab data for audit):

| Sheet | Contents |
|-------|----------|
| **Reconciliation** | Every Addepar security with its `Match Status`, `Match Basis` (which identifiers agreed), `Issue` (the reason, if any), `Notes` (provenance), and the Addepar-vs-Schwab values side by side |
| **Breaks** | Just the exceptions — anything unmatched, flagged, or resolved only by hand. The operator's worklist |
| **Schwab Clean** | The parsed, normalized Schwab records |

Every security lands in one of three buckets:

- **Matched** — found in the custodian files. `Match Basis` records the confidence:
  a three-identifier agreement (`ISIN + CUSIP + Ticker`) is far stronger than a
  `Ticker only` hit.
- **Matched (overlay)** — not in the custodian files, but a Schwab number was
  supplied by hand in the Addepar export.
- **Unmatched** — no custodian record found.

And for anything needing attention, the `Issue` column says *why*:

| Reason | What it means |
|--------|---------------|
| `not present in custodian files (held-away or not yet set up)` | Good identifiers, no Schwab record — typically an asset custodied elsewhere, or a position not yet onboarded |
| `ticker not found in custodian files` | Only a ticker to match on, and it isn't at the custodian |
| `ticker differs (Addepar X vs Schwab Y)` | Matched on ISIN/CUSIP, but the tickers disagree — usually a stale ticker or the wrong share class |
| `low-confidence match (ticker only)` | Matched, but on ticker alone — worth a human glance |
| `not in custodian files; Schwab # supplied by hand` | Resolved via the manual overlay rather than an automated match |

The reason list is a plain if/elif ladder in `classify_matches()` and is easy to extend.

## How it works

The hard part isn't joining two lists. It's (a) matching confidently when no
single identifier is reliably populated on both sides, and (b) turning every
"no match" into an actionable reason.

**Hierarchical matching — with a confidence score.** Seven lookup tables are built
from the Schwab side, keyed by every useful combination of ISIN, CUSIP and Ticker
(`ISIN+CUSIP+Ticker → ISIN+CUSIP → ISIN+Ticker → CUSIP+Ticker →` each identifier
alone). Each security tries the tightest key its data supports and falls back to
looser keys, stopping at the first hit. The tier that matched is recorded as the
**Match Basis** — the confidence signal that distinguishes a solid three-identifier
agreement from a shaky ticker-only one.

**Classifying the outcome — the actual point.** A match or non-match on its own
isn't useful; the reason is. After the lookup, each security is labelled
*matched* / *matched (overlay)* / *unmatched*, and then:

- **conflicts are caught even among matches** — if a security matches on ISIN and
  CUSIP but the two sources disagree on the ticker, that's flagged (a stale ticker
  or wrong share class), not silently accepted;
- **unmatched securities are diagnosed by the identifiers they *did* have** — an
  instrument with a valid ISIN but no custodian record is almost certainly
  held-away or not yet onboarded, a very different break from one where all we had
  was a ticker the custodian doesn't recognise.

This is the step that turns a raw join into a reconciliation report someone can act
on. It lives in `classify_matches()`.

The remaining pieces are supporting cleanup that keeps the matching honest:

**CUSIP repair from the ISIN.** A US ISIN is `US` + the 9-character CUSIP + a check
digit, so when a CUSIP arrives blank, over-long, or mangled into scientific notation
(`3.7833E+08`), it's rebuilt from characters 3–11 of the ISIN
(`US0378331005 → 037833100`). Surfaced in `Notes`, not treated as a headline.

**Backfilling missing identifiers.** A CUSIP present in Schwab but blank in Addepar
(or an ISIN only Schwab has) is copied across so the side-by-side comparison is
complete — also recorded in `Notes`.

**Modal ("most-common") collapse.** The same security can appear on several Schwab
lots with minor inconsistencies; each lookup key collapses its matches to the
*modal* value of every field, so duplicates become a consensus instead of an error.

**Layout-resilient header parsing.** Schwab files use a two-row header (`H2|` groups
+ `H3|` names) and columns move between deliveries, so columns are located by
fuzzy-matching their combined label rather than by fixed position.

## Input formats

**Schwab** — pipe-delimited flat files named `CRS<YYYYMMDD>.ULT` (taxable) and
`CRS<YYYYMMDD>.ULN` (non-taxable), with a two-line header and `DL|`-prefixed data
rows. Required columns (matched by label): Ticker Symbol, CUSIP, ISIN, and the
Schwab security number.

**Addepar** — an Excel export with a single `Portfolio View` sheet containing:
`Security`, `Entity ID`, `Schwab Security Number`, `Ticker Symbol`, `ISIN`,
`CUSIP`. Everything is read as text so leading zeros and long numeric IDs survive.

## Usage

```bash
pip install -r requirements.txt

# Run against the bundled synthetic sample data:
python reconcile.py

# Or point it at your own inputs:
python reconcile.py \
    --data-dir      path/to/schwab/files \
    --addepar-file  path/to/addepar_export.xlsx \
    --out-dir       output
```

Options:

| Flag | Default | Description |
|------|---------|-------------|
| `--data-dir` | `sample_data` | Folder containing the Schwab `.ULT` / `.ULN` files |
| `--addepar-file` | `sample_data/addepar_export.xlsx` | Addepar Excel export |
| `--sheet` | `Portfolio View` | Sheet name in the export |
| `--out-dir` | `output` | Where the reconciled workbook + CSV are written |
| `--date` | latest available | Optional `YYYYMMDD` to pin a specific delivery |

When no `--date` is given, the script uses the requested date if present, then
yesterday's file, then the most recent `CRS*` file in the directory.

## Regenerating the sample data

The sample files are checked in so the demo runs out of the box, but the
generator is included and documents exactly how each edge case is constructed:

```bash
cd sample_data
python make_sample_data.py
```

Running `reconcile.py` on this data produces a mix of clean matches and five
deliberate breaks — a held-away security, a ticker the custodian doesn't
recognise, a stale-ticker conflict (Addepar `FB` vs Schwab `META`), a ticker-only
low-confidence match, and an overlay-only resolution — each landing in the Breaks
sheet with its reason. The console prints a status-and-reason summary.

## Requirements

Python 3.9+, `pandas`, `openpyxl`.
