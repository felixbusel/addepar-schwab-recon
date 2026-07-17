#!/usr/bin/env python3
"""
Addepar <-> Schwab securities reconciliation.

Wealth managers commonly keep a securities master in a portfolio-management
platform (here: an Addepar "Portfolio View" export) while their custodian
(here: Charles Schwab) delivers its own security records as pipe-delimited flat
files. The same instrument is often described inconsistently across the two
systems: a CUSIP present in one but blank in the other, an Excel-mangled CUSIP
stored in scientific notation, a ticker with no ISIN, and so on.

This tool reconciles the two sources and — the main point — flags every
security that does NOT line up and explains why. For each security in the
Addepar export it attempts a hierarchical, most-specific-first match against
the Schwab records (ISIN -> CUSIP -> Ticker), then classifies the outcome:

  * matched, and how confidently (which identifiers agreed);
  * matched but with a data-quality conflict (e.g. the ISIN agrees but the
    tickers disagree - a stale ticker or the wrong share class);
  * unmatched, with the likely reason (held-away / not custodied at Schwab, a
    security not yet set up, or only a ticker to go on).

Supporting steps along the way repair CUSIPs that Excel has mangled into
scientific notation and backfill identifiers missing on either side.

Output is an Excel workbook: the cleaned Schwab data, the full reconciliation
(status + confidence + reason per security), and a Breaks sheet holding just the
exceptions an operator needs to work.

Requires: Python 3.9+, pandas, openpyxl.

Usage:
    python reconcile.py                        # runs against ./sample_data
    python reconcile.py --data-dir path/to/schwab/files \\
                        --addepar-file path/to/export.xlsx \\
                        --out-dir output

Inputs:
  --data-dir     Folder containing Schwab CRS<YYYYMMDD>.ULT / .ULN files.
  --addepar-file Excel export with a single "Portfolio View" sheet.
  --sheet        Sheet name in the Addepar export (default: "Portfolio View").
  --out-dir      Where to write the reconciled workbook + CSV.
  --date         Optional YYYYMMDD; otherwise the latest available file is used.
"""

import re
import sys
import argparse
import datetime as dt
from pathlib import Path

import pandas as pd

ADDEPAR_SHEET_DEFAULT = "Portfolio View"

# A CUSIP is exactly nine alphanumeric characters.
CUSIP_RE = re.compile(r"^[A-Z0-9]{9}$")


# ---------- small helpers ----------
def clean_id(x):
    """Uppercase and strip everything except A-Z / 0-9; empty -> NA."""
    if pd.isna(x):
        return pd.NA
    s = re.sub(r"[^A-Z0-9]", "", str(x).upper())
    return s if s else pd.NA


def clean_ticker(x):
    if pd.isna(x):
        return pd.NA
    s = str(x).strip().upper()
    return s if s else pd.NA


def safe_str(x):
    return "" if pd.isna(x) else str(x)


def is_valid_cusip(x):
    s = clean_id(x)
    return isinstance(s, str) and CUSIP_RE.match(s) is not None


def is_scientific_like(s):
    """True for values Excel has coerced into scientific notation, e.g. '3.7833E+08'."""
    s = safe_str(s).strip().upper()
    return bool(re.match(r"^\d+(\.\d+)?E\+\d+$", s))


def cusip_from_us_isin(isin):
    """A US ISIN embeds its CUSIP as characters 3-11 (e.g. US0378331005 -> 037833100)."""
    s = clean_id(isin)
    if isinstance(s, str) and len(s) == 12 and s.startswith("US"):
        return s[2:11]  # 9 chars
    return None


def most_common(series):
    """Return the modal (most frequently occurring) non-null value in a series."""
    s = series.dropna()
    if s.empty:
        return pd.NA
    return s.mode().iloc[0]


def pick_yesterdays(data_dir: Path, ext: str, date: str | None = None) -> Path:
    """
    Pick the Schwab file to use. Prefer an explicit --date, then yesterday's
    file, then the most recent CRS<date>.<ext> present in the directory.
    """
    if date:
        p = data_dir / f"CRS{date}.{ext.upper()}"
        if p.exists():
            return p
    y = (dt.date.today() - dt.timedelta(days=1)).strftime("%Y%m%d")
    y_path = data_dir / f"CRS{y}.{ext.upper()}"
    if y_path.exists():
        return y_path
    all_files = sorted(data_dir.glob(f"CRS*.{ext.upper()}"))
    if not all_files:
        raise FileNotFoundError(f"No CRS*.{ext} found in {data_dir}")

    def date_key(p: Path):
        m = re.search(r"CRS(\d{8})\." + re.escape(ext.upper()) + r"$", p.name)
        return m.group(1) if m else "00000000"

    all_files.sort(key=date_key, reverse=True)
    return all_files[0]


def normalize_header_tokens(h2: str, h3: str) -> str:
    """Schwab files use a two-row header; join the two levels into one label."""
    label = f"{h2.strip()} {h3.strip()}".strip()
    label = label.replace("Schwab#", "Schwab Nbr")
    label = re.sub(r"\s+", " ", label)
    return label


def make_match_key(label: str) -> str:
    """Fold a header label to a loose match key: lowercase, alphanumerics only."""
    k = label.lower().replace("#", " nbr")
    k = re.sub(r"[^a-z0-9]+", " ", k).strip()
    return k
# -----------------------------------


def parse_schwab_to_rows(path: Path, tax_status: str):
    """
    Yield {Ticker Symbol, CUSIP, Schwab Sec Nbr, ISIN, Tax Status} rows from a
    Schwab ULT/ULN flat file.

    The format is pipe-delimited with a two-line header: an H2| line (column
    groups) and an H3| line (column names). Data rows are prefixed DL|. Rather
    than assume fixed column positions, columns are located by fuzzy-matching
    their combined H2+H3 label, which keeps the parser resilient to layout
    changes and reordered columns.
    """
    with path.open("r", encoding="utf-8", errors="replace") as f:
        h2_line = h3_line = None
        for line in f:
            if line.startswith("H2|"):
                h2_line = line.rstrip("\n")
            elif line.startswith("H3|"):
                h3_line = line.rstrip("\n")
                break
        if not h2_line or not h3_line:
            raise ValueError(f"Could not find H2/H3 headers in {path.name}")

        h2_tokens = h2_line.split("|")[1:]
        h3_tokens = h3_line.split("|")[1:]
        maxlen = max(len(h2_tokens), len(h3_tokens))
        h2_tokens += [""] * (maxlen - len(h2_tokens))
        h3_tokens += [""] * (maxlen - len(h3_tokens))

        labels = [normalize_header_tokens(a, b) for a, b in zip(h2_tokens, h3_tokens)]
        keys = [make_match_key(lbl) for lbl in labels]

        def find_idx(patterns):
            for pat in patterns:
                for i, k in enumerate(keys):
                    if re.search(pat, k):
                        return i
            return None

        idx_ticker = find_idx([r"\bticker\s+symbol\b"])
        idx_cusip = find_idx([r"^cusip$"])
        idx_isin = find_idx([r"^isin$"])
        idx_schwab = find_idx([r"\bschwab(\s*sec)?\s*nbr\b", r"\bschwab\s*nbr\b"])

        if None in (idx_ticker, idx_cusip, idx_isin, idx_schwab):
            raise ValueError(
                f"{path.name}: missing one or more required columns.\n"
                f"ticker={idx_ticker}, cusip={idx_cusip}, schwab={idx_schwab}, isin={idx_isin}\n"
                f"Sample headers: {labels[:80]}"
            )

    with path.open("r", encoding="utf-8", errors="replace") as f:
        for line in f:
            if not line.startswith("DL|"):
                continue
            parts = line.rstrip("\n").split("|")[1:]
            if len(parts) < len(labels):
                parts += [""] * (len(labels) - len(parts))

            ticker = clean_ticker(parts[idx_ticker])
            cusip = clean_id(parts[idx_cusip])
            isin = clean_id(parts[idx_isin])
            schwab = str(parts[idx_schwab]).strip().lstrip("+")
            schwab = schwab if schwab else pd.NA

            if pd.isna(ticker) and pd.isna(cusip) and pd.isna(isin) and pd.isna(schwab):
                continue

            yield {
                "Ticker Symbol": ticker,
                "CUSIP": cusip,
                "Schwab Sec Nbr": schwab,
                "ISIN": isin,
                "Tax Status": tax_status,
            }


# -------- Addepar load (as text; CUSIP repair happens here) --------
def load_addepar(path: Path, sheet: str) -> pd.DataFrame:
    """
    Load the Addepar export as text (so leading zeros and long numeric IDs
    survive) and repair CUSIPs damaged by Excel along the way.
    """
    required = ["Entity ID", "Ticker Symbol", "CUSIP", "ISIN"]

    if not path.exists():
        print(f"Addepar file not found: {path} (skipping)")
        return pd.DataFrame(
            columns=["Addepar Entity ID", "Addepar Ticker", "Addepar Cusip",
                     "Addepar ISIN", "Addepar Cusip Repaired"]
        )

    adf = pd.read_excel(
        path,
        sheet_name=sheet,
        converters={  # force text
            "Entity ID":     lambda x: "" if x is None else str(x),
            "Ticker Symbol": lambda x: "" if x is None else str(x),
            "CUSIP":         lambda x: "" if x is None else str(x),
            "ISIN":          lambda x: "" if x is None else str(x),
        },
    )
    print("Addepar sheet columns detected:", list(adf.columns))
    for col in required:
        if col not in adf.columns:
            adf[col] = ""

    out = pd.DataFrame({
        "Addepar Entity ID": adf["Entity ID"].map(lambda x: str(x).strip() or pd.NA),
        "Addepar Ticker":    adf["Ticker Symbol"].map(clean_ticker),
        "Addepar ISIN":      adf["ISIN"].map(clean_id),
    })

    raw_cusip = adf["CUSIP"].astype(str)

    # Fix scientific-like or over-long numeric expansions via the US ISIN,
    # otherwise keep the pre-cleaned value (which may be NA). Also report
    # whether a repair happened, so it can be surfaced later as provenance.
    def fix_cusip(c_raw, isin_text):
        c_txt = safe_str(c_raw)
        c_pre = clean_id(c_txt)
        if is_valid_cusip(c_pre):
            return c_pre, False
        if is_scientific_like(c_txt) or (c_txt.isdigit() and len(c_txt) > 12):
            from_isin = cusip_from_us_isin(isin_text)
            if from_isin:
                return from_isin, True
        return c_pre, False  # may be NA

    fixed = [fix_cusip(c_raw, i_txt)
             for c_raw, i_txt in zip(raw_cusip.tolist(), out["Addepar ISIN"].tolist())]
    out["Addepar Cusip"] = [v for v, _ in fixed]
    out["Addepar Cusip Repaired"] = [r for _, r in fixed]

    return out


# -------- pull manual Schwab Sec# by Entity ID from the Portfolio View --------
def load_manual_secnum_from_portfolio(path: Path, sheet: str) -> dict:
    """
    Some Schwab security numbers are maintained by hand directly in the Addepar
    export. Read {Entity ID -> Schwab Security Number} for non-blank values so
    those can be used as a last-resort overlay.
    """
    if not path.exists():
        return {}
    try:
        df = pd.read_excel(path, sheet_name=sheet, dtype=str)
    except Exception:
        return {}
    if "Entity ID" not in df.columns or "Schwab Security Number" not in df.columns:
        return {}
    df["Entity ID"] = df["Entity ID"].astype(str).str.strip()
    df["Schwab Security Number"] = (
        df["Schwab Security Number"].astype(str).str.strip().str.lstrip("+")
    )
    df = df[(df["Entity ID"] != "") & (df["Schwab Security Number"] != "")]
    df = df.drop_duplicates(subset=["Entity ID"], keep="first")
    return dict(zip(df["Entity ID"], df["Schwab Security Number"]))


# -------- Build Schwab lookup dicts (keyed by identifier combinations) --------
def build_schwab_maps(df_norm: pd.DataFrame):
    """
    Build seven lookup tables keyed by every useful combination of ISIN, CUSIP
    and Ticker. Where a key maps to several Schwab rows, the modal value of each
    remaining field is kept, so noisy duplicates collapse to their consensus.
    """
    sch = df_norm.rename(columns={
        "Ticker Symbol": "Ticker",
        "Schwab Sec Nbr": "SecNbr",
    }).copy()

    def make_map(keys):
        cols_to_agg = ["ISIN", "CUSIP", "Ticker", "SecNbr"]
        agg_dict = {c: most_common for c in cols_to_agg if c not in keys}
        g = sch.dropna(subset=[k for k in keys]).groupby(keys, dropna=False)
        out = g.agg(agg_dict).reset_index()

        def val_from_key(keys_list, key_tuple, name):
            if name in keys_list:
                return key_tuple[keys_list.index(name)]
            return pd.NA

        recs = {}
        for _, r in out.iterrows():
            key_tuple = tuple(r[k] for k in keys)
            recs[key_tuple] = {
                "Schwab ISIN":    r["ISIN"]   if "ISIN"   in out.columns else val_from_key(keys, key_tuple, "ISIN"),
                "Schwab Cusip":   r["CUSIP"]  if "CUSIP"  in out.columns else val_from_key(keys, key_tuple, "CUSIP"),
                "Schwab Ticker":  r["Ticker"] if "Ticker" in out.columns else val_from_key(keys, key_tuple, "Ticker"),
                "Schwab Sec Nbr": r["SecNbr"] if "SecNbr" in out.columns else pd.NA,
            }
        return recs

    return {
        "ict": make_map(["ISIN", "CUSIP", "Ticker"]),
        "ic":  make_map(["ISIN", "CUSIP"]),
        "it":  make_map(["ISIN", "Ticker"]),
        "ct":  make_map(["CUSIP", "Ticker"]),
        "i":   make_map(["ISIN"]),
        "c":   make_map(["CUSIP"]),
        "t":   make_map(["Ticker"]),
    }


# -------- Fill Schwab fields for a single Addepar row --------
# Human-readable labels for the identifier combination that produced a match.
BASIS_LABEL = {
    "ict": "ISIN + CUSIP + Ticker",
    "ic":  "ISIN + CUSIP",
    "it":  "ISIN + Ticker",
    "ct":  "CUSIP + Ticker",
    "i":   "ISIN",
    "c":   "CUSIP",
    "t":   "Ticker only",
}


def fill_from_schwab(ad_isin, ad_cusip, ad_ticker, maps):
    """
    Hierarchical, most-specific-first lookup. Try the tightest key available
    (all three identifiers), then progressively looser keys, stopping at the
    first hit. Returns (filled_fields, basis_code) where basis_code names the
    key that matched (or None if the security wasn't found). The basis is the
    match-confidence signal: a three-identifier hit is far stronger than a
    ticker-only hit.
    """
    res = {"Schwab ISIN": pd.NA, "Schwab Cusip": pd.NA, "Schwab Ticker": pd.NA, "Schwab Sec Nbr": pd.NA}
    i = ad_isin if pd.notna(ad_isin) else None
    c = ad_cusip if pd.notna(ad_cusip) else None
    t = ad_ticker if pd.notna(ad_ticker) else None

    attempts = [
        ("ict", (i, c, t), all((i, c, t))),
        ("ic",  (i, c),    all((i, c))),
        ("it",  (i, t),    all((i, t))),
        ("ct",  (c, t),    all((c, t))),
        ("i",   (i,),      i is not None),
        ("c",   (c,),      c is not None),
        ("t",   (t,),      t is not None),
    ]
    for code, key, ok in attempts:
        if not ok:
            continue
        rec = maps[code].get(key)
        if rec:
            for k, v in rec.items():
                res[k] = v
            return res, code
    return res, None


# -------- Classify every row: matched / unmatched, confidence, and reason --------
def classify_matches(m: pd.DataFrame) -> pd.DataFrame:
    """
    The core of the tool. For each reconciled security, assign:
      * Match Status  — Matched / Matched (overlay) / Unmatched
      * Match Basis   — which identifiers agreed (the confidence level)
      * Issue         — for anything that isn't a clean match, the underlying
                        reason: an unmatched break, a low-confidence match, or a
                        data-quality conflict between the two sources
      * Notes         — benign provenance (what was repaired or backfilled)
    """
    statuses, bases, issues, notes = [], [], [], []
    for _, r in m.iterrows():
        basis = r["_basis"]
        sec = r["Schwab Security Number"]
        a_tk, s_tk = r["Addepar Ticker"], r["Schwab Ticker"]
        a_cu, s_cu = r["Addepar Cusip"], r["Schwab Cusip"]
        a_is, s_is = r["Addepar ISIN"], r["Schwab ISIN"]

        matched_files = isinstance(basis, str)
        matched_overlay = (not matched_files) and pd.notna(sec) and safe_str(sec) != ""

        note = []
        if r["_repaired"]:
            note.append("CUSIP repaired from ISIN")
        if r["_cusip_backfilled"]:
            note.append("CUSIP backfilled from Schwab")
        if r["_isin_from_schwab"]:
            note.append("ISIN taken from Schwab")

        issue = []
        if matched_files:
            status = "Matched"
            basis_label = BASIS_LABEL.get(basis, basis)
            # Data-quality conflicts, flagged even though the security matched:
            if pd.notna(a_tk) and pd.notna(s_tk) and clean_ticker(a_tk) != clean_ticker(s_tk):
                issue.append(f"ticker differs (Addepar {a_tk} vs Schwab {s_tk})")
            if is_valid_cusip(a_cu) and is_valid_cusip(s_cu) and clean_id(a_cu) != clean_id(s_cu):
                issue.append(f"CUSIP differs (Addepar {a_cu} vs Schwab {s_cu})")
            if pd.notna(a_is) and pd.notna(s_is) and clean_id(a_is) != clean_id(s_is):
                issue.append(f"ISIN differs (Addepar {a_is} vs Schwab {s_is})")
            if basis == "t":
                issue.append("low-confidence match (ticker only)")
        elif matched_overlay:
            status = "Matched (overlay)"
            basis_label = "Manual overlay"
            issue.append("not in custodian files; Schwab # supplied by hand")
        else:
            status = "Unmatched"
            basis_label = "\u2014"
            if not (r["_has_isin"] or r["_has_cusip"] or r["_has_ticker"]):
                issue.append("no usable identifiers")
            elif r["_has_isin"] or r["_has_cusip"]:
                issue.append("not present in custodian files (held-away or not yet set up)")
            else:
                issue.append("ticker not found in custodian files")

        statuses.append(status)
        bases.append(basis_label)
        issues.append("; ".join(issue))
        notes.append("; ".join(note))

    m = m.copy()
    m["Match Status"] = statuses
    m["Match Basis"] = bases
    m["Issue"] = issues
    m["Notes"] = notes
    return m.drop(columns=[c for c in m.columns if c.startswith("_")])


# -------- Addepar Entity ID mapping (modal) --------
def build_entity_maps(ad: pd.DataFrame):
    """Same hierarchical-key idea, used to recover a missing Addepar Entity ID."""
    d = ad.copy()
    d["Entity"] = d["Addepar Entity ID"]

    def make_map(keys):
        sub = (
            d.dropna(subset=[k for k in keys])
            .groupby(keys, dropna=False)["Entity"]
            .agg(most_common)
            .reset_index()
        )
        recs = {}
        for _, r in sub.iterrows():
            key = tuple(r[k] for k in keys)
            recs[key] = r["Entity"]
        return recs

    return {
        "ict": make_map(["Addepar ISIN", "Addepar Cusip", "Addepar Ticker"]),
        "ic":  make_map(["Addepar ISIN", "Addepar Cusip"]),
        "it":  make_map(["Addepar ISIN", "Addepar Ticker"]),
        "ct":  make_map(["Addepar Cusip", "Addepar Ticker"]),
        "i":   make_map(["Addepar ISIN"]),
        "c":   make_map(["Addepar Cusip"]),
        "t":   make_map(["Addepar Ticker"]),
    }


def fill_entity(ad_isin, ad_cusip, ad_ticker, emaps):
    i = ad_isin if pd.notna(ad_isin) else None
    c = ad_cusip if pd.notna(ad_cusip) else None
    t = ad_ticker if pd.notna(ad_ticker) else None
    if i and c and t:
        v = emaps["ict"].get((i, c, t))
        if v is not None:
            return v
    if i and c:
        v = emaps["ic"].get((i, c))
        if v is not None:
            return v
    if i and t:
        v = emaps["it"].get((i, t))
        if v is not None:
            return v
    if c and t:
        v = emaps["ct"].get((c, t))
        if v is not None:
            return v
    if i:
        v = emaps["i"].get((i,))
        if v is not None:
            return v
    if c:
        v = emaps["c"].get((c,))
        if v is not None:
            return v
    if t:
        v = emaps["t"].get((t,))
        if v is not None:
            return v
    return pd.NA


# -------- Main --------
def parse_args():
    p = argparse.ArgumentParser(description="Reconcile an Addepar export against Schwab ULT/ULN files.")
    p.add_argument("--data-dir", type=Path, default=Path("sample_data"),
                   help="Folder with Schwab CRS<YYYYMMDD>.ULT/.ULN files.")
    p.add_argument("--addepar-file", type=Path, default=Path("sample_data/addepar_export.xlsx"),
                   help="Addepar Excel export.")
    p.add_argument("--sheet", default=ADDEPAR_SHEET_DEFAULT, help="Sheet name in the Addepar export.")
    p.add_argument("--out-dir", type=Path, default=Path("output"), help="Where to write outputs.")
    p.add_argument("--date", default=None, help="Optional YYYYMMDD to pin a specific delivery.")
    return p.parse_args()


def main():
    args = parse_args()
    data_dir = args.data_dir
    if not data_dir.exists():
        raise FileNotFoundError(f"Data directory not found: {data_dir}")
    args.out_dir.mkdir(parents=True, exist_ok=True)

    # Pick the latest (or requested) Schwab files
    paths = []
    try:
        paths.append(("ULT", pick_yesterdays(data_dir, "ULT", args.date)))
    except FileNotFoundError:
        pass
    try:
        paths.append(("ULN", pick_yesterdays(data_dir, "ULN", args.date)))
    except FileNotFoundError:
        pass
    if not paths:
        raise FileNotFoundError(f"No CRS*.ULT or CRS*.ULN files found in {data_dir}")

    print("Using files:")
    for kind, p in paths:
        print(f" - {kind}: {p}")

    # Parse Schwab lots (ULT = taxable, ULN = non-taxable)
    all_rows = []
    for kind, p in paths:
        tax = "taxable" if kind == "ULT" else "nontaxable"
        before = len(all_rows)
        all_rows.extend(list(parse_schwab_to_rows(p, tax_status=tax)))
        print(f"  {kind}: parsed {len(all_rows) - before} DL rows")
    if not all_rows:
        print("No DL rows parsed from either file.")
        return

    df = pd.DataFrame(all_rows).replace("", pd.NA).dropna(how="all")

    # Derive an output date stamp from the source file names
    dates = []
    for _, p in paths:
        m = re.search(r"CRS(\d{8})\.(ULT|ULN)$", p.name)
        dates.append(m.group(1) if m else "")
    date_part = dates[0] if dates and all(x == dates[0] for x in dates) else dt.date.today().strftime("%Y%m%d")
    out_csv = args.out_dir / f"Schwab_Clean_{date_part}_ULT_ULN.csv"
    out_xlsx = args.out_dir / f"Schwab_Clean_{date_part}_ULT_ULN.xlsx"
    df.to_csv(out_csv, index=False)

    # Normalize Schwab IDs for matching
    df_norm = df.copy()
    df_norm["Ticker Symbol"] = df_norm["Ticker Symbol"].map(clean_ticker)
    df_norm["CUSIP"] = df_norm["CUSIP"].map(clean_id)
    df_norm["ISIN"] = df_norm["ISIN"].map(clean_id)
    df_norm["Schwab Sec Nbr"] = df_norm["Schwab Sec Nbr"].map(
        lambda x: str(x).strip() if pd.notna(x) else pd.NA
    )

    # Build Schwab modal lookup maps
    sch_maps = build_schwab_maps(df_norm)

    # Load Addepar (as text; CUSIP repaired inside)
    ad = load_addepar(args.addepar_file, args.sheet)
    ad["Addepar Ticker"] = ad["Addepar Ticker"].map(clean_ticker)
    ad["Addepar Cusip"] = ad["Addepar Cusip"].map(clean_id)
    ad["Addepar ISIN"] = ad["Addepar ISIN"].map(clean_id)

    # Unique Addepar securities drive the output rows
    ad_uni = ad[["Addepar ISIN", "Addepar Cusip", "Addepar Ticker",
                 "Addepar Entity ID", "Addepar Cusip Repaired"]].copy()
    ad_uni = (
        ad_uni.dropna(how="all", subset=["Addepar ISIN", "Addepar Cusip", "Addepar Ticker"])
        .drop_duplicates()
        .reset_index(drop=True)
    )

    eid_maps = build_entity_maps(ad)

    # Enrich each Addepar security from the Schwab maps, tracking what happened
    enriched = []
    for _, r in ad_uni.iterrows():
        ai, ac, at = r["Addepar ISIN"], r["Addepar Cusip"], r["Addepar Ticker"]

        sch, basis = fill_from_schwab(ai, ac, at, sch_maps)

        # Backfill a bad/missing Addepar CUSIP from Schwab when possible
        ad_cusip_final = ac
        cusip_backfilled = False
        if not is_valid_cusip(ad_cusip_final) and is_valid_cusip(sch["Schwab Cusip"]):
            ad_cusip_final = clean_id(sch["Schwab Cusip"])
            cusip_backfilled = True

        # Note when the only ISIN available came from the Schwab side
        isin_from_schwab = pd.isna(ai) and pd.notna(sch["Schwab ISIN"])

        # Recover a missing Entity ID via the hierarchical maps
        eid = r["Addepar Entity ID"]
        if pd.isna(eid) or safe_str(eid) == "":
            eid = fill_entity(ai, ad_cusip_final, at, eid_maps)

        enriched.append({
            "Addepar Entity ID":      eid,
            "Schwab Security Number": sch["Schwab Sec Nbr"],
            "Addepar Ticker":         at,
            "Schwab Ticker":          sch["Schwab Ticker"],
            "Addepar Cusip":          ad_cusip_final,
            "Schwab Cusip":           sch["Schwab Cusip"],
            "Addepar ISIN":           ai,
            "Schwab ISIN":            sch["Schwab ISIN"],
            # helper fields consumed by classify_matches(), then dropped
            "_basis":            basis,
            "_repaired":         bool(r["Addepar Cusip Repaired"]),
            "_cusip_backfilled": cusip_backfilled,
            "_isin_from_schwab": bool(isin_from_schwab),
            "_has_isin":         pd.notna(ai),
            "_has_cusip":        is_valid_cusip(ad_cusip_final),
            "_has_ticker":       pd.notna(at),
        })

    out_match = pd.DataFrame(enriched)

    # ---------- OVERLAY: hand-maintained Schwab Sec# from the Portfolio View ----------
    manual_from_portfolio = load_manual_secnum_from_portfolio(args.addepar_file, args.sheet)
    if manual_from_portfolio:
        mapped_sec = out_match["Addepar Entity ID"].astype(str).map(manual_from_portfolio)
        out_match["Schwab Security Number"] = out_match["Schwab Security Number"].fillna(mapped_sec)
    # -----------------------------------------------------------------------------------

    # ---------- CLASSIFY: status, confidence, and the reason for every break ----------
    out_match = classify_matches(out_match)

    # The exceptions list: anything that isn't a clean, unambiguous match
    needs_review = (out_match["Match Status"] != "Matched") | (out_match["Issue"].str.len() > 0)
    breaks = out_match[needs_review].reset_index(drop=True)

    # Diagnostics first, then the side-by-side identifiers for auditing
    col_order = [
        "Addepar Entity ID", "Match Status", "Match Basis", "Issue", "Notes",
        "Schwab Security Number",
        "Addepar Ticker", "Schwab Ticker",
        "Addepar Cusip", "Schwab Cusip",
        "Addepar ISIN", "Schwab ISIN",
    ]
    out_match = out_match[col_order]
    breaks = breaks[col_order]

    # Write the workbook: cleaned Schwab data, full reconciliation, and the breaks
    with pd.ExcelWriter(out_xlsx, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Schwab Clean")
        out_match.to_excel(writer, index=False, sheet_name="Reconciliation")
        breaks.to_excel(writer, index=False, sheet_name="Breaks")

    # ---------- Summary ----------
    n = len(out_match)
    print(f"\nWrote outputs:\n - {out_csv}\n - {out_xlsx}")
    print(f"\nReconciled {n} Addepar securities.")
    print("\nMatch status:")
    for status, cnt in out_match["Match Status"].value_counts().items():
        print(f"  {status:18s} {cnt}")
    if len(breaks):
        print(f"\n{len(breaks)} securities need review — reasons:")
        for issue, cnt in breaks["Issue"].value_counts().items():
            print(f"  [{cnt}] {issue}")
    print("\nPreview (Reconciliation):")
    with pd.option_context("display.max_columns", None, "display.width", 220):
        print(out_match.head(20).to_string(index=False))


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print("ERROR:", e, file=sys.stderr)
        sys.exit(1)
