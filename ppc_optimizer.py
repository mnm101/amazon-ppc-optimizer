#!/usr/bin/env python3
"""
Amazon PPC Bid & Placement Optimizer (suggestions only)
=======================================================
Reads a keyword performance file and a placement performance file
(covering one or many campaigns) and generates bid / placement-adjustment
RECOMMENDATIONS in a review workbook. It never changes anything in Amazon.

Usage:
    python ppc_optimizer.py                          (uses the input/ folder - see below)
    python ppc_optimizer.py keywords.xlsx placements.xlsx [options]
    python ppc_optimizer.py combined.xlsx            (one file w/ 'Keywords' & 'Placements' tabs)

Standard workflow: rename the two Amazon exports to input/keywords.xlsx and
input/placements.xlsx, then run with no arguments. Results go to outputs/recs.xlsx.

Options:
    -o, --output PATH      Output workbook path (default: ppc_recommendations.xlsx)
    -t, --target FLOAT     Target ACOS as a fraction (default: 0.29 for 28-30% goal)
    --targets-file PATH    Optional CSV with columns: Campaign ID, Target ACOS
                           (per-campaign override of the global target)
    --max-cpc FLOAT        Hard effective-CPC ceiling in dollars (default 1.85)
    --window-days INT      Days the export covers; drives the late-conversion
                           allowance (default 7)

Assumptions (documented in the output 'Settings & Notes' tab):
    * All campaigns run Dynamic bids - DOWN ONLY (per account owner).
      Effective max CPC = base bid x (1 + placement adjustment).
    * Sales are derived as Cost / ACOS when no Sales column exists.
    * Sponsored Products placement adjustments are NON-NEGATIVE: 0%..+900%.
      A placement can be bid UP from the base bid but never below it. The value
      itself moves freely in both directions inside that range, so a too-high
      adjustment IS lowered when that alone fixes the placement. What is
      impossible is going below 0% - so a placement already AT 0% can only be
      suppressed by cutting base bids and compensating the others (that is lam).
      (Negative settings on Sponsored Brands rows are held, never raised.)

Bids and placement adjustments are solved as ONE coupled system, because the
effective CPC at a placement is base bid x (1 + that placement's adjustment):

    * Each campaign's base-bid factor (lam) and ALL its placement adjustments
      are solved as one equation system: new adj_p = (1+adj_p)(1+pi_p)/lam - 1,
      lam = min(1, min_p (1+adj_p)(1+pi_p)). lam < 1 IS suppression via base
      bids; the compensating raises on the other placements come from the same
      equation, so effective bids are held exactly. Keyword moves compose as
      base = bid x (1 + m) x lam - the two levers cannot compound by accident.
      The "Effective Change" column reports the delivered effective move.
    * Short windows understate conversions (orders still attributing to recent
      clicks), so sales are matured once up-front and every threshold, tier and
      ceiling reads the matured figures.
    * Nothing is acted on inside a tolerance band around target - being a couple
      of points off target is not a signal.
    * Move size scales with the evidence behind it; thin samples move at half
      step. Over-target converters step down gradually and never below their
      revenue-justified bid; a cut can never raise a bid.
    * Every move is bounded per run: merit raises +20% effective, suppression
      15% effective, bid steps 10-30% - all reviewable, converging over cycles.
    * A hard effective-CPC ceiling outranks every other rule.
    * Thin data => HOLD/WATCH, never a cut or pause on noise.
"""

import argparse
import math
import os
import sys
from dataclasses import dataclass, field

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ============================================================================
# TUNABLE CONSTANTS - edit these to change behavior
# ============================================================================
TARGET_ACOS_DEFAULT = 0.29     # global target ACOS (fraction). 28-30% goal -> 0.29
MIN_BID = 0.02                 # Amazon minimum bid
PAUSE_SPEND_MULT = 2.0         # pause zero-sale kw when spend >= 2x Target CPA ...
PAUSE_CLICK_FACTOR = 1.5       # ... AND clicks >= 1.5x expected clicks-per-order
MIN_IMPR_VISIBILITY = 50       # below this many impressions, kw is "starved"
NO_TRAFFIC_RAISE = 0.15        # suggested raise for starved keywords (+15%)

MIN_PLACEMENT_CLICKS = 10      # min clicks before judging a placement
PLACEMENT_STEP = 0.20          # raise a winning placement's adj by +20 points
# --- File convention: run with no arguments and these are used. Paths are
# resolved relative to THIS script, so it works from any working directory.
INPUT_DIR = "input"
DEFAULT_KEYWORDS = "keywords.xlsx"
DEFAULT_PLACEMENTS = "placements.xlsx"
DEFAULT_OUTPUT = "outputs/recs.xlsx"

OFF_AMAZON_ACOS_FLAG = 0.40    # Off-Amazon flagged only when ACOS > 40% (or ACOS=0
                               # while still spending). Otherwise it is ignored.
# --- Engine behaviour -------------------------------------------------------
# Guiding principles:
#   (1) short windows understate conversions -> mature the data before judging,
#   (2) act on evidence, not on noise -> step size scales with sample size,
#   (3) never act at a boundary -> tolerance bands on both sides of target,
#   (4) every move is small and self-stepping, converging over runs.
PLACEMENT_MAX = 1.00           # cap a placement adjustment at +100%
PLACEMENT_MIN = 0.00           # SP adjustments are NON-NEGATIVE (0%..+900%): a
                               # placement can be bid up from base, never below it.
                               # The value is still freely lowerable within range -
                               # what is impossible is going below 0%, so a placement
                               # already at 0% can only be pushed down by cutting the
                               # BASE BID and compensating the others (see lam).
SUPPRESS_STEP = 0.15           # max EFFECTIVE suppression of one placement per run
MAX_EFFECTIVE_CPC = 1.85       # hard ceiling: base bid x (1 + adjustment) is never
                               # allowed to exceed this at any placement (--max-cpc).
# Late-conversion maturity. A report pulled over the last N days has not yet been
# credited with the orders that will still attribute to its most recent clicks, so
# observed Sales are understated and observed ACOS is overstated. Roughly one day's
# worth of conversions is still in flight; over a 7-day window that is ~15%.
# Sales are grossed up by 1/(1-allowance) ONCE, so matured ACOS, RPC and every
# downstream decision stay mutually consistent.
WINDOW_DAYS_DEFAULT = 7        # data window length in days (--window-days)
LATE_CONV_DAYS_PENDING = 1.05  # days of conversions still in flight at pull time
LATE_CONV_ALLOWANCE_MAX = 0.20 # never assume more than 20% of sales are pending
# Tolerance bands around target (applied to MATURED ACOS). Nothing is touched
# inside the band - being a couple of points off target is not a signal.
RAISE_BUFFER = 0.80            # raise only when matured ACOS <= target x 0.80
LOWER_BUFFER = 1.20            # cut only when matured ACOS >= target x 1.20
STRONG_WINNER_RATIO = 0.50     # matured ACOS < 50% of target -> bigger raise step
RAISE_STEP = 0.10              # standard raise (+10%)
RAISE_STEP_STRONG = 0.20       # raise for a far-under-target winner (+20%)
# Over-target converters step DOWN gradually, sized by how far matured ACOS sits
# over target, one notch harder when spend is heavy.
OVERTARGET_CUT_TIERS = ((1.5, 0.10), (2.0, 0.15), (3.0, 0.20))  # (acos/target, cut)
OVERTARGET_CUT_MAX = 0.30      # cut when matured ACOS is more than 3x target
HIGH_SPEND_MULT = 2.0          # spend >= this x Target CPA -> one notch harder
HIGH_SPEND_BUMP = 0.05
MAX_BID_CHANGE = 0.30          # hard cap on any single-run bid change
# Evidence gating. A decision backed by a thin sample moves at half step, and a
# zero-order keyword is not judged until it has spent past a buffer on Target CPA.
EVIDENCE_ORDERS = 3            # orders needed for a full-size converter move
EVIDENCE_CLICKS = 25           # clicks needed for a full-size zero-order move
THIN_STEP_FACTOR = 0.50        # thin-sample moves are halved
ZERO_SALE_BUFFER = 1.25        # act on a zero-order keyword only past 1.25x CPA
DEAD_PLACEMENT_CLICK_MULT = 1.5  # a 0-order placement is only "dead" once it has
                                 # >= 1.5x the clicks an order normally takes
LOW_CLICK_CLICKS = 10          # active kw/target with <= this many clicks in the
LOW_CLICK_BUMP = 0.05          # window gets a +5% exposure bump (never overrides a
                               # cut/pause or a larger raise).
# ============================================================================

KW_REQUIRED = ["Campaign Name", "Campaign ID", "Ad Group", "Target", "Match Type",
               "Impressions", "Clicks", "Conversions", "Cost", "ACOS", "Current Bid"]
PL_REQUIRED = ["Campaign Name", "Campaign ID", "Placement",
               "Impressions", "Clicks", "Conversions", "Cost", "ACOS",
               "Current Adjustment"]

# Native Amazon export columns -> canonical names the engine uses.
# {canonical: [accepted aliases]}. Old hand-built exports already use the
# canonical names, so harmonization is a no-op for them.
KW_ALIASES = {
    "Campaign Name": ["Campaign name"],
    "Ad Group": ["Ad group"],
    "Match Type": ["Match type"],
    "Conversions": ["Orders"],
    "Cost": ["Spend"],
    "Current Bid": ["Bid"],
    "Sales": ["Ad Sales"],
    # optional pass-through so a row can be found/actioned in the ads platform
    "Target ID": ["Keyword/Target ID", "Keyword ID", "Targeting ID"],
    "Ad Group ID": ["Ad group ID"],
}
# Identifiers are carried through as TEXT: they run to 15+ digits, and Excel
# silently reformats numbers that long into scientific notation or drops
# precision past 15 significant figures - which would break copy-paste.
ID_COLS = ("Target ID", "Ad Group ID")
PL_ALIASES = {
    "Campaign Name": ["Campaign name"],
    "Conversions": ["Orders"],
    "Cost": ["Spend"],
    "Current Adjustment": ["Bid adj. %"],
    "Sales": ["Ad Sales"],
}

PLACEMENT_LABELS = {
    "top_of_search": "Top of Search",
    "rest_of_search": "Rest of Search",
    "product_page": "Product Pages",
    "off_amazon": "Off-Amazon",
}

# Raw placement label (stripped, lower-cased) -> canonical category.
# "__IGNORE__" rows are dropped from the recommendations entirely.
PLACEMENT_ALIASES = {
    "top of search on-amazon": "top_of_search",
    "top of search": "top_of_search",
    "top_of_search": "top_of_search",
    "detail page on-amazon": "product_page",
    "product page": "product_page",
    "product pages": "product_page",
    "product_page": "product_page",
    "other on-amazon": "rest_of_search",
    "rest of search": "rest_of_search",
    "rest_of_search": "rest_of_search",
    "off amazon": "off_amazon",
    "off-amazon": "off_amazon",
    "off_amazon": "off_amazon",
    "other placements": "__IGNORE__",
}


# ============================================================================
# Loading & validation
# ============================================================================
def _here(*parts) -> str:
    """A path relative to this script, so defaults do not depend on the cwd."""
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), *parts)


def resolve_default_inputs(kw_arg, pl_arg):
    """Fall back to the input/ folder convention when paths are not given."""
    if kw_arg is not None:
        return kw_arg, pl_arg
    kw = _here(INPUT_DIR, DEFAULT_KEYWORDS)
    pl = _here(INPUT_DIR, DEFAULT_PLACEMENTS)
    missing = [p for p in (kw, pl) if not os.path.isfile(p)]
    if missing:
        fail("No input files given, and the default ones are missing:\n       "
             + "\n       ".join(missing)
             + f"\n\n       Rename your two Amazon exports and drop them in the "
               f"'{INPUT_DIR}' folder:\n"
               f"         {INPUT_DIR}/{DEFAULT_KEYWORDS}    <- the keyword/target "
               f"export (kt_export_*.xlsx)\n"
               f"         {INPUT_DIR}/{DEFAULT_PLACEMENTS}  <- the placement export "
               f"(pl_export_*.xlsx)\n"
               f"       then run:  python ppc_optimizer.py\n"
               f"       Or pass paths explicitly: python ppc_optimizer.py KW.xlsx "
               f"PL.xlsx")
    return kw, pl


def fail(msg: str):
    print(f"ERROR: {msg}", file=sys.stderr)
    sys.exit(1)


def harmonize_columns(df: pd.DataFrame, aliases: dict, is_kw: bool) -> pd.DataFrame:
    """Rename native Amazon-export columns to the canonical names the engine
    uses, so both the owner's hand-built schema and raw report exports load.
    Canonical names already present are left untouched (old files unaffected)."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    for canon, alts in aliases.items():
        if canon in df.columns:
            continue
        for alt in alts:
            if alt in df.columns:
                df = df.rename(columns={alt: canon})
                break
    # Keyword target may be split across 'Keyword' (blank for auto/ASIN targets)
    # and 'Targeting value'. Build a single canonical 'Target' column.
    if is_kw and "Target" not in df.columns:
        kw_col = df["Keyword"] if "Keyword" in df.columns else None
        tv_col = df["Targeting value"] if "Targeting value" in df.columns else None
        if kw_col is not None and tv_col is not None:
            df["Target"] = kw_col.fillna(tv_col).fillna("(auto/ASIN target)")
        elif kw_col is not None:
            df["Target"] = kw_col.fillna("(auto/ASIN target)")
        elif tv_col is not None:
            df["Target"] = tv_col.fillna("(auto/ASIN target)")
    if is_kw and "Match Type" in df.columns:
        df["Match Type"] = df["Match Type"].fillna("AUTO")
    return df


def load_sheet(path: str, sheet, required: list, label: str,
               aliases: dict, is_kw: bool) -> pd.DataFrame:
    try:
        df = pd.read_excel(path, sheet_name=sheet) if sheet else pd.read_excel(path)
    except Exception as e:
        fail(f"Could not read {label} data from '{path}': {e}")
    df = harmonize_columns(df, aliases, is_kw)
    missing = [c for c in required if c not in df.columns]
    if missing:
        fail(f"{label} file '{path}' is missing required column(s): {missing}\n"
             f"       Found columns: {list(df.columns)}")
    return df


def load_inputs(kw_path: str, pl_path: str | None):
    """Two files, or one workbook containing 'Keywords' and 'Placements' tabs."""
    if pl_path is None:
        try:
            sheets = pd.ExcelFile(kw_path).sheet_names
        except Exception as e:
            fail(f"Could not open '{kw_path}': {e}")
        if "Keywords" in sheets and "Placements" in sheets:
            kw = load_sheet(kw_path, "Keywords", KW_REQUIRED, "Keyword", KW_ALIASES, True)
            pl = load_sheet(kw_path, "Placements", PL_REQUIRED, "Placement", PL_ALIASES, False)
            return kw, pl
        # Native single-file exports name the sheets differently; accept the
        # keyword/target tab here and require a separate placements file.
        fail("Only one file given and it does not contain both a 'Keywords' and a "
             "'Placements' sheet. Provide two files: keywords.xlsx placements.xlsx")
    kw = load_sheet(kw_path, None, KW_REQUIRED, "Keyword", KW_ALIASES, True)
    pl = load_sheet(pl_path, None, PL_REQUIRED, "Placement", PL_ALIASES, False)
    return kw, pl


def _clean_id(v) -> str:
    """An identifier as exact text. Excel/pandas may have read a 15-digit ID as a
    float, so 2.9051584403233e+14 and 290515844032329.0 must both come back as
    '290515844032329'."""
    if pd.isna(v):
        return ""
    if isinstance(v, float) and float(v).is_integer():
        return str(int(v))
    return str(v).strip()


def normalize(df: pd.DataFrame, is_kw: bool) -> pd.DataFrame:
    df = df.copy()
    num_cols = ["Impressions", "Clicks", "Conversions", "Cost", "ACOS"]
    num_cols += ["Current Bid"] if is_kw else ["Current Adjustment"]
    # A blank keyword bid means "inherits the ad group default", NOT zero. Remember
    # that before filling, so the engine can refuse to do percentage math on it.
    if is_kw:
        bid_num = pd.to_numeric(df["Current Bid"], errors="coerce")
        df["_bid_missing"] = bid_num.isna() | (bid_num <= 0)
        for c in ID_COLS:                       # exact text, never float-mangled
            if c in df.columns:
                df[c] = df[c].map(_clean_id)
    for c in num_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    # ACOS may arrive as 15.67 (percent) instead of 0.1567 (fraction)
    nonzero = df.loc[df["ACOS"] > 0, "ACOS"]
    if len(nonzero) and nonzero.median() > 1.5:
        df["ACOS"] = df["ACOS"] / 100.0
    # Adjustment may arrive as 10 (percent) instead of 0.10
    if not is_kw:
        adj = df.loc[df["Current Adjustment"] > 0, "Current Adjustment"]
        if len(adj) and adj.max() > 9.0:  # >900% impossible as fraction
            df["Current Adjustment"] = df["Current Adjustment"] / 100.0
    # Derived metrics ------------------------------------------------------
    if "Sales" in df.columns:                       # prefer a real Sales column
        df["Sales"] = pd.to_numeric(df["Sales"], errors="coerce").fillna(0.0)
        df["_sales_src"] = "column"
    else:
        df["Sales"] = df.apply(
            lambda r: r["Cost"] / r["ACOS"] if r["ACOS"] > 0 else 0.0, axis=1)
        df["_sales_src"] = "derived (Cost/ACOS)"
    df["_data_issue"] = df.apply(
        lambda r: "Conversions>0 but ACOS=0 - revenue unknown"
        if (r["Conversions"] > 0 and r["ACOS"] == 0) else "", axis=1)
    df["CPC"] = df.apply(lambda r: r["Cost"] / r["Clicks"] if r["Clicks"] else 0, axis=1)
    df["CVR"] = df.apply(lambda r: r["Conversions"] / r["Clicks"] if r["Clicks"] else 0, axis=1)
    df["CTR"] = df.apply(lambda r: r["Clicks"] / r["Impressions"] if r["Impressions"] else 0, axis=1)
    return df


# ============================================================================
# Campaign-level statistics (fallback hierarchy for CVR / AOV)
# ============================================================================
@dataclass
class CampaignStats:
    clicks: float = 0
    conversions: float = 0
    cost: float = 0
    sales: float = 0
    cvr: float = 0.0
    aov: float = 0.0
    acos: float = 0.0
    target: float = TARGET_ACOS_DEFAULT
    target_cpa: float = 0.0
    clicks_per_order: float = 0.0
    notes: list = field(default_factory=list)


def campaign_stats(kw: pd.DataFrame, targets: dict, global_target: float,
                   account_cvr: float, account_aov: float) -> dict:
    out = {}
    for cid, g in kw.groupby("Campaign ID"):
        s = CampaignStats(
            clicks=g["Clicks"].sum(), conversions=g["Conversions"].sum(),
            cost=g["Cost"].sum(), sales=g["Sales"].sum())
        s.cvr = s.conversions / s.clicks if s.clicks else 0.0
        s.aov = s.sales / s.conversions if s.conversions else 0.0
        s.acos = s.cost / s.sales if s.sales else 0.0
        # fallbacks to account level when the campaign itself is thin
        if s.cvr == 0:
            s.cvr = account_cvr
            s.notes.append("campaign CVR unavailable - using account CVR")
        if s.aov == 0:
            s.aov = account_aov
            s.notes.append("campaign AOV unavailable - using account AOV")
        s.target = targets.get(str(cid), global_target)
        s.target_cpa = s.aov * s.target
        s.clicks_per_order = (1.0 / s.cvr) if s.cvr > 0 else float("inf")
        out[cid] = s
    return out


# ============================================================================
# Keyword rules engine
# ============================================================================
def _kw_rec_row(r: pd.Series, rec: dict) -> dict:
    row = {
        "Campaign Name": r["Campaign Name"], "Campaign ID": r["Campaign ID"],
        "Ad Group": r["Ad Group"], "Keyword": r["Target"],
        "Match Type": r["Match Type"],
    }
    # identifiers, when the export carries them - for finding the row in the
    # ads platform and applying the change against the right target
    for c in ID_COLS:
        if c in r.index:
            row[c] = r[c]
    row.update({
        "Impressions": int(r["Impressions"]), "Clicks": int(r["Clicks"]),
        "Orders": int(r["Conversions"]), "Spend": round(r["Cost"], 2),
        "Sales": round(r["Sales"], 2),
        "ACOS": round(r["ACOS"], 4), "CVR": round(r["CVR"], 4),
        "CPC": round(r["CPC"], 2),
        "Current Bid": round(r["Current Bid"], 2),
        "Recommended Bid": rec["new_bid"],
        "Effective Change": round(rec.get("eff_change", 0.0), 4),
        "Action": rec["action"], "Confidence": rec["confidence"],
        "Change Capped": "YES" if rec["capped"] else "",
        "Reason": rec["reason"] + (f" [{r['_data_issue']}]" if r["_data_issue"] else ""),
    })
    return row


# ============================================================================
# Placement rules engine
# ============================================================================
def canonicalize_placements(pl: pd.DataFrame):
    """Map raw export placement labels to canonical categories and split off
    the rows that get special treatment.

    Returns (engine_rows, off_amazon_rows, ignored_rows):
      * engine_rows  - Top of Search / Rest of Search / Product Pages, fed to
                       the placement solver (Placement replaced with the
                       canonical key so labels and Mode-B detection work).
      * off_amazon   - handled by flag_off_amazon() (mostly ignored).
      * ignored      - 'Other Placements' etc., dropped entirely.
    'Top of Search' and 'Top of Search on-Amazon' both fold into top_of_search.
    """
    pl = pl.copy()
    raw = pl["Placement"].astype(str).str.strip().str.lower()
    pl["Placement"] = raw.map(PLACEMENT_ALIASES).fillna(raw)
    off = pl[pl["Placement"] == "off_amazon"].copy()
    ignored = pl[pl["Placement"] == "__IGNORE__"].copy()
    engine = pl[~pl["Placement"].isin(["off_amazon", "__IGNORE__"])].copy()
    return engine, off, ignored


def flag_off_amazon(off: pd.DataFrame, stats: dict) -> pd.DataFrame:
    """Off-Amazon is ignored unless it is clearly leaking money: ACOS above the
    flag threshold, or spending with zero sales. Only flagged rows are emitted;
    no bid change is suggested (review/exclude decision is the seller's)."""
    rows = []
    for _, r in off.iterrows():
        s = stats.get(r["Campaign ID"])
        t = s.target if s else TARGET_ACOS_DEFAULT
        acos, spend = r["ACOS"], r["Cost"]
        if not (acos > OFF_AMAZON_ACOS_FLAG or (acos == 0 and spend > 0)):
            continue
        if acos == 0:
            reason = (f"Off-Amazon spent ${spend:.2f} with 0 sales - review or "
                      f"exclude this placement.")
        else:
            reason = (f"Off-Amazon ACOS {acos:.1%} exceeds the "
                      f"{OFF_AMAZON_ACOS_FLAG:.0%} flag threshold - review or "
                      f"exclude this placement.")
        cur = r["Current Adjustment"]
        rows.append(_pl_row(r, t, cur, cur, "REVIEW - OFF-AMAZON",
                            "HIGH" if r["Clicks"] >= MIN_PLACEMENT_CLICKS else "LOW",
                            reason))
    return pd.DataFrame(rows)


def _pl_row(r, target, cur, new, action, conf, reason):
    label = PLACEMENT_LABELS.get(str(r["Placement"]).strip().lower(),
                                 str(r["Placement"]))
    return {
        "Campaign Name": r["Campaign Name"], "Campaign ID": r["Campaign ID"],
        "Placement": label,
        "Impressions": int(r["Impressions"]), "Clicks": int(r["Clicks"]),
        "Orders": int(r["Conversions"]), "Spend": round(r["Cost"], 2),
        "ACOS": round(r["ACOS"], 4), "CPC": round(r["CPC"], 2),
        "Target ACOS": target,
        "Current Adjustment": round(cur, 2), "Recommended Adjustment": round(new, 2),
        "Action": action, "Confidence": conf, "Reason": reason,
    }


# ============================================================================
# Joint bid + placement solve
# ============================================================================
# Bid and placement adjustments are one coupled system: effective CPC at a
# placement = base bid x (1 + adjustment). Two facts drive the solve:
#   1. The observed keyword RPC (Sales/Clicks) was earned on clicks whose cost
#      already includes the placement uplift, so the target BASE bid must be
#      DEFLATED by the campaign's click-weighted average adjustment:
#          base = RPC x target / (1 + weighted_avg_adjustment)
#      (Modes A/B omit this divisor - fine when adjustments are small, a large
#       over-bid when they are +100%+.) The per-run +/-MAX_BID_CHANGE clamp
#      phases the correction in gradually rather than in one jump.
#   2. Suppression is incremental, not a one-shot swing. A bad placement is
#      reined in via its OWN adjustment first (-20 pts/run toward 0). Once it
#      hits 0% (Amazon has no negative adjustment) it is suppressed indirectly:
#      profitable placements step UP +20 pts/run to pull spend toward them, and
#      base bids deflate as those adjustments rise. Everything moves in small
#      steps and converges over a few export cycles - the same self-stepping
#      philosophy as the keyword bids, so each run is reviewable.
#
# Data limit: Amazon exports carry no keyword x placement performance, so each
# keyword's placement mix is approximated by its CAMPAIGN's click mix. Exact at
# campaign level, approximate per keyword - unavoidable without a report Amazon
# does not produce.
def _campaign_adj_weights(g: pd.DataFrame) -> dict:
    """Click-share weight per placement within a campaign (from placement sheet)."""
    total = g["Clicks"].sum()
    if total <= 0:
        return {}
    return {row["Placement"]: row["Clicks"] / total for _, row in g.iterrows()}


def _dead_placement_clicks(s: CampaignStats) -> float:
    """Clicks a 0-order placement must have before it can be called dead: enough
    that an order would normally have happened by now. Getting 0 orders in 11
    clicks at a 7% CVR is ordinary noise, not evidence of a bad placement."""
    cpo = s.clicks_per_order if s and s.clicks_per_order != float("inf") else 0.0
    return max(float(MIN_PLACEMENT_CLICKS), DEAD_PLACEMENT_CLICK_MULT * cpo)


def desired_placement_effect(r: pd.Series, s: CampaignStats, t: float,
                             allowance: float, dead_floor: float):
    """The EFFECTIVE bid change this placement deserves on its own merit, as a
    multiplier offset (+0.20 = push its effective bid 20% higher, -0.15 = 15%
    lower). Judged on MATURED ACOS inside the same tolerance band as keywords.

    Returns (pi, cls, matured_acos).
    """
    clicks, conv, cur = r["Clicks"], r["Conversions"], r["Current Adjustment"]
    _, acos, _ = mature(r["Cost"], r["Sales"], clicks, allowance, r["ACOS"])
    if clicks < MIN_PLACEMENT_CLICKS:
        return 0.0, "thin", acos
    if conv > 0 and acos <= t * RAISE_BUFFER:
        return PLACEMENT_STEP, "strong", acos
    if conv > 0 and acos < t * LOWER_BUFFER:
        return 0.0, "buffer", acos
    if conv > 0:                                  # over target: steer toward target
        return max(t / acos - 1.0, -SUPPRESS_STEP), "over", acos
    if clicks >= dead_floor:                      # 0 orders with real traffic
        return -SUPPRESS_STEP, "dead", acos
    return 0.0, "thin", acos                      # 0 orders but ordinary noise


def solve_campaign_placements(g: pd.DataFrame, s: CampaignStats,
                              allowance: float = 0.0):
    """Solve one campaign's base-bid factor and ALL its placement adjustments as a
    single system.

    Sponsored Products placement adjustments are NON-NEGATIVE (0%..+900%): a
    placement can be bid up from the base bid but never below it. A too-high
    adjustment can therefore simply be LOWERED when that alone fixes the placement.
    But one already at the 0% floor cannot go further, and the only remaining lever
    is the owner's strategy: cut the BASE BIDS and COMPENSATE the other placements
    upward, leaving the bad placement to absorb the cut.

    Writing the desired effective bid at placement p as
        E_p = base x (1 + a_p) x (1 + pi_p)
    (pi_p = the effective change placement p deserves on its own merit) and
    requiring base' x (1 + a'_p) = E_p for every p SIMULTANEOUSLY gives, for one
    campaign-level constant lam,
        base'    = base x lam          (further x (1 + m) per keyword, later)
        1 + a'_p = (1 + a_p)(1 + pi_p) / lam
    The floor a'_p >= 0, with minimal churn, fixes lam:
        lam = min(1, min_p (1 + a_p)(1 + pi_p))
    lam < 1 exactly when some placement must fall below the base level - that IS
    suppression via base bids, and the other placements' compensating raises come
    out of the same equation, so effective bids are held exactly. lam = 1 when
    nothing needs suppressing, so healthy campaigns do not churn.

    Grandfathered negative settings (Sponsored Brands rows in the same export) are
    held as-is and never raised; their constraint is (1 + pi_p) alone, so the base
    cut still delivers their reduction gradually.

    Returns (rows, lam, max_adj).
    """
    t = s.target if s else TARGET_ACOS_DEFAULT
    dead_floor = _dead_placement_clicks(s)
    dec = []                                  # [row, cur, pi, cls, acos]
    for _, r in g.iterrows():
        pi, cls, acos = desired_placement_effect(r, s, t, allowance, dead_floor)
        dec.append([r, r["Current Adjustment"], pi, cls, acos])

    # The one campaign-level degree of freedom: how far base bids must come down.
    # Each placement's constraint is the base factor it would need if its own
    # adjustment were pushed to its floor; lam is the tightest of them (capped at 1).
    cons, drivers = [1.0], []
    for row, cur, pi, cls, _ in dec:
        if cur > PLACEMENT_MAX:               # legacy over-cap rows step down instead
            continue
        floor_ = min(0.0, cur)                # grandfathered negatives keep their level
        v = (1 + cur) * (1 + pi) / (1 + floor_) if 1 + floor_ > 0 else 1.0
        if v > 0:
            cons.append(v)
            drivers.append((row["Placement"], v))
    lam = min(cons)
    suppressing = lam < 1 - 1e-9
    # the placement(s) whose own floor forced the cut - the cause, not a passenger
    driver_set = {p for p, v in drivers if abs(v - lam) < 1e-9} if suppressing else set()
    supp_note = (f" Base bids cut {1 - lam:.0%} campaign-wide to suppress the bad "
                 f"placement." if suppressing else "")

    rows, adjs = [], []
    for r, cur, pi, cls, acos in dec:
        conf = _confidence(r["Conversions"], r["Clicks"])
        spent = f"{int(r['Clicks'])} clicks, ${r['Cost']:.2f}"
        if cur > PLACEMENT_MAX:               # legacy setting above the cap
            new = max(PLACEMENT_MAX, round(cur - PLACEMENT_STEP, 2))
            reason = (f"Adjustment {cur:.0%} exceeds the +{PLACEMENT_MAX:.0%} "
                      f"ceiling - step down {PLACEMENT_STEP:.0%} pts toward it.")
            row = _pl_row(r, t, cur, new, "LOWER ADJUSTMENT", conf, reason)
            row["Effective Change"] = round((1 + new) * lam / (1 + cur) - 1, 4)
            rows.append(row); adjs.append(new)
            continue
        if cur < 0:                           # grandfathered negative (SB row)
            new = cur
            reason = (f"Current adjustment {cur:.0%} is below the 0% floor this "
                      f"engine works in (Sponsored Brands setting) - held as set, "
                      f"never raised.{supp_note}")
            row = _pl_row(r, t, cur, new, "HOLD", conf, reason)
            row["Effective Change"] = round(lam - 1, 4)
            rows.append(row); adjs.append(new)
            continue
        ideal = (1 + cur) * (1 + pi) / lam - 1
        if pi < -1e-9:
            ideal = min(ideal, cur)           # never raise a bad/over placement
        new = min(max(round(ideal, 2), PLACEMENT_MIN), PLACEMENT_MAX)
        capped = ideal > PLACEMENT_MAX + 1e-9
        adjs.append(new)
        cap_note = (f" Compensation capped at +{PLACEMENT_MAX:.0%}, so this "
                    f"placement's effective bid slips too - consider splitting the "
                    f"campaign if it persists." if capped else "")
        if cls == "strong":
            action = "RAISE ADJUSTMENT" if new > cur else "HOLD"
            comp = (f" (includes compensation for the {1 - lam:.0%} base cut)"
                    if suppressing else "")
            reason = (f"Converting at {acos:.1%} matured ACOS, well under target "
                      f"{t:.1%} - lean in {pi:+.0%} effective{comp}.{cap_note}")
        elif pi < -1e-9:                      # over-target or dead
            what = (f"Matured ACOS {acos:.1%} over target {t:.1%}" if cls == "over"
                    else f"{spent}, 0 orders (past the ~{dead_floor:.0f} clicks an "
                         f"order takes here)")
            if new < cur - 1e-9:
                action = "LOWER ADJUSTMENT"
                reason = (f"{what} - push its effective bid {pi:+.0%}; its own "
                          f"adjustment ({cur:.0%}) is high enough to absorb that on "
                          f"its own, so no base-bid cut is needed for it.{cap_note}")
            elif r["Placement"] in driver_set:
                action = "SUPPRESS (VIA BASE BIDS)"
                reason = (f"{what} - needs {pi:+.0%} effective but its adjustment is "
                          f"already at the {cur:.0%} floor and cannot go lower. THIS "
                          f"is why base bids are cut {1 - lam:.0%} campaign-wide, with "
                          f"the other placements compensated so only this placement "
                          f"loses ground. Steps down again each run while it stays "
                          f"bad.{cap_note}")
            else:
                action = "SUPPRESS (BASE CUT, ADJ HELD)"
                reason = (f"{what} - needs {pi:+.0%} effective. Another placement in "
                          f"this campaign is at its 0% floor and already forces a "
                          f"{1 - lam:.0%} base-bid cut, which delivers this "
                          f"placement's full {pi:+.0%} on its own - so its "
                          f"{cur:.0%} adjustment is held rather than lowered "
                          f"(lowering it too would over-suppress past the "
                          f"{SUPPRESS_STEP:.0%}/run limit).{cap_note}")
        elif suppressing and new != cur:
            action = "COMPENSATE ADJUSTMENT"
            reason = (f"Base bids cut {1 - lam:.0%} to suppress another placement - "
                      f"raise this one to (1{cur:+.0%})/{lam:.2f}-1 = {new:.0%} so "
                      f"its effective bid is held where it was.{cap_note}")
        else:
            action = "HOLD"
            if cls == "buffer":
                reason = (f"Matured ACOS {acos:.1%} within target tolerance "
                          f"({t * RAISE_BUFFER:.1%}-{t * LOWER_BUFFER:.1%}) - at "
                          f"target, hold.")
            elif r["Conversions"] == 0 and r["Clicks"] >= MIN_PLACEMENT_CLICKS:
                reason = (f"Only {int(r['Clicks'])} clicks and 0 orders - fewer than "
                          f"the ~{dead_floor:.0f} clicks an order normally takes "
                          f"here, so this is noise, not a verdict.")
            else:
                reason = f"Only {int(r['Clicks'])} clicks - insufficient data."
        row = _pl_row(r, t, cur, new, action, conf, reason)
        row["Effective Change"] = round((1 + new) * lam / (1 + cur) - 1, 4)
        rows.append(row)
    return rows, lam, (max(adjs) if adjs else 0.0)


def solve_placements(pl: pd.DataFrame, stats: dict, allowance: float = 0.0):
    """Solve every campaign's placements. Returns
    (placement recommendations, {Campaign ID: (lam, max_adj)}) where lam is the
    campaign-wide base-bid factor the keyword engine must apply and max_adj is the
    highest recommended adjustment (the binding placement for the CPC ceiling)."""
    rows, coeffs = [], {}
    for cid, g in pl.groupby("Campaign ID"):
        pl_rows, lam, max_adj = solve_campaign_placements(g, stats.get(cid), allowance)
        rows.extend(pl_rows)
        coeffs[cid] = (lam, max_adj)
    return pd.DataFrame(rows), coeffs


def late_conv_allowance(window_days: int) -> float:
    """Fraction of conversions still in flight for a window of this length."""
    if window_days <= 0:
        return 0.0
    return min(LATE_CONV_ALLOWANCE_MAX, LATE_CONV_DAYS_PENDING / window_days)


def mature(cost: float, sales: float, clicks: float, allowance: float,
           reported_acos: float = 0.0):
    """Gross up observed sales for conversions not yet attributed, then derive the
    matured ACOS and RPC from it. One transformation, so every downstream number
    stays consistent. Returns (sales, acos, rpc).

    Prefers Amazon's own ACOS figure over Cost/Sales: exports round Cost to whole
    dollars, which can throw a recomputed ACOS off by several points on low-spend
    rows - enough to flip a decision near a threshold."""
    m_sales = sales / (1.0 - allowance) if 0.0 < allowance < 1.0 else sales
    keep = (1.0 - allowance) if 0.0 < allowance < 1.0 else 1.0
    if reported_acos > 0:
        m_acos = reported_acos * keep
    else:
        m_acos = cost / m_sales if m_sales > 0 else 0.0
    m_rpc = m_sales / clicks if clicks else 0.0
    return m_sales, m_acos, m_rpc


def _evidence_full(orders: float, clicks: float) -> bool:
    """Is the sample big enough to justify a full-size move?"""
    return orders >= EVIDENCE_ORDERS or clicks >= EVIDENCE_CLICKS


def _scale_step(step: float, orders: float, clicks: float) -> float:
    """Halve a step when the sample behind it is thin."""
    return step if _evidence_full(orders, clicks) else step * THIN_STEP_FACTOR


def _confidence(orders: float, clicks: float) -> str:
    if orders >= 5 and clicks >= 30:
        return "HIGH"
    if orders >= 2 or clicks >= EVIDENCE_CLICKS:
        return "MEDIUM"
    return "LOW"


def _overtarget_cut(acos: float, t: float, spend: float, target_cpa: float) -> float:
    """Gentle, tiered step-down for over-target converters. The further
    matured ACOS is over target the bigger the cut; heavy spend nudges it a notch."""
    ratio = acos / t if t else 0.0
    cut = OVERTARGET_CUT_MAX
    for bound, c in OVERTARGET_CUT_TIERS:
        if ratio <= bound:
            cut = c
            break
    if target_cpa and spend >= HIGH_SPEND_MULT * target_cpa:
        cut = min(cut + HIGH_SPEND_BUMP, MAX_BID_CHANGE)
    return cut


def clamp_bid(new_bid: float, current: float) -> tuple[float, bool]:
    """Bid clamp: cap the single-run change at +/-MAX_BID_CHANGE. Rounds
    to the cent INSIDE the bounds, so the stated cap is never exceeded by rounding."""
    lo, hi = current * (1 - MAX_BID_CHANGE), current * (1 + MAX_BID_CHANGE)
    capped = new_bid < lo or new_bid > hi
    val = round(min(max(new_bid, lo), hi), 2)
    if val < lo:                                  # rounded past the floor -> ceil
        val = math.ceil(lo * 100 - 1e-9) / 100
    elif val > hi:                                # rounded past the cap -> floor
        val = math.floor(hi * 100 + 1e-9) / 100
    return max(MIN_BID, val), capped


def _reconcile_action(rec: dict, current_bid: float) -> dict:
    """Keep Action honest about the EFFECTIVE move (rec['eff']) - what the keyword
    actually experiences at the compensated placements. In a suppressing campaign
    the base bid falls while compensating adjustments hold exposure, so 'base went
    down' is not the economic story."""
    if rec["action"] in ("PAUSE", "REVIEW - NO BID IN EXPORT", "REVIEW - NO TRAFFIC"):
        return rec
    eff = rec.get("eff", 0.0)
    if eff > 1e-6:
        rec["action"] = "RAISE BID"
    elif eff < -1e-6:
        rec["action"] = "LOWER BID"
    elif rec["new_bid"] < current_bid - 1e-9:
        rec["action"] = "LOWER BASE (COMPENSATED)"
    elif rec["action"] in ("RAISE BID", "LOWER BID"):
        rec["action"] = "HOLD"
    return rec


def optimize_keyword(r: pd.Series, s: CampaignStats, lam: float = 1.0,
                     allowance: float = 0.0) -> dict:
    """Keyword bid engine, working in EFFECTIVE-CPC terms.

    The keyword's observed CPC already contains its campaign's placement
    adjustments, so the decision is a multiplier m on the effective bid: to land
    blended ACOS on target the effective CPC must scale by target/ACOS (the
    revenue-justified limit); the actual move is a gentle, evidence-sized step
    toward it.

    The result composes EXACTLY with the placement solve's base factor lam:
        new base = bid x (1 + m) x lam
    and since the compensating adjustments carry 1/lam, the keyword's effective
    bid at every surviving placement moves by exactly (1 + m), while a suppressed
    placement gets (1 + m)(1 + pi_p). The two levers cannot compound by accident -
    they are two variables of one equation system.
    """
    t = s.target
    bid, clicks, conv = r["Current Bid"], r["Clicks"], r["Conversions"]
    spend = r["Cost"]
    m_sales, acos, rpc = mature(spend, r["Sales"], clicks, allowance, r["ACOS"])
    rec = dict(action="HOLD", new_bid=bid, confidence="HIGH", reason="",
               capped=False, eff=0.0)
    mat_note = (f" [sales matured +{allowance / (1 - allowance):.0%} for pending "
                f"conversions]" if allowance > 0 and m_sales > 0 else "")
    lam_note = (f" Base bid also x{lam:.2f}: this campaign is suppressing a "
                f"placement via base bids; the compensating adjustments hold this "
                f"keyword's exposure at the surviving placements."
                if lam < 1 - 1e-9 else "")

    def finish(m, action, conf, reason):
        m = max(-MAX_BID_CHANGE, min(MAX_BID_CHANGE, m))
        nb = max(MIN_BID, round(bid * (1 + m) * lam, 2))
        rec.update(action=action, new_bid=nb, confidence=conf, eff=m,
                   capped=abs(m) >= MAX_BID_CHANGE - 1e-9,
                   reason=reason + lam_note)
        return rec

    # No bid in the export (auto/ASIN targets inherit the ad group default).
    if r.get("_bid_missing", False):
        justified = rpc * t if rpc > 0 else 0.0
        detail = (f"Data justifies about ${justified:.2f} effective "
                  f"(matured ACOS {acos:.1%})." if justified > 0
                  else f"{int(clicks)} clicks, {int(conv)} orders - no revenue "
                       f"signal to price it from.")
        rec.update(action="REVIEW - NO BID IN EXPORT", new_bid=0.0,
                   confidence=_confidence(conv, clicks),
                   reason=f"No bid in the export - this target inherits its ad "
                          f"group's default bid. {detail} Set it at the ad group, "
                          f"or give the target an explicit bid{mat_note}.")
        return rec

    if conv > 0 and m_sales > 0 and clicks == 0:
        rec.update(confidence="LOW",
                   reason=f"{int(conv)} order(s), ${m_sales:.2f} sales but 0 clicks "
                          f"in window (late attribution). No per-click bid math "
                          f"possible; hold.")
        return rec

    if conv > 0 and m_sales > 0:
        conf = _confidence(conv, clicks)
        justified = t / acos - 1.0       # effective move that lands exactly on target
        if acos >= t * LOWER_BUFFER:
            cut = _scale_step(_overtarget_cut(acos, t, spend, s.target_cpa),
                              conv, clicks)
            m = min(max(-cut, justified), 0.0)
            if m >= -1e-6:
                return finish(0.0, "HOLD", conf,
                              f"Matured ACOS {acos:.1%} over target {t:.1%} "
                              f"({acos / t:.1f}x) but already at its "
                              f"revenue-justified effective bid - hold{mat_note}.")
            floor_note = (" (held at the revenue-justified level)"
                          if m <= justified + 1e-9 else "")
            return finish(m, "LOWER BID", conf,
                          f"Matured ACOS {acos:.1%} over target {t:.1%} "
                          f"({acos / t:.1f}x) - step effective bid {m:+.0%}"
                          f"{floor_note}{mat_note}.")
        if acos <= t * RAISE_BUFFER:
            base_step = (RAISE_STEP_STRONG if acos < t * STRONG_WINNER_RATIO
                         else RAISE_STEP)
            m = min(_scale_step(base_step, conv, clicks), max(justified, 0.0))
            if m <= 1e-6:
                return finish(0.0, "HOLD", conf,
                              f"Matured ACOS {acos:.1%} under target but already at "
                              f"its revenue-justified effective bid{mat_note}.")
            return finish(m, "RAISE BID", conf,
                          f"Matured ACOS {acos:.1%} well under target {t:.1%} - "
                          f"step effective bid {m:+.0%} (revenue-justified limit "
                          f"{justified:+.0%}){mat_note}.")
        return finish(0.0, "HOLD", conf,
                      f"Matured ACOS {acos:.1%} within target tolerance "
                      f"({t * RAISE_BUFFER:.1%}-{t * LOWER_BUFFER:.1%}) - at "
                      f"target, no change{mat_note}.")

    if clicks == 0:
        if r["Impressions"] < MIN_IMPR_VISIBILITY:
            return finish(NO_TRAFFIC_RAISE, "REVIEW - NO TRAFFIC", "LOW",
                          f"Only {int(r['Impressions'])} impressions - starved. "
                          f"Suggest {NO_TRAFFIC_RAISE:+.0%} effective to test "
                          f"visibility.")
        return finish(0.0, "HOLD", "LOW",
                      f"{int(r['Impressions'])} impressions, 0 clicks (CTR issue, "
                      f"not a bid issue).")

    # clicks > 0, no orders yet
    act_at = s.target_cpa * ZERO_SALE_BUFFER
    if spend > act_at:
        exp_clicks = s.clicks_per_order
        if (spend >= PAUSE_SPEND_MULT * s.target_cpa
                and clicks >= PAUSE_CLICK_FACTOR * exp_clicks
                and _evidence_full(conv, clicks)):
            rec.update(action="PAUSE", new_bid=0.0, confidence="HIGH", eff=0.0,
                       reason=f"${spend:.2f} spent (>= {PAUSE_SPEND_MULT:.0f}x "
                              f"target CPA ${s.target_cpa:.2f}) over {int(clicks)} "
                              f"clicks (>= {PAUSE_CLICK_FACTOR:.1f}x the "
                              f"~{exp_clicks:.0f} clicks an order takes here) with "
                              f"0 orders.")
            return rec
        # anticipated-RPC level in effective terms: desired eff CPC / current
        justified = s.target_cpa / spend - 1.0 if spend > 0 else 0.0
        cut = _scale_step(_overtarget_cut(spend / act_at * t, t, spend,
                                          s.target_cpa), conv, clicks)
        m = min(max(-cut, justified), 0.0)
        if m >= -1e-6:
            return finish(0.0, "WATCH", _confidence(conv, clicks),
                          f"{int(clicks)} clicks, ${spend:.2f} spend, 0 orders - "
                          f"but already at or below the anticipated-RPC level; "
                          f"hold and re-check next run.")
        return finish(m, "LOWER BID", _confidence(conv, clicks),
                      f"${spend:.2f} spend (> {ZERO_SALE_BUFFER:.2f}x target CPA "
                      f"${s.target_cpa:.2f}) over {int(clicks)} clicks with 0 "
                      f"orders - step effective bid {m:+.0%} toward the "
                      f"anticipated-RPC level. Steps down again each run while "
                      f"clicks keep not converting.")

    return finish(0.0, "WATCH", "LOW",
                  f"{int(clicks)} clicks, ${spend:.2f} spend - still under the "
                  f"${act_at:.2f} action threshold ({ZERO_SALE_BUFFER:.2f}x target "
                  f"CPA ${s.target_cpa:.2f}). Too early to act on zero orders.")


def _cpc_cap_bid(max_adj: float, max_cpc: float) -> float:
    """Highest base bid allowed so base x (1 + adjustment) <= the ceiling (floored).

    Negative (grandfathered) adjustments are deliberately NOT credited: the base
    bid itself is never allowed above the ceiling. Conservative by design."""
    cap = max_cpc / (1.0 + max(0.0, max_adj))
    return max(MIN_BID, math.floor(cap * 100 + 1e-9) / 100)


def _apply_cpc_ceiling(rec: dict, current_bid: float, lam: float,
                       max_adj: float, max_cpc: float) -> dict:
    """Cap the base bid so base x (1 + highest campaign adjustment) never exceeds
    the effective-CPC ceiling; rec['eff'] is recomputed so the label and the
    Effective Change column stay truthful."""
    if rec["action"] in ("PAUSE", "REVIEW - NO BID IN EXPORT"):
        return rec
    capped_bid = _cpc_cap_bid(max_adj, max_cpc)
    if capped_bid < rec["new_bid"]:
        rec["new_bid"] = capped_bid
        if current_bid > 0 and lam > 0:
            rec["eff"] = capped_bid / (current_bid * lam) - 1.0
        rec["reason"] += (f" Capped to ${capped_bid:.2f} so the top placement's "
                          f"effective bid (x{1 + max_adj:.2f}) stays <= "
                          f"${max_cpc:.2f}.")
        rec["capped"] = True
    return rec


_STATUS_COLS = ("Campaign status", "Ad group status", "Status")


def _is_active(r: pd.Series) -> bool:
    """Active = every present status column is enabled/active. Columns absent
    (the owner's hand-built schema) -> treated as active."""
    for col in _STATUS_COLS:
        if col in r.index and str(r[col]).strip().lower() not in ("enabled", "active"):
            return False
    return True


def _apply_low_click_bump(rec: dict, r: pd.Series, max_adj: float, max_cpc: float,
                          lam: float = 1.0) -> dict:
    """Give an active, low-click keyword/target a small EFFECTIVE exposure bump.
    Runs AFTER the CPC ceiling and stays under it; never overrides a cut/pause and
    never shrinks a larger raise."""
    if not _is_active(r) or r["Clicks"] > LOW_CLICK_CLICKS:
        return rec
    if rec["action"] in ("LOWER BID", "PAUSE", "REVIEW - NO BID IN EXPORT"):
        return rec
    if rec.get("eff", 0.0) >= LOW_CLICK_BUMP - 1e-9:
        return rec                                    # already a bigger raise
    bumped = min(round(r["Current Bid"] * (1 + LOW_CLICK_BUMP) * lam, 2),
                 _cpc_cap_bid(max_adj, max_cpc))
    if bumped > rec["new_bid"] + 1e-9:
        rec["new_bid"] = max(MIN_BID, bumped)
        # actual delivered lift (the CPC ceiling may have trimmed the +5%)
        rec["eff"] = (bumped / (r["Current Bid"] * lam) - 1.0
                      if r["Current Bid"] > 0 and lam > 0 else LOW_CLICK_BUMP)
        rec["reason"] += (f" {LOW_CLICK_BUMP:+.0%} effective low-traffic exposure "
                          f"bump ({int(r['Clicks'])} clicks <= {LOW_CLICK_CLICKS}, "
                          f"active).")
    return rec


def run_keyword_engine(kw: pd.DataFrame, stats: dict, coeffs: dict,
                       max_cpc: float = MAX_EFFECTIVE_CPC,
                       allowance: float = 0.0) -> pd.DataFrame:
    rows = []
    for _, r in kw.iterrows():
        s = stats[r["Campaign ID"]]
        lam, max_adj = coeffs.get(r["Campaign ID"], (1.0, 0.0))
        rec = optimize_keyword(r, s, lam, allowance)
        rec = _apply_cpc_ceiling(rec, r["Current Bid"], lam, max_adj, max_cpc)
        rec = _apply_low_click_bump(rec, r, max_adj, max_cpc, lam)
        rec = _reconcile_action(rec, r["Current Bid"])
        rec["eff_change"] = 0.0 if rec["action"] == "PAUSE" else rec.get("eff", 0.0)
        rows.append(_kw_rec_row(r, rec))
    return pd.DataFrame(rows)


# ============================================================================
# Campaign summary
# ============================================================================
def build_summary(kw: pd.DataFrame, kw_rec: pd.DataFrame,
                  stats: dict) -> pd.DataFrame:
    rows = []
    for cid, s in stats.items():
        g = kw[kw["Campaign ID"] == cid]
        rec = kw_rec[kw_rec["Campaign ID"] == cid]
        wasted = g[(g["Conversions"] == 0)]["Cost"].sum()
        rows.append({
            "Campaign Name": g["Campaign Name"].iloc[0], "Campaign ID": cid,
            "Target ACOS": s.target,
            "Clicks": int(s.clicks), "Orders": int(s.conversions),
            "Spend": round(s.cost, 2), "Sales": round(s.sales, 2),
            "ACOS": round(s.acos, 4), "CVR": round(s.cvr, 4),
            "AOV": round(s.aov, 2), "Target CPA": round(s.target_cpa, 2),
            "Zero-Order Spend": round(wasted, 2),
            "Zero-Order Spend %": round(wasted / s.cost, 3) if s.cost else 0,
            "Raises": int((rec["Action"] == "RAISE BID").sum()),
            "Cuts": int((rec["Action"] == "LOWER BID").sum()),
            "Pauses": int((rec["Action"] == "PAUSE").sum()),
            # everything that is not an explicit bid move, however it is labelled
            "Holds/Watch": int((~rec["Action"].isin(
                ["RAISE BID", "LOWER BID", "PAUSE"])).sum()),
            "Notes": "; ".join(s.notes),
        })
    return pd.DataFrame(rows)


# ============================================================================
# Output workbook
# ============================================================================
HEADER_FILL = PatternFill("solid", fgColor="1F3864")
HEADER_FONT = Font(name="Arial", bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name="Arial", size=10)
ACTION_FILLS = {
    "RAISE BID": PatternFill("solid", fgColor="C6EFCE"),
    "RAISE ADJUSTMENT": PatternFill("solid", fgColor="C6EFCE"),
    "LOWER BID": PatternFill("solid", fgColor="FFEB9C"),
    "LOWER ADJUSTMENT": PatternFill("solid", fgColor="FFEB9C"),
    "PAUSE": PatternFill("solid", fgColor="FFC7CE"),
    "SUPPRESS (VIA BASE BIDS)": PatternFill("solid", fgColor="FFC7CE"),
    "SUPPRESS (BASE CUT, ADJ HELD)": PatternFill("solid", fgColor="FFE1E4"),
    "COMPENSATE ADJUSTMENT": PatternFill("solid", fgColor="C6EFCE"),
    "LOWER BASE (COMPENSATED)": PatternFill("solid", fgColor="DDEBF7"),
    "REVIEW - OFF-AMAZON": PatternFill("solid", fgColor="FCE4D6"),
    "REVIEW - NO BID IN EXPORT": PatternFill("solid", fgColor="D9D9D9"),
}
PCT_COLS = {"ACOS", "CVR", "CTR", "Target ACOS", "Zero-Order Spend %",
            "Effective Change",
            "Current Adjustment", "Recommended Adjustment"}
MONEY_COLS = {"Spend", "Sales", "CPC", "Current Bid", "Recommended Bid",
              "AOV", "Target CPA", "Zero-Order Spend"}


def style_sheet(ws, df):
    for j, col in enumerate(df.columns, start=1):
        c = ws.cell(row=1, column=j)
        c.fill, c.font = HEADER_FILL, HEADER_FONT
        c.alignment = Alignment(wrap_text=True, vertical="center")
        width = max(12, min(60, int(df[col].astype(str).str.len().max() or 10) + 2))
        if col == "Reason":
            width = 70
        ws.column_dimensions[get_column_letter(j)].width = width
    act_idx = list(df.columns).index("Action") + 1 if "Action" in df.columns else None
    for i in range(2, ws.max_row + 1):
        for j, col in enumerate(df.columns, start=1):
            cell = ws.cell(row=i, column=j)
            cell.font = BODY_FONT
            if col in PCT_COLS:
                cell.number_format = "0.0%"
            elif col in MONEY_COLS:
                cell.number_format = "$#,##0.00"
            elif col in ID_COLS:
                # text format: stops Excel showing 2.9E+14 or rounding away the
                # last digits of a 15+ digit identifier
                cell.number_format = "@"
            if col == "Reason":
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        if act_idx:
            fill = ACTION_FILLS.get(ws.cell(row=i, column=act_idx).value)
            if fill:
                ws.cell(row=i, column=act_idx).fill = fill
    ws.freeze_panes = "A2"


def settings_notes(target: float, window_days: int, max_cpc: float,
                   sales_src: str) -> pd.DataFrame:
    allowance = late_conv_allowance(window_days)
    rows = [
        ("Run assumptions", ""),
        ("Global target ACOS", f"{target:.1%} (28-30% goal -> midpoint)"),
        ("Data window", f"{window_days} days -> {allowance:.0%} of sales assumed still "
                        f"in flight (late conversions)"),
        ("Bidding strategy", "Dynamic bids - DOWN ONLY assumed for all campaigns "
                             "(per account owner). Effective max CPC = base bid x "
                             "(1 + placement adjustment)."),
        ("Sales figures", sales_src),
        ("Placement adjustments", "Amazon allows 0% to +900% only. No negative "
                                  "placement adjustments exist."),
        ("", ""),
        ("How a recommendation is reached", ""),
        ("1. Mature the data", f"Observed sales are grossed up by 1/(1-{allowance:.2f}) "
                               f"because a {window_days}-day pull has not been credited "
                               f"yet with orders still attributing to its recent "
                               f"clicks. Every threshold below reads MATURED figures. "
                               f"Amazon's reported ACOS is preferred over Cost/Sales "
                               f"(exports round Cost to whole dollars)."),
        ("2. Solve placements jointly", f"Per campaign, one base-bid factor lam and "
                                f"every adjustment are solved together: new adj = "
                                f"(1+adj)(1+pi)/lam-1 with lam = min(1, "
                                f"min (1+adj)(1+pi)). SP adjustments cannot go below "
                                f"0%, so a bad placement at the floor is suppressed "
                                f"by cutting base bids (lam < 1) while the other "
                                f"placements get compensating raises from the same "
                                f"equation - their effective bids are held exactly. "
                                f"Merit lean-in +{PLACEMENT_STEP:.0%} effective, "
                                f"suppression {SUPPRESS_STEP:.0%} effective per run, "
                                f"adjustments capped at +{PLACEMENT_MAX:.0%}."),
        ("3. Compose keyword bids", "Each keyword's own effective move m (judged "
                                    "on matured ACOS) composes with the campaign "
                                    "factor: new base = bid x (1+m) x lam. At "
                                    "compensated placements the keyword experiences "
                                    "exactly (1+m); at a suppressed placement "
                                    "(1+m)(1+pi). No compounding by accident."),
        ("4. Cap the effective CPC", f"Base bids are capped so base x (1 + the "
                                     f"campaign's highest adjustment) never exceeds "
                                     f"${max_cpc:.2f} (--max-cpc). A hard safety rail: "
                                     f"it outranks the gradual-step cap."),
        ("", ""),
        ("Key thresholds (editable constants at top of ppc_optimizer.py)", ""),
        ("Tolerance band", f"nothing is touched while matured ACOS sits within "
                           f"{RAISE_BUFFER:.0%}-{LOWER_BUFFER:.0%} of target "
                           f"({target * RAISE_BUFFER:.1%}-{target * LOWER_BUFFER:.1%}) "
                           f"- a couple of points off target is not a signal"),
        ("Raise steps", f"+{RAISE_STEP:.0%} standard, +{RAISE_STEP_STRONG:.0%} when "
                        f"matured ACOS < {STRONG_WINNER_RATIO:.0%} of target; never "
                        f"above the revenue-justified (RPC x target) bid"),
        ("Over-target cuts", f"tiered by how far over target: "
                             + ", ".join(f"<={b:.1f}x -> {c:.0%}"
                                         for b, c in OVERTARGET_CUT_TIERS)
                             + f", else {OVERTARGET_CUT_MAX:.0%}; "
                               f"+{HIGH_SPEND_BUMP:.0%} when spend >= "
                               f"{HIGH_SPEND_MULT:.0f}x Target CPA. Never below the "
                               f"revenue-justified bid, and a cut never raises a bid."),
        ("Max bid change per run", f"+/-{MAX_BID_CHANGE:.0%} (the effective-CPC "
                                   f"ceiling may exceed this - it is a safety rail)"),
        ("Evidence gating", f"a move backed by < {EVIDENCE_ORDERS} orders AND "
                            f"< {EVIDENCE_CLICKS} clicks is halved"),
        ("Zero-order keywords", f"not judged until spend > {ZERO_SALE_BUFFER:.2f}x "
                                f"Target CPA (= AOV x target ACOS); then stepped toward "
                                f"the anticipated-RPC bid (AOV/clicks x target), which "
                                f"self-steps down each run while clicks do not convert"),
        ("Pause threshold", f"spend >= {PAUSE_SPEND_MULT:.0f}x Target CPA AND clicks "
                            f">= {PAUSE_CLICK_FACTOR:.1f}x campaign clicks-per-order "
                            f"AND a full-size evidence sample, with 0 orders"),
        ("Placement judgment floor", f"{MIN_PLACEMENT_CLICKS} clicks; a 0-order "
                                     f"placement is only called dead once it has "
                                     f">= {DEAD_PLACEMENT_CLICK_MULT:.1f}x the clicks "
                                     f"an order normally takes in that campaign"),
        ("Placement suppression", f"an over-target placement whose own adjustment "
                                  f"can absorb the cut is simply lowered; one already "
                                  f"at the 0% floor is suppressed via base bids "
                                  f"({SUPPRESS_STEP:.0%} effective per run) with the "
                                  f"other placements compensated. Never raised."),
        ("Low-traffic bump", f"an active kw/target in an active ad group with "
                             f"<= {LOW_CLICK_CLICKS} clicks gets +{LOW_CLICK_BUMP:.0%} "
                             f"for exposure; never overrides a cut/pause, a larger "
                             f"raise, or the effective-CPC ceiling"),
        ("Blank bids", "auto/ASIN targets that inherit the ad group default are "
                       "flagged REVIEW - NO BID IN EXPORT, never treated as $0"),
        ("", ""),
        ("Data window advice", f"Longer windows need less maturing and give more "
                               f"confident calls: {window_days} days -> "
                               f"{allowance:.0%} allowance, 30 days -> "
                               f"{late_conv_allowance(30):.0%}. Pass --window-days to "
                               f"match your export."),
        ("Determinism", "Same inputs always produce identical output. No randomness, "
                        "no API calls."),
        ("Disclaimer", "Suggestions only. Review before applying in Amazon Ads."),
    ]
    return pd.DataFrame(rows, columns=["Setting", "Value"])


def write_output(path: str, kw_rec, pl_rec, summary, notes):
    with pd.ExcelWriter(path, engine="openpyxl") as xl:
        summary.to_excel(xl, sheet_name="Campaign Summary", index=False)
        kw_rec.to_excel(xl, sheet_name="Keyword Recommendations", index=False)
        pl_rec.to_excel(xl, sheet_name="Placement Recommendations", index=False)
        notes.to_excel(xl, sheet_name="Settings & Notes", index=False)
        wb = xl.book
        style_sheet(wb["Campaign Summary"], summary)
        style_sheet(wb["Keyword Recommendations"], kw_rec)
        style_sheet(wb["Placement Recommendations"], pl_rec)
        ws = wb["Settings & Notes"]
        ws.column_dimensions["A"].width = 45
        ws.column_dimensions["B"].width = 100
        for row in ws.iter_rows():
            for c in row:
                c.font = BODY_FONT
                c.alignment = Alignment(wrap_text=True, vertical="top")
        for c in ws[1]:
            c.fill, c.font = HEADER_FILL, HEADER_FONT


# ============================================================================
# Main
# ============================================================================
def main():
    p = argparse.ArgumentParser(
        description="Amazon PPC bid & placement optimizer (suggestions only)",
        epilog=f"Standard workflow: rename the two Amazon exports to "
               f"{INPUT_DIR}/{DEFAULT_KEYWORDS} and {INPUT_DIR}/{DEFAULT_PLACEMENTS}, "
               f"then run with no arguments. Results go to {DEFAULT_OUTPUT}.",
        formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("keywords", nargs="?", default=None,
                   help=f"Keyword performance .xlsx (or combined workbook). "
                        f"Omit to use {INPUT_DIR}/{DEFAULT_KEYWORDS}")
    p.add_argument("placements", nargs="?", default=None,
                   help=f"Placement performance .xlsx. Omit to use "
                        f"{INPUT_DIR}/{DEFAULT_PLACEMENTS}")
    p.add_argument("-o", "--output", default=None,
                   help=f"Output workbook path (default {DEFAULT_OUTPUT})")
    p.add_argument("-t", "--target", type=float, default=TARGET_ACOS_DEFAULT,
                   help="Global target ACOS as fraction (default 0.29)")
    p.add_argument("--targets-file", default=None,
                   help="CSV with 'Campaign ID','Target ACOS' per-campaign overrides")
    p.add_argument("--max-cpc", type=float, default=MAX_EFFECTIVE_CPC,
                   help=f"Hard effective-CPC ceiling in dollars: base bid x "
                        f"(1+adjustment) never exceeds this at any placement. "
                        f"Default {MAX_EFFECTIVE_CPC}")
    p.add_argument("--window-days", type=int, default=WINDOW_DAYS_DEFAULT,
                   help=f"How many days the export covers. Shorter "
                        f"windows get a larger late-conversion allowance, so recent "
                        f"clicks are not judged on sales that have not landed yet. "
                        f"Default {WINDOW_DAYS_DEFAULT}")
    args = p.parse_args()

    if not 0.01 <= args.target <= 1.0:
        fail("--target must be a fraction between 0.01 and 1.0 (e.g. 0.29)")
    if args.window_days < 1:
        fail("--window-days must be at least 1")
    if args.max_cpc <= MIN_BID:
        fail(f"--max-cpc must be greater than the Amazon minimum bid ${MIN_BID:.2f}")

    kw_path, pl_path = resolve_default_inputs(args.keywords, args.placements)
    out_path = args.output or _here(DEFAULT_OUTPUT)
    parent = os.path.dirname(os.path.abspath(out_path))
    if parent:
        os.makedirs(parent, exist_ok=True)

    kw_raw, pl_raw = load_inputs(kw_path, pl_path)
    kw, pl = normalize(kw_raw, True), normalize(pl_raw, False)

    targets = {}
    if args.targets_file:
        tf = pd.read_csv(args.targets_file, dtype={"Campaign ID": str})
        for c in ("Campaign ID", "Target ACOS"):
            if c not in tf.columns:
                fail(f"--targets-file must contain columns 'Campaign ID' and "
                     f"'Target ACOS'; found {list(tf.columns)}")
        t = pd.to_numeric(tf["Target ACOS"], errors="coerce")
        if t.max() > 1.5:
            t = t / 100.0
        targets = dict(zip(tf["Campaign ID"].astype(str), t))

    account_cvr = kw["Conversions"].sum() / kw["Clicks"].sum() if kw["Clicks"].sum() else 0.05
    conv_rows = kw[kw["Conversions"] > 0]
    account_aov = (conv_rows["Sales"].sum() / conv_rows["Conversions"].sum()
                   if conv_rows["Conversions"].sum() else 0.0)
    if account_aov == 0:
        fail("No conversions anywhere in the keyword file - cannot derive AOV. "
             "Add a Sales column or provide data containing at least one order.")

    stats = campaign_stats(kw, targets, args.target, account_cvr, account_aov)

    pl_main, pl_off, pl_ignored = canonicalize_placements(pl)

    # Placements are solved first; each campaign's base factor lam and its
    # compensating adjustments come out of one equation system, and keyword
    # bids then compose with lam exactly (base = bid x (1+m) x lam).
    allowance = late_conv_allowance(args.window_days)
    pl_rec, coeffs = solve_placements(pl_main, stats, allowance)
    kw_rec = run_keyword_engine(kw, stats, coeffs, args.max_cpc, allowance)

    off_rec = flag_off_amazon(pl_off, stats)
    if not off_rec.empty:
        pl_rec = pd.concat([pl_rec, off_rec], ignore_index=True)

    summary = build_summary(kw, kw_rec, stats)
    sales_src = kw["_sales_src"].iloc[0] if len(kw) else "n/a"
    write_output(out_path, kw_rec, pl_rec, summary,
                 settings_notes(args.target, args.window_days, args.max_cpc,
                                sales_src))

    print(f"Done. {len(kw_rec)} keywords and {len(pl_rec)} placement rows analyzed "
          f"across {len(stats)} campaign(s).")
    print(f"Actions: {kw_rec['Action'].value_counts().to_dict()}")
    off_flagged = len(off_rec) if not off_rec.empty else 0
    print(f"Placements: {len(pl_off)} Off-Amazon rows ({off_flagged} flagged, "
          f"rest ignored); {len(pl_ignored)} 'Other Placements' rows ignored.")
    print(f"Output: {out_path}")


if __name__ == "__main__":
    main()
