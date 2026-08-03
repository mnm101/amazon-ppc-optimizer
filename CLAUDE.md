# PPC Optimizer — project brief for Claude Code

## What this is
A deterministic, rules-based Amazon Sponsored Products bid & placement
optimizer. Reads keyword + placement performance Excel exports (one or many
campaigns), writes a styled recommendations workbook. **Suggestions only** —
it never calls the Amazon Ads API or changes live campaigns (API integration
is a planned future phase, not yet approved by the owner).

## Owner's operating context (do not change without asking)
- Target ACOS: **28–30%**, engine midpoint **0.29** (global default; per-campaign
  override via `--targets-file` CSV).
- All campaigns run **Dynamic bids – down only**. Effective max CPC =
  base bid × (1 + placement adjustment). No up-and-down doubling.
- Sponsored Products placement adjustments are **non-negative: 0%–+900%**
  (owner re-confirmed 2026-07-30). They bid a placement UP from the base bid and
  never below it — but the value is freely raisable *and lowerable* in range, so
  reducing a too-high adjustment is normal and expected. Avoid the term
  "increase-only": it wrongly implies an existing adjustment can't be reduced.
  The −20% values in the placement export are **Sponsored Brands** rows — do NOT
  generalize from them; the engine holds them as set and never emits a negative
  recommendation.
- Input schema matches the owner's hand-built exports (see `data/` samples).
  Sales column absent → derived as Cost ÷ ACOS.

## Methodology
Bids and placement adjustments are solved as **one coupled system**: the
effective CPC at a placement is `base bid × (1 + that placement's adjustment)`.
Per campaign, a single base-bid factor λ and every adjustment come out of one
equation set:
`new adj_p = (1+adj_p)(1+π_p)/λ − 1`, `λ = min(1, min_p (1+adj_p)(1+π_p))`,
`new base = bid × (1+m) × λ`. Since adjustments cannot go below 0%,
**λ < 1 IS suppression via base bids**, and the other placements' compensating
raises come from the same equation — their effective bids are held exactly while
the bad placement absorbs the cut. λ = 1 when nothing needs suppressing, so
healthy campaigns don't churn. Keyword and placement moves compose as
`(1+m)(1+π_p)` — they cannot compound by accident. This is the owner's Mode-B
strategy `(1+cur)/(1−s)−1` derived from first principles and generalized: any
placement can trigger it, and it is sized to need (≤15% effective per run).

`Action` is labelled by the *effective* direction (the `Effective Change` column
on both sheets); a keyword whose base falls only because its campaign suppresses
is `LOWER BASE (COMPENSATED)`, not `LOWER BID`.

Known approximation: each placement's merit (π_p) is judged at current bids, and
each keyword's placement mix is approximated by its campaign's mix (Amazon
exports carry no keyword×placement data). The per-run steps converge over
cycles.

Four principles govern every rule (owner-directed, 2026-07-30):
1. **Mature before judging.** A short window hasn't been credited with orders
   still attributing to recent clicks, so sales are grossed up once
   (`--window-days`, ~15% on 7 days) and every threshold reads matured figures.
   Consequence: 30–32% observed ACOS on a 7-day pull is *on target*.
2. **Act on evidence, not noise.** Move size scales with sample size (<3 orders
   and <25 clicks → half step). Thin data → HOLD/WATCH, never a cut.
3. **Never act at a boundary.** Nothing is touched inside 80–120% of target.
4. **Small, self-stepping moves** that converge over runs, so each run is
   reviewable.

Concrete rules:
- Over-target converters: gradual step-down tiered by how far over target
  (~10/15/20%, +5% on heavy spend), never below the revenue-justified bid
  (`RPC × target`), capped at 30% per run. A cut may never raise a bid.
- Winners: +10% (or +20% under half target), never above the RPC ceiling.
- Zero-sale keywords: act only once `spend > 1.25 × Target CPA`; then step
  toward `(AOV / clicks) × target` (anticipated RPC — self-stepping).
- Pause: spend ≥ 2× Target CPA AND clicks ≥ 1.5× clicks-per-order AND a
  full-size evidence sample, with zero orders.
- Low-traffic bump: active kw/target in an active ad group with ≤10 clicks gets
  +5% for exposure. A floor only — never overrides a cut/pause or a larger raise.
- Hard effective-CPC ceiling **$1.85** (`--max-cpc`): the one rule allowed to
  exceed the 30% step cap. Owner constraint: never pay more than this anywhere.
- CVR/AOV fallback hierarchy: campaign → account.
- Determinism is a design requirement: same inputs must always produce
  identical output.

## Placements
Single joint engine (the λ solve above), adjustments in **0% … +100%**
(pre-existing settings above the cap step down toward it; grandfathered SB
negatives are held, never raised).

**History note (2026-07-30):** a session briefly switched to suppressing via
negative adjustments after seeing −20% values in the export. Those rows are
Sponsored Brands; SP adjustments cannot be negative, and Amazon would reject
negative SP recommendations. The owner corrected this — suppression must run through base
bids + compensation, which is what λ implements.
- Converting well under target → +20% effective lean-in.
- Over target with a high adjustment → its own adjustment absorbs the cut.
- Over target or dead at the 0% floor → SUPPRESS (VIA BASE BIDS): λ < 1 cuts the
  campaign's base bids, others get COMPENSATE ADJUSTMENT raises.
- An over-target or non-converting placement is never raised.
- A 0-order placement is only "dead" once it has ≥1.5× the clicks an order
  normally takes in that campaign — 0 orders in 11 clicks is noise.
- `Off Amazon` is ignored unless ACOS > 40% or it spends with zero sales (then
  flagged for review, no bid change). `Other Placements` is dropped.
  `Top of Search` and `Top of Search on-Amazon` are the same placement.

Modes A and B were removed on 2026-07-30 (owner decision) — the coupled engine
superseded them, and leaving Mode A as the default was a trap: it had no CPC
ceiling (effective bids to $4.42), cut at the exact 29% boundary, and mislabeled
45 raises as cuts.

## Layout
- `ppc_optimizer.py` — single-file program; all thresholds are labeled
  constants at the top.
- `input/` — the owner's working folder. Exports are renamed by hand to
  `keywords.xlsx` + `placements.xlsx` each run (the export filenames are
  timestamped, so fixed names are the convention). Gitignored.
- `data/` — sample inputs (real 7-day Mothers Day campaign export).
- `tests/` — pytest regression suite pinning validated numbers. Run
  `pytest -q` after ANY change to the engine; the expected values were
  hand-verified against the sample campaign.
- `outputs/` — gitignored scratch for generated workbooks; default output is
  `outputs/recs.xlsx`.

## Run
The owner's normal flow — exports renamed into `input/`, no arguments needed
(`--window-days` must match the export's window):
```
python ppc_optimizer.py --window-days 7
```
Defaults are resolved relative to the script, so the cwd does not matter.
Explicit paths still work, and the sample data is run the same way:
```
python ppc_optimizer.py data/sample_keywords.xlsx data/sample_placements.xlsx \
    -o outputs/sample.xlsx
pytest -q
```

## Known backlog (owner-approved direction, not yet built)
1. Per-keyword effective ToS/RoS/PP bid columns in the output, so the coupled
   result is visible per row rather than inferred from base × (1+adj).
2. Amazon Ads API integration (reports in, recommendations out) — discuss
   with owner before starting; the engines return plain DataFrames so the
   seam is `load_inputs()`.
3. Bulk-operations-file output format as an alternative to the review sheet.

## Input schema
Accepts native Amazon report exports *and* the hand-built schema in `data/`.
`harmonize_columns()` maps aliases (`Campaign name`→`Campaign Name`,
`Orders`→`Conversions`, `Bid`→`Current Bid`, `Bid adj. %`→`Current Adjustment`,
`Ad Sales`→`Sales`); `Target` is built from `Keyword` falling back to
`Targeting value` for auto/ASIN targets. Two data quirks are handled
deliberately: `Cost` is rounded to whole dollars in exports (so Amazon's
reported ACOS is preferred over recomputing it), and a blank `Bid` means the
target inherits the ad-group default (flagged, never treated as $0).

## Guardrails
- Never turn this into an auto-executor without explicit owner request.
- Don't loosen thin-data thresholds silently; they're the safety layer.
- Keep the tool deterministic (no randomness, no API calls in the engine).
