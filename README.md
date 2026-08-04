# Amazon PPC Bid & Placement Optimizer

Deterministic, rules-based optimizer for Amazon Sponsored Products.
Reads keyword + placement performance exports (Excel), writes a styled
recommendations workbook. Suggestions only — never touches live campaigns.

## Quick start
```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

**Every run:** rename your two Amazon exports and drop them in `input/`, then run
with no arguments:

| Export | Rename to |
|---|---|
| `kt_export_*.xlsx` (keywords/targets) | `input/keywords.xlsx` |
| `pl_export_*.xlsx` (placements) | `input/placements.xlsx` |

```bash
python ppc_optimizer.py --window-days 7
```

Results land in `outputs/recs.xlsx`. The defaults resolve relative to the script,
so this works from any working directory. Overwrite the two input files after each
new export — nothing else changes. Close `outputs/recs.xlsx` in Excel first, or the
write fails with *Permission denied*.

To use different files just once, pass them explicitly:
```bash
python ppc_optimizer.py SOME_KW.xlsx SOME_PL.xlsx -o outputs/other.xlsx
```

Native Amazon report exports work directly (`Campaign name`, `Orders`, `Bid`,
`Bid adj. %`, `Detail Page on-Amazon`, …) alongside the hand-built schema in
`data/`. A single combined workbook also works, but only if its tabs are named
exactly `Keywords` and `Placements` — the native keyword export's tab is
`Keywords & Targets`, so two files is the simpler path. Run `pytest -q` for the
regression suite.

## Options
| Flag | Meaning | Default |
|---|---|---|
| `-o` | output path | `outputs/recs.xlsx` |
| `-t` | global target ACOS (fraction) | `0.29` |
| `--targets-file` | CSV `Campaign ID,Target ACOS` per-campaign overrides | — |
| `--max-cpc` | hard effective-CPC ceiling `base × (1+adj)` | `1.85` |
| `--window-days` | days the export covers (drives the late-conversion allowance) | `7` |

Enter `Recommended Bid` (the base bid) and `Recommended Adjustment` (per
placement) into Amazon.

`--window-days` must match the number of days your export actually covers — it
drives the late-conversion allowance, so running a 30-day export as `7` would
over-cut everything.

## How it works

Bids and placement adjustments are solved as **one coupled system**, because the
effective CPC at a placement is `base bid × (1 + that placement's adjustment)`.
Per campaign, one base-bid factor **λ** and *all* the placement adjustments come
out of a single set of equations:

```
new adj_p = (1 + adj_p)(1 + π_p) / λ − 1        (π_p = placement p's merited move)
λ         = min(1, min_p (1 + adj_p)(1 + π_p))
new base  = bid × (1 + m) × λ                    (m = the keyword's merited move)
```

Placement adjustments are **non-negative (0%–+900%)** — a placement can be bid
*up* from the base bid but never below it. A too-high adjustment is simply lowered
when that alone fixes the placement; but one already at the 0% floor can be pushed
below the campaign's base level only one way: **cut the base bids (λ < 1) and
compensate the other placements upward** — and both halves
fall out of the same equation, so the surviving placements' effective bids are held
*exactly* while the bad one absorbs the cut. When nothing needs suppressing, λ = 1
and healthy settings don't churn. Because λ and the compensations are algebraically
linked, a keyword raise and a placement raise can never silently compound: the
keyword experiences exactly `(1 + m)` at compensated placements and
`(1 + m)(1 + π_p)` at the suppressed one. The `Effective Change` column on both
sheets reports the delivered effective move; `Action` is labelled by that
effective direction. A keyword whose base falls only because its campaign is
suppressing (exposure held by the compensations) is labelled
`LOWER BASE (COMPENSATED)`, not `LOWER BID`.

Everything moves in **small, self-stepping increments** and converges over a few
export cycles, so every run is reviewable.

- **Late-conversion maturity** (`--window-days`, default 7): a short window has
  not been credited yet with orders still attributing to its recent clicks, so
  observed sales are understated and observed ACOS overstated. Sales are grossed
  up once (~15% on 7 days) and *everything* — tolerance bands, cut tiers, RPC
  ceilings — reads the matured figures. Practical effect: 30–32% observed ACOS on
  a 7-day pull is treated as on-target, while 50%/60% get the 15%/20% cuts you'd
  apply by hand.
- **Tolerance band**: nothing is touched while matured ACOS sits within 80–120%
  of target. Being a couple of points off target is not a signal.
- **Gradual cuts**: a converter past the band steps down by ACOS band
  (~10/15/20%, +5% on heavy spend), never below its revenue-justified bid, and a
  "cut" can never raise a bid. Single-run change capped at 30%.
- **Evidence gating**: any move backed by fewer than 3 orders *and* fewer than 25
  clicks is halved. Zero-order keywords aren't judged until spend clears 1.25×
  Target CPA; a 0-order placement isn't called dead until it has ~1.5× the clicks
  an order normally takes there.
- **Placements**: merit lean-in is +20% effective per run; suppression is 15%
  effective per run (λ ≥ 0.85), both capped at +100% adjustment. An over-target
  placement whose own adjustment can absorb the cut is simply lowered (no base
  cut); one already at the 0% floor is suppressed via base bids with the others
  compensated. An over-target or dead placement is never raised, and a
  non-converting placement isn't judged "dead" until it has ~1.5× the clicks an
  order normally takes there.
- **Effective-CPC ceiling** (`--max-cpc`, default $1.85): base bids are capped so
  `base × (1 + highest campaign adjustment) ≤ ceiling`. A hard safety rail — it
  is the one rule allowed to exceed the 30% gradual-step cap.
- **Low-traffic exposure bump**: an *active* keyword/target in an *active* ad
  group (campaign, ad group, and keyword all `enabled`) with ≤10 clicks gets +5%
  to buy more data. A floor only — never overrides a cut/pause, never shrinks a
  larger raise, never breaches the ceiling.
- **Off-Amazon** placements are ignored unless ACOS > 40% or they spend with zero
  sales, in which case they're flagged for review (no bid change).
  `Other Placements` rows are dropped; `Top of Search` and `Top of Search
  on-Amazon` are treated as the same placement.
- **Blank bids** (auto/ASIN targets inheriting the ad-group default) are never
  treated as $0 — they're flagged `REVIEW - NO BID IN EXPORT` with the
  revenue-justified figure, since no percentage step is meaningful without a base.
- **The bid cap must actually be binding.** Under down-only the bid is a true
  ceiling, so realized CPC can never exceed it. When CPC sits below **70% of the
  base bid** (with ≥10 clicks, since `Cost` rounding makes CPC unreliable below
  that), the auctions are clearing well under the cap and the bid is not the
  lever: a raise could only win the thin, empty slice just above the old cap, and
  a cut merely shaves a ceiling nobody reaches.
  - Raises are held as **`HOLD - BID NOT THE CONSTRAINT`**, pointing at
    impressions/relevance instead.
  - Cuts still proceed (they do trim volume at the margin) but the reason says
    the bid must reach ~CPC before what you *pay* changes.
  - Either way the row is flagged **`Bid Cap Unused = YES`**, a separate column
    from `Action` so the diagnostic survives even when a suppressing campaign
    relabels the row.
  - Measured against the **base bid**, not `base × (1+adj)`: the adjustment
    uplift is already inside realized CPC, and using the campaign's *max*
    adjustment invents headroom for keywords that never serve on the boosted
    placement (it wrongly froze winners paying 100% of their base bid).
- **Sponsored Products only.** Exports often mix campaign types; Sponsored Brands
  campaigns are detected via `Campaign type` and skipped, with a count printed at
  run time. SB has a different placement taxonomy (`Other Placements`) and permits
  negative adjustments, so every rule here would be wrong for it.

### Data limits worth knowing
- Amazon exports carry no keyword×placement performance, so each keyword's
  placement mix is approximated by its **campaign's** click mix. Exact at
  campaign level, approximate per keyword.
- Amazon's reported ACOS is preferred over `Cost ÷ Sales`, because exports round
  Cost to whole dollars (up to ~4 points of error on low-spend rows).
- Sponsored Products placement adjustments are **non-negative (0%–+900%)**:
  they bid a placement up from base, never below it, though the value itself can
  be raised *or lowered* freely within that range. The −20% settings that appear
  in these exports belong to **Sponsored Brands** rows only — the engine holds
  those as set (never raises them) and never emits a negative recommendation.
- Longer windows need less maturing and yield more confident calls. 7-day pulls
  produce many LOW-confidence rows by design.

## Notes
- All thresholds are labeled constants at the top of `ppc_optimizer.py`.
- Deterministic: same inputs always produce identical output.
- See `CLAUDE.md` for project context and owner operating constraints.
