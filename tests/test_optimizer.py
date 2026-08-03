"""Regression tests for ppc_optimizer.

Pins the behaviour of the joint bid+placement engine: the per-campaign lambda
solve (suppression via base bids + compensating adjustments, floor 0%),
late-conversion maturity, tolerance bands, evidence-scaled steps, the
effective-CPC ceiling, and the input-schema plumbing.
"""
import subprocess
import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
import ppc_optimizer as eng                                        # noqa: E402

KW = ROOT / "data" / "sample_keywords.xlsx"
PL = ROOT / "data" / "sample_placements.xlsx"


def run(out, *extra):
    cmd = [sys.executable, str(ROOT / "ppc_optimizer.py"),
           str(KW), str(PL), "-o", str(out), *extra]
    r = subprocess.run(cmd, capture_output=True, text=True)
    assert r.returncode == 0, r.stderr
    return out


@pytest.fixture(scope="module")
def recs(tmp_path_factory):
    out = tmp_path_factory.mktemp("run") / "recs.xlsx"
    run(out)
    return {s: pd.read_excel(out, s) for s in
            ("Keyword Recommendations", "Placement Recommendations",
             "Campaign Summary", "Settings & Notes")}


def kw_row(df, name, match="BROAD"):
    m = df[(df["Keyword"] == name) & (df["Match Type"] == match)]
    assert len(m) == 1
    return m.iloc[0]


# ---------------- infrastructure ----------------
def test_rerun_is_deterministic(tmp_path):
    a = run(tmp_path / "a.xlsx")
    b = run(tmp_path / "b.xlsx")
    pd.testing.assert_frame_equal(pd.read_excel(a, "Keyword Recommendations"),
                                  pd.read_excel(b, "Keyword Recommendations"))


def test_missing_column_fails_loudly(tmp_path):
    broken = tmp_path / "broken.xlsx"
    pd.read_excel(KW).drop(columns=["Current Bid"]).to_excel(broken, index=False)
    r = subprocess.run([sys.executable, str(ROOT / "ppc_optimizer.py"),
                        str(broken), str(PL)], capture_output=True, text=True)
    assert r.returncode == 1
    assert "Current Bid" in r.stderr


def test_percent_format_acos_normalized(tmp_path):
    """ACOS given as 15.67 must be read the same as 0.1567 - percent-format input
    should produce byte-identical recommendations to fraction-format input."""
    kw = pd.read_excel(KW); kw["ACOS"] *= 100
    pl = pd.read_excel(PL); pl["ACOS"] *= 100; pl["Current Adjustment"] *= 100
    kp, pp = tmp_path / "k.xlsx", tmp_path / "p.xlsx"
    kw.to_excel(kp, index=False); pl.to_excel(pp, index=False)
    pct = tmp_path / "pct.xlsx"
    subprocess.run([sys.executable, str(ROOT / "ppc_optimizer.py"),
                    str(kp), str(pp), "-o", str(pct)], check=True)
    frac = run(tmp_path / "frac.xlsx")
    pd.testing.assert_frame_equal(pd.read_excel(pct, "Keyword Recommendations"),
                                  pd.read_excel(frac, "Keyword Recommendations"))
    pd.testing.assert_frame_equal(pd.read_excel(pct, "Placement Recommendations"),
                                  pd.read_excel(frac, "Placement Recommendations"))


# ---------------- sample-data pins ----------------
def test_sample_adjustments_within_sp_bounds(recs):
    """SP placement adjustments are never negative: 0%..+100% in this engine.
    (The value is freely lowerable within range - see the LOWER ADJUSTMENT tests.)"""
    pl = recs["Placement Recommendations"]
    assert (pl["Recommended Adjustment"] >= -1e-9).all()
    assert (pl["Recommended Adjustment"] <= eng.PLACEMENT_MAX + 1e-9).all()


def test_sample_tos_effective_lean_in(recs):
    # strong converter at +10%: (1+0.10)(1+0.20)-1 = 0.32, effective +20%
    pl = recs["Placement Recommendations"]
    tos = pl[pl["Placement"] == "Top of Search"].iloc[0]
    assert tos["Recommended Adjustment"] == pytest.approx(0.32, abs=0.01)
    assert tos["Effective Change"] == pytest.approx(0.20, abs=0.01)


def test_effective_change_column_signs(recs):
    k = recs["Keyword Recommendations"]
    assert "Effective Change" in k.columns
    assert (k[k["Action"] == "RAISE BID"]["Effective Change"] > 0).all()
    assert (k[k["Action"] == "LOWER BID"]["Effective Change"] < 0).all()


# ---------------- placement solver: helpers ----------------
def _pl(placement, clicks, orders, cost, sales, adj=0.0):
    acos = cost / sales if sales else 0.0
    return {"Campaign Name": "C", "Campaign ID": 1, "Placement": placement,
            "Impressions": clicks * 20, "Clicks": clicks, "Conversions": orders,
            "Cost": cost, "Sales": sales, "ACOS": acos,
            "CPC": cost / clicks if clicks else 0.0, "Current Adjustment": adj}


def _stats():
    s = eng.CampaignStats(target=0.29)
    s.aov = 34.0
    s.target_cpa = 10.0
    s.clicks_per_order = 10.0      # -> dead floor = 1.5 x 10 = 15 clicks
    return s


# ---------------- placement solver: the lambda joint solve ----------------
def test_thin_dead_placement_does_not_suppress():
    g = pd.DataFrame([
        _pl("top_of_search", clicks=40, orders=5, cost=25, sales=170),  # strong
        _pl("product_page", clicks=11, orders=0, cost=5, sales=0),      # dead but THIN
    ])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    acts = {r["Placement"]: r["Action"] for r in rows}
    assert lam == pytest.approx(1.0, abs=1e-9)               # no base cut on noise
    assert "SUPPRESS (VIA BASE BIDS)" not in acts.values()
    tos = next(r for r in rows if r["Placement"] == "Top of Search")
    assert tos["Recommended Adjustment"] == pytest.approx(0.20, abs=0.01)


def test_dead_placement_suppressed_via_base_bids_with_compensation():
    """The owner's mechanism: the bad placement sits at the 0% floor, so base bids
    are cut (lam < 1) and the other placements get compensating raises from the
    same equation."""
    g = pd.DataFrame([
        _pl("top_of_search", clicks=40, orders=5, cost=25, sales=170),  # strong
        _pl("product_page", clicks=60, orders=0, cost=60, sales=0),     # dead + heavy
    ])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    by = {r["Placement"]: r for r in rows}
    assert lam == pytest.approx(0.85, abs=0.001)             # 15% base cut
    assert by["Product Pages"]["Action"] == "SUPPRESS (VIA BASE BIDS)"
    assert by["Product Pages"]["Recommended Adjustment"] == 0.0
    assert by["Product Pages"]["Effective Change"] == pytest.approx(-0.15, abs=0.01)
    # ToS: merit +20% AND compensation: (1)(1.2)/0.85 - 1 = 0.41
    assert by["Top of Search"]["Action"] == "RAISE ADJUSTMENT"
    assert by["Top of Search"]["Recommended Adjustment"] == pytest.approx(0.41, abs=0.01)
    assert by["Top of Search"]["Effective Change"] == pytest.approx(0.20, abs=0.01)


def test_compensation_holds_effective_bids_exactly():
    # a thin neighbour is compensated so its effective bid does not move at all
    g = pd.DataFrame([
        _pl("rest_of_search", clicks=60, orders=0, cost=60, sales=0),  # dead + heavy
        _pl("top_of_search", clicks=5, orders=0, cost=1, sales=0),     # thin
    ])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    by = {r["Placement"]: r for r in rows}
    assert lam == pytest.approx(0.85, abs=0.001)
    assert by["Rest of Search"]["Action"] == "SUPPRESS (VIA BASE BIDS)"
    tos = by["Top of Search"]
    assert tos["Action"] == "COMPENSATE ADJUSTMENT"
    assert (1 + tos["Recommended Adjustment"]) * lam == pytest.approx(1.0, abs=0.01)
    assert abs(tos["Effective Change"]) <= 0.01


def test_high_adjustment_absorbs_the_cut_without_base_suppression():
    """'If its existing placement adjustment is high enough that lowering fixes
    the issue, then ok' - no base-bid cut is triggered."""
    g = pd.DataFrame([_pl("rest_of_search", clicks=80, orders=6, cost=60,
                          sales=60 / 0.58, adj=0.50)])       # ACOS 2x target
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    assert lam == pytest.approx(1.0, abs=1e-9)               # no suppression needed
    assert rows[0]["Action"] == "LOWER ADJUSTMENT"
    assert rows[0]["Recommended Adjustment"] == pytest.approx(0.275, abs=0.01)
    assert rows[0]["Effective Change"] == pytest.approx(-0.15, abs=0.01)


def test_at_floor_bad_placement_triggers_base_suppression():
    # same placement but already at 0%: only base bids can push it down
    g = pd.DataFrame([_pl("rest_of_search", clicks=80, orders=6, cost=60,
                          sales=60 / 0.58, adj=0.0)])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    assert lam == pytest.approx(0.85, abs=0.001)
    assert rows[0]["Action"] == "SUPPRESS (VIA BASE BIDS)"
    assert rows[0]["Recommended Adjustment"] == 0.0


def test_suppression_is_gradual_per_run():
    # 15% effective per run, not one violent jump; repeats while it stays bad
    for _ in range(2):
        g = pd.DataFrame([_pl("product_page", clicks=60, orders=0, cost=60,
                              sales=0, adj=0.0)])
        rows, lam, _ = eng.solve_campaign_placements(g, _stats())
        assert lam == pytest.approx(1 - eng.SUPPRESS_STEP, abs=0.001)
        assert rows[0]["Recommended Adjustment"] == 0.0


def test_grandfathered_negative_setting_is_held_never_raised():
    # Sponsored Brands rows can carry -20%; the engine holds them as set and the
    # base cut still delivers the reduction
    g = pd.DataFrame([
        _pl("top_of_search", clicks=40, orders=5, cost=25, sales=170),
        _pl("product_page", clicks=60, orders=0, cost=60, sales=0, adj=-0.20),
    ])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    pp = next(r for r in rows if r["Placement"] == "Product Pages")
    assert pp["Recommended Adjustment"] == pytest.approx(-0.20, abs=1e-9)
    assert lam == pytest.approx(0.85, abs=0.001)


def test_placement_near_target_is_not_touched():
    g = pd.DataFrame([_pl("top_of_search", clicks=200, orders=20,
                          cost=60, sales=60 / 0.298, adj=0.30)])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    assert lam == pytest.approx(1.0, abs=1e-9)
    assert rows[0]["Action"] == "HOLD"
    assert rows[0]["Recommended Adjustment"] == pytest.approx(0.30, abs=0.001)


def test_zero_order_placement_below_expected_clicks_is_noise():
    # 12 clicks / 0 orders where an order takes ~10 clicks: floor is 15 clicks
    g = pd.DataFrame([_pl("product_page", clicks=12, orders=0, cost=6, sales=0,
                          adj=0.20)])
    rows, _, _ = eng.solve_campaign_placements(g, _stats())
    assert rows[0]["Action"] == "HOLD"
    assert rows[0]["Recommended Adjustment"] == pytest.approx(0.20, abs=0.001)
    assert "noise" in rows[0]["Reason"]


# ---------------- keyword engine ----------------
def _kseries(bid, clicks, orders, cost, sales, impr=1000):
    acos = cost / sales if sales else 0.0
    return pd.Series({"Current Bid": bid, "Clicks": clicks, "Conversions": orders,
                      "Cost": cost, "Sales": sales, "ACOS": acos,
                      "Impressions": impr, "CVR": orders / clicks if clicks else 0})


def test_overtarget_cut_is_gentle_not_35pct():
    # 51% ACOS, 2 orders / 16 clicks = thin -> half of the 15% tier = ~7.5%
    r = _kseries(bid=1.05, clicks=16, orders=2, cost=8.0, sales=8.0 / 0.51)
    rec = eng.optimize_keyword(r, _stats())
    assert rec["action"] == "LOWER BID"
    assert 0.90 < rec["new_bid"] < 1.00


def test_near_target_is_held_not_cut():
    r = _kseries(bid=0.60, clicks=480, orders=80, cost=267.0, sales=267.0 / 0.298)
    rec = eng.optimize_keyword(r, _stats())
    assert rec["action"] == "HOLD"
    assert rec["new_bid"] == pytest.approx(0.60, abs=0.001)


def test_overtarget_high_spend_cuts_harder():
    # 51% ACOS, full evidence, heavy spend -> 15% + 5% = 20%
    r = _kseries(bid=1.05, clicks=40, orders=5, cost=25.0, sales=25.0 / 0.51)
    rec = eng.optimize_keyword(r, _stats())
    assert rec["new_bid"] == pytest.approx(1.05 * 0.80, abs=0.02)


# ---- late-conversion maturity ----
def test_late_conv_allowance_shrinks_with_window():
    assert eng.late_conv_allowance(7) == pytest.approx(0.15, abs=0.001)
    assert eng.late_conv_allowance(30) == pytest.approx(0.035, abs=0.001)


@pytest.mark.parametrize("observed,expect_cut", [
    (0.32, False),   # 32% observed -> ~27% matured: at target, leave alone
    (0.34, False),   # 34% -> ~29%: still at target
    (0.50, True),    # 50% -> ~43%: genuinely over, cut
])
def test_maturity_spares_marginally_over_target_keywords(observed, expect_cut):
    a = eng.late_conv_allowance(7)
    r = _kseries(bid=1.00, clicks=100, orders=10, cost=30.0, sales=30.0 / observed)
    rec = eng.optimize_keyword(r, _stats(), 1.0, a)
    assert (rec["action"] == "LOWER BID") is expect_cut


def test_maturity_matches_owner_practice_on_50_and_60_pct():
    # owner's rule of thumb: 50-60% ACOS -> lower 15-20% depending on spend
    a = eng.late_conv_allowance(7)
    for observed, expected in ((0.50, 0.15), (0.60, 0.20)):
        r = _kseries(bid=1.00, clicks=100, orders=10, cost=30.0,
                     sales=30.0 / observed)
        rec = eng.optimize_keyword(r, _stats(), 1.0, a)
        assert rec["new_bid"] == pytest.approx(1.00 * (1 - expected), abs=0.01)


# ---- bid/placement coupling: composition with lambda ----
def test_winner_move_composes_exactly_with_base_suppression():
    """A winner due +10% effective in a campaign suppressing at lam=0.85: base
    becomes bid x 1.10 x 0.85 while compensating adjustments carry 1/0.85, so the
    keyword's effective move at surviving placements is exactly +10%."""
    r = _kseries(bid=0.50, clicks=100, orders=12, cost=45.0, sales=45.0 / 0.22)
    rec = eng.optimize_keyword(r, _stats(), lam=0.85)
    assert rec["eff"] == pytest.approx(0.10, abs=0.001)
    assert rec["new_bid"] == pytest.approx(0.50 * 1.10 * 0.85, abs=0.01)
    assert rec["action"] == "RAISE BID"        # labelled by the effective move


def test_overtarget_effective_bid_never_rises_at_any_lambda():
    for lam in (1.0, 0.9, 0.85):
        r = _kseries(bid=1.00, clicks=100, orders=10, cost=30.0, sales=30.0 / 0.55)
        rec = eng.optimize_keyword(r, _stats(), lam=lam)
        assert rec["eff"] <= 1e-6
        assert rec["new_bid"] < 1.00


def test_action_label_follows_the_effective_move():
    # base falls purely because of campaign suppression -> not a "LOWER BID"
    out = eng._reconcile_action({"action": "HOLD", "new_bid": 0.85, "eff": 0.0}, 1.00)
    assert out["action"] == "LOWER BASE (COMPENSATED)"
    out = eng._reconcile_action({"action": "HOLD", "new_bid": 1.02, "eff": 0.05}, 1.00)
    assert out["action"] == "RAISE BID"
    out = eng._reconcile_action({"action": "RAISE BID", "new_bid": 0.90,
                                 "eff": -0.10}, 1.00)
    assert out["action"] == "LOWER BID"
    out = eng._reconcile_action({"action": "PAUSE", "new_bid": 0.0, "eff": 0.0}, 1.00)
    assert out["action"] == "PAUSE"


def test_effective_cpc_ceiling_enforced():
    kw = pd.DataFrame([{
        "Campaign Name": "C", "Campaign ID": 1, "Ad Group": "g", "Target": "kw",
        "Match Type": "EXACT", "Impressions": 1000, "Clicks": 20, "Conversions": 5,
        "Cost": 6.0, "ACOS": 0.10, "Current Bid": 1.00, "Sales": 60.0}]
    )
    kwn = eng.normalize(kw, True)
    coeffs = {1: (1.0, 2.0)}                   # campaign max adjustment = +200%
    out = eng.run_keyword_engine(kwn, {1: _stats()}, coeffs, max_cpc=1.85)
    bid = out.iloc[0]["Recommended Bid"]
    assert bid * (1 + 2.0) <= 1.85 + 1e-9
    assert bid == pytest.approx(0.61, abs=0.01)


# ---- low-click exposure bump ----
def _kw_df(bid, clicks, orders, cost, sales, camp="enabled", ag="enabled",
           status="enabled"):
    acos = cost / sales if sales else 0.0
    return eng.normalize(pd.DataFrame([{
        "Campaign Name": "C", "Campaign ID": 1, "Ad Group": "g", "Target": "kw",
        "Match Type": "EXACT", "Impressions": max(clicks * 20, 60), "Clicks": clicks,
        "Conversions": orders, "Cost": cost, "ACOS": acos, "Current Bid": bid,
        "Sales": sales, "Campaign status": camp, "Ad group status": ag,
        "Status": status}]), True)


def _run1(df, coeffs=None):
    return eng.run_keyword_engine(df, {1: _stats()},
                                  coeffs or {1: (1.0, 0.0)}).iloc[0]


def test_low_click_bump_lifts_a_watch_keyword():
    out = _run1(_kw_df(bid=1.00, clicks=6, orders=0, cost=3.0, sales=0))
    assert out["Action"] == "RAISE BID"
    assert out["Recommended Bid"] == pytest.approx(1.05, abs=0.001)


def test_low_click_bump_skipped_when_paused():
    for kw in (_kw_df(bid=1.00, clicks=6, orders=0, cost=3.0, sales=0, status="paused"),
               _kw_df(bid=1.00, clicks=6, orders=0, cost=3.0, sales=0, camp="paused")):
        out = _run1(kw)
        assert out["Recommended Bid"] == pytest.approx(1.00, abs=0.001)
        assert out["Action"] != "RAISE BID"


def test_low_click_bump_does_not_override_a_cut():
    out = _run1(_kw_df(bid=1.00, clicks=8, orders=1, cost=6.0, sales=6.0 / 0.60))
    assert out["Action"] == "LOWER BID"
    assert out["Recommended Bid"] < 1.00


def test_low_click_bump_not_applied_above_threshold():
    out = _run1(_kw_df(bid=1.00, clicks=11, orders=0, cost=4.0, sales=0))
    assert out["Recommended Bid"] == pytest.approx(1.00, abs=0.001)


def test_driver_vs_passenger_of_a_base_cut_are_labelled_differently():
    """A placement at the 0% floor CAUSES the base cut. A second bad placement that
    still has adjustment headroom merely RIDES that cut - its adjustment is held,
    and it must not be labelled as though it could not be fixed on its own."""
    g = pd.DataFrame([
        _pl("top_of_search", clicks=40, orders=9, cost=12, sales=130, adj=0.35),
        _pl("rest_of_search", clicks=20, orders=0, cost=20, sales=0, adj=0.0),
        _pl("product_page", clicks=42, orders=3, cost=40, sales=40 / 0.68, adj=0.20),
    ])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    by = {r["Placement"]: r for r in rows}
    assert lam == pytest.approx(0.85, abs=0.001)
    # the floor-bound placement is the cause
    assert by["Rest of Search"]["Action"] == "SUPPRESS (VIA BASE BIDS)"
    assert "THIS is why" in by["Rest of Search"]["Reason"]
    # the one with headroom rides it, adjustment untouched
    pp = by["Product Pages"]
    assert pp["Action"] == "SUPPRESS (BASE CUT, ADJ HELD)"
    assert pp["Recommended Adjustment"] == pytest.approx(0.20, abs=1e-9)
    assert pp["Effective Change"] == pytest.approx(-0.15, abs=0.01)
    # both bad placements land on the same intended step, no double-suppression
    assert by["Rest of Search"]["Effective Change"] == pytest.approx(-0.15, abs=0.01)


def test_same_placement_alone_uses_its_own_adjustment_instead():
    """Counterfactual for the test above: with no floor-bound neighbour there is no
    base cut at all - the adjustment absorbs the whole reduction."""
    g = pd.DataFrame([_pl("product_page", clicks=42, orders=3, cost=40,
                          sales=40 / 0.68, adj=0.20)])
    rows, lam, _ = eng.solve_campaign_placements(g, _stats())
    assert lam == pytest.approx(1.0, abs=1e-9)
    assert rows[0]["Action"] == "LOWER ADJUSTMENT"
    assert rows[0]["Recommended Adjustment"] == pytest.approx(0.02, abs=0.01)
    assert rows[0]["Effective Change"] == pytest.approx(-0.15, abs=0.01)


# ---- identifier pass-through (for finding/applying rows in the ads platform) ----
def test_identifier_columns_survive_as_exact_text():
    """15-digit IDs must round-trip digit-for-digit. Excel and pandas may read them
    as floats (2.9051584403233e+14), and any precision loss or scientific notation
    breaks copy-paste into the ads platform."""
    raw = pd.DataFrame([{
        "Campaign name": "C", "Campaign ID": 1, "Ad group": "g",
        "Keyword": "kw", "Match type": "EXACT", "Impressions": 100, "Clicks": 10,
        "Orders": 1, "Cost": 3.0, "ACOS": 0.20, "Bid": 0.50, "Ad Sales": 15.0,
        "Keyword/Target ID": 290515844032329,          # int64 from the export
        "Ad group ID": 5.24347575094442e14,            # float, as pandas may read it
    }])
    h = eng.harmonize_columns(raw, eng.KW_ALIASES, True)
    n = eng.normalize(h, True)
    assert n["Target ID"].iloc[0] == "290515844032329"
    assert n["Ad Group ID"].iloc[0] == "524347575094442"


def test_identifiers_appear_in_output_next_to_the_keyword():
    raw = pd.DataFrame([{
        "Campaign name": "C", "Campaign ID": 1, "Ad group": "g",
        "Keyword": "kw", "Match type": "EXACT", "Impressions": 100, "Clicks": 10,
        "Orders": 1, "Cost": 3.0, "ACOS": 0.20, "Bid": 0.50, "Ad Sales": 15.0,
        "Keyword/Target ID": 290515844032329, "Ad group ID": 524347575094442,
    }])
    n = eng.normalize(eng.harmonize_columns(raw, eng.KW_ALIASES, True), True)
    out = eng.run_keyword_engine(n, {1: _stats()}, {1: (1.0, 0.0)})
    cols = list(out.columns)
    assert cols.index("Target ID") == cols.index("Match Type") + 1
    assert out["Target ID"].iloc[0] == "290515844032329"


def test_missing_identifier_columns_are_simply_omitted():
    """The hand-built sample schema has no ID columns - the engine must not invent
    empty ones or fail."""
    n = _kw_df(bid=1.00, clicks=6, orders=0, cost=3.0, sales=0)
    out = eng.run_keyword_engine(n, {1: _stats()}, {1: (1.0, 0.0)})
    assert "Target ID" not in out.columns
