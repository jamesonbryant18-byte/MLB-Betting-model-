"""
scan_games.py
Runs the model engine across all eligible games today (5/7/2026),
ranks by edge, and prints results. Excludes Yankees/Rangers and Twins/Nationals.
Sources: ESPN API (odds), FanGraphs (pitcher/offense stats), Covers.com (bullpen)
"""

import math
import shutil
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# GAME DATA — 5/7/2026  (excluded: NYY vs TEX, MIN vs WSH)
# ─────────────────────────────────────────────────────────────────────────────

GAMES = {

    # ── GAME 1: Cleveland Guardians at Kansas City Royals ────────────────────
    # Venue: Kauffman Stadium | 2:10 PM ET | Park Factor: 98
    # Weather: 63°F, 16 mph Cross wind  (outdoors, afternoon)
    # Starters: Slade Cecconi (CLE) vs Seth Lugo (KC)
    # Moneyline: KC -143 / CLE +119
    "CLE @ KC": {
        "home_team": "Kansas City Royals",
        "away_team": "Cleveland Guardians",
        "home_sp": "Seth Lugo",
        "away_sp": "Slade Cecconi",
        # Pitcher stats — Seth Lugo (FanGraphs 2026)
        "home_era": 2.68, "home_fip": 2.64, "home_xfip": 3.81, "home_siera": 3.50,
        "home_k9": 7.63, "home_bb9": 2.68, "home_babip": 0.325,
        "home_hrfb": 2.2, "home_gb": 39.3, "home_form": 2.68,
        # Pitcher stats — Slade Cecconi (FanGraphs 2026)
        "away_era": 6.56, "away_fip": 5.83, "away_xfip": 4.71, "away_siera": 5.20,
        "away_k9": 7.07, "away_bb9": 3.28, "away_babip": 0.313,
        "away_hrfb": 17.8, "away_gb": 44.1, "away_form": 6.56,
        # Team offense — KC (FanGraphs 2026)
        "home_woba": 0.315, "home_wrc": 94, "home_ops": 0.703,
        "home_tbabip": 0.287, "home_barrel": 7.5,
        # Team offense — CLE (FanGraphs 2026)
        "away_woba": 0.311, "away_wrc": 96, "away_ops": 0.690,
        "away_tbabip": 0.295, "away_barrel": 7.5,
        # Bullpen — Covers.com 2026
        "home_bp_era": 4.80, "home_bp_whip": 1.46,
        "away_bp_era": 4.03, "away_bp_whip": 1.26,
        # Context
        "park_factor": 98, "wind_speed": 16, "wind_dir": "Cross", "temp": 63,
        "home_rest": 0, "away_rest": 0,
        # Vegas (ESPN / covers.com)
        "home_ml": -143, "away_ml": 119,
    },

    # ── GAME 2: Cincinnati Reds at Chicago Cubs ──────────────────────────────
    # Venue: Wrigley Field | 2:20 PM ET | Park Factor: 103
    # Weather: 55°F, 9 mph Cross wind  (outdoors, afternoon)
    # Starters: Rhett Lowder (CIN) vs Shota Imanaga (CHC)
    # Moneyline: CHC -199 / CIN +163
    "CIN @ CHC": {
        "home_team": "Chicago Cubs",
        "away_team": "Cincinnati Reds",
        "home_sp": "Shota Imanaga",
        "away_sp": "Rhett Lowder",
        # Pitcher stats — Shota Imanaga (FanGraphs 2026)
        "home_era": 2.40, "home_fip": 2.73, "home_xfip": 3.47, "home_siera": 3.20,
        "home_k9": 9.36, "home_bb9": 2.18, "home_babip": 0.214,
        "home_hrfb": 6.1, "home_gb": 34.0, "home_form": 2.40,
        # Pitcher stats — Rhett Lowder (FanGraphs 2026)
        "away_era": 5.09, "away_fip": 3.22, "away_xfip": 4.32, "away_siera": 4.44,
        "away_k9": 6.62, "away_bb9": 3.57, "away_babip": 0.304,
        "away_hrfb": 2.5, "away_gb": 43.2, "away_form": 5.09,
        # Team offense — CHC (FanGraphs 2026): OPS .781, wRC+ 121
        "home_woba": 0.340, "home_wrc": 121, "home_ops": 0.781,
        "home_tbabip": 0.285, "home_barrel": 9.5,
        # Team offense — CIN (FanGraphs 2026)
        "away_woba": 0.309, "away_wrc": 88, "away_ops": 0.689,
        "away_tbabip": 0.280, "away_barrel": 7.5,
        # Bullpen — Covers.com 2026
        "home_bp_era": 3.86, "home_bp_whip": 1.20,
        "away_bp_era": 4.23, "away_bp_whip": 1.58,
        # Context
        "park_factor": 103, "wind_speed": 9, "wind_dir": "Cross", "temp": 55,
        "home_rest": 0, "away_rest": 0,
        # Vegas (ESPN)
        "home_ml": -199, "away_ml": 163,
    },

    # ── GAME 3: New York Mets at Colorado Rockies ────────────────────────────
    # Venue: Coors Field | 3:10 PM ET | Park Factor: 115
    # Weather: 72°F, 8 mph Out (afternoon Denver, typical Coors blow-out)
    # Starters: Christian Scott (NYM) vs Jose Quintana (COL)
    # Moneyline: NYM -149 / COL +123
    "NYM @ COL": {
        "home_team": "Colorado Rockies",
        "away_team": "New York Mets",
        "home_sp": "Jose Quintana",
        "away_sp": "Christian Scott",
        # Pitcher stats — Jose Quintana (ESPN/DK 2026 — 24.1 IP, 1-2, 4.07 ERA)
        "home_era": 4.07, "home_fip": 5.00, "home_xfip": 5.13, "home_siera": 5.00,
        "home_k9": 4.44, "home_bb9": 4.08, "home_babip": 0.310,
        "home_hrfb": 12.0, "home_gb": 42.0, "home_form": 4.07,
        # Pitcher stats — Christian Scott (FanGraphs 2026 — 6.1 IP, very small sample)
        "away_era": 4.26, "away_fip": 4.70, "away_xfip": 5.67, "away_siera": 5.50,
        "away_k9": 12.79, "away_bb9": 7.11, "away_babip": 0.167,
        "away_hrfb": 10.0, "away_gb": 25.0, "away_form": 4.26,
        # Team offense — COL (FanGraphs 2026)
        "home_woba": 0.324, "home_wrc": 90, "home_ops": 0.726,
        "home_tbabip": 0.322, "home_barrel": 8.5,
        # Team offense — NYM (FanGraphs 2026)
        "away_woba": 0.290, "away_wrc": 84, "away_ops": 0.640,
        "away_tbabip": 0.273, "away_barrel": 6.5,
        # Bullpen — Covers.com 2026
        "home_bp_era": 4.42, "home_bp_whip": 1.36,
        "away_bp_era": 3.82, "away_bp_whip": 1.27,
        # Context — Coors outdoor, afternoon
        "park_factor": 115, "wind_speed": 8, "wind_dir": "Out", "temp": 72,
        "home_rest": 0, "away_rest": 0,
        # Vegas (ESPN)
        "home_ml": 123, "away_ml": -149,
    },

    # ── GAME 4: Pittsburgh Pirates at Arizona Diamondbacks ───────────────────
    # Venue: Chase Field | 3:40 PM ET | Park Factor: 99
    # Weather: 72°F indoor (retractable roof closed — 101°F outside)
    # Starters: Mitch Keller (PIT) vs Zac Gallen (ARI)
    # Moneyline: ARI +112 / PIT -132
    "PIT @ ARI": {
        "home_team": "Arizona Diamondbacks",
        "away_team": "Pittsburgh Pirates",
        "home_sp": "Zac Gallen",
        "away_sp": "Mitch Keller",
        # Pitcher stats — Zac Gallen (FanGraphs 2026)
        "home_era": 4.45, "home_fip": 3.63, "home_xfip": 4.46, "home_siera": 3.80,
        "home_k9": 5.57, "home_bb9": 2.78, "home_babip": 0.339,
        "home_hrfb": 5.4, "home_gb": 49.1, "home_form": 4.45,
        # Pitcher stats — Mitch Keller (FanGraphs 2026)
        "away_era": 2.85, "away_fip": 3.76, "away_xfip": 2.89, "away_siera": 4.49,
        "away_k9": 6.80, "away_bb9": 2.63, "away_babip": 0.254,
        "away_hrfb": 2.5, "away_gb": 39.8, "away_form": 2.85,
        # Team offense — ARI (FanGraphs 2026)
        "home_woba": 0.314, "home_wrc": 96, "home_ops": 0.708,
        "home_tbabip": 0.291, "home_barrel": 7.5,
        # Team offense — PIT (FanGraphs 2026)
        "away_woba": 0.324, "away_wrc": 102, "away_ops": 0.718,
        "away_tbabip": 0.295, "away_barrel": 8.0,
        # Bullpen — Covers.com 2026
        "home_bp_era": 4.78, "home_bp_whip": 1.38,
        "away_bp_era": 3.80, "away_bp_whip": 1.28,
        # Context — Chase Field indoor AC
        "park_factor": 99, "wind_speed": 3, "wind_dir": "Cross", "temp": 72,
        "home_rest": 0, "away_rest": 0,
        # Vegas (ESPN / FanDuel)
        "home_ml": 112, "away_ml": -132,
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# MODEL ENGINE — mirrors the Excel formulas exactly
# ─────────────────────────────────────────────────────────────────────────────

def pitcher_score(era, fip, xfip, siera, k9, bb9, form):
    raw = (
        ((4.20 - fip)  / 4.20 * 30) +
        ((4.15 - xfip) / 4.15 * 25) +
        ((4.10 - siera)/ 4.10 * 20) +
        ((k9   - 8.9)  / 8.9  * 10) +
        ((3.2  - bb9)  / 3.2  * 10) +
        ((4.20 - form) / 4.20 *  5) +
        50
    )
    return max(0, min(100, raw))

def offense_score(wrc, woba, barrel, tbabip):
    raw = (
        ((wrc    - 100)   / 100   * 40) +
        ((woba   - 0.315) / 0.315 * 35) +
        ((barrel - 8.5)   / 8.5   * 15) +
        ((tbabip - 0.315) / 0.315 * 10) +
        50
    )
    return max(0, min(100, raw))

def bullpen_score(bp_era, bp_whip):
    raw = (
        ((4.30 - bp_era)  / 4.30 * 60) +
        ((2.0  - bp_whip) / 2.0  * 40) +
        50
    )
    return max(0, min(100, raw))

def park_adj(pf):
    return (pf - 100) / 100 * 0.3

def wind_adj(speed, direction):
    if direction == "Out":
        return speed * 0.02
    elif direction == "In":
        return -speed * 0.015
    return 0.0

def temp_adj(temp):
    if temp < 72:
        return -(72 - temp) / 10 * 0.005
    return 0.0

def rest_adj(home_rest, away_rest):
    if home_rest - away_rest >= 2:
        return 0.03
    elif away_rest - home_rest >= 2:
        return -0.03
    return 0.0

def expected_runs(off_score, opp_pitch_score, bp_score,
                  pk_adj, w_adj, t_adj, r_adj, home=True):
    runs = (
        3.5 +
        (off_score / 100) * 3 +
        (-opp_pitch_score / 100) * 1.5 +
        (bp_score / 100) * 1 +
        pk_adj + w_adj + t_adj + (r_adj if home else -r_adj)
    )
    return max(1.5, min(12, runs))

def win_prob(run_diff):
    return 1 / (1 + math.exp(-run_diff * 0.8))

def implied_prob(ml):
    if ml < 0:
        return abs(ml) / (abs(ml) + 100)
    return 100 / (ml + 100)

# ─────────────────────────────────────────────────────────────────────────────
# RUN MODEL ACROSS ALL GAMES
# ─────────────────────────────────────────────────────────────────────────────

results = []

for label, g in GAMES.items():
    # Scores
    h_pitch = pitcher_score(g["home_era"], g["home_fip"], g["home_xfip"],
                             g["home_siera"], g["home_k9"], g["home_bb9"], g["home_form"])
    a_pitch = pitcher_score(g["away_era"], g["away_fip"], g["away_xfip"],
                             g["away_siera"], g["away_k9"], g["away_bb9"], g["away_form"])
    h_off   = offense_score(g["home_wrc"], g["home_woba"], g["home_barrel"], g["home_tbabip"])
    a_off   = offense_score(g["away_wrc"], g["away_woba"], g["away_barrel"], g["away_tbabip"])
    h_bp    = bullpen_score(g["home_bp_era"], g["home_bp_whip"])
    a_bp    = bullpen_score(g["away_bp_era"], g["away_bp_whip"])

    # Context
    pk  = park_adj(g["park_factor"])
    wd  = wind_adj(g["wind_speed"], g["wind_dir"])
    td  = temp_adj(g["temp"])
    rd  = rest_adj(g["home_rest"], g["away_rest"])

    # Expected runs
    h_runs = expected_runs(h_off, a_pitch, h_bp, pk, wd, td, rd, home=True)
    a_runs = expected_runs(a_off, h_pitch, a_bp, pk, wd, td, rd, home=False)

    # Win probability
    diff   = h_runs - a_runs
    h_wp   = win_prob(diff)
    a_wp   = 1 - h_wp

    # Vegas implied
    h_imp  = implied_prob(g["home_ml"])
    a_imp  = implied_prob(g["away_ml"])

    # Edge
    h_edge = h_wp - h_imp
    a_edge = a_wp - a_imp
    best_edge = max(h_edge, a_edge)
    best_side = g["home_team"] if h_edge >= a_edge else g["away_team"]
    best_ml   = g["home_ml"]   if h_edge >= a_edge else g["away_ml"]

    results.append({
        "label": label,
        "game": g,
        "h_pitch": h_pitch, "a_pitch": a_pitch,
        "h_off": h_off, "a_off": a_off,
        "h_bp": h_bp, "a_bp": a_bp,
        "h_runs": h_runs, "a_runs": a_runs,
        "total_runs": h_runs + a_runs,
        "h_wp": h_wp, "a_wp": a_wp,
        "h_imp": h_imp, "a_imp": a_imp,
        "h_edge": h_edge, "a_edge": a_edge,
        "best_edge": best_edge,
        "best_side": best_side,
        "best_ml": best_ml,
    })

# Sort by best edge descending
results.sort(key=lambda x: x["best_edge"], reverse=True)

# ─────────────────────────────────────────────────────────────────────────────
# PRINT RESULTS
# ─────────────────────────────────────────────────────────────────────────────

print("=" * 65)
print("  MLB EDGE SCAN — MAY 7, 2026")
print("  Excluded: Yankees/Rangers | Twins/Nationals")
print("=" * 65)

for i, r in enumerate(results, 1):
    g = r["game"]
    edge_flag = " *** VALUE BET ***" if r["best_edge"] > 0.05 else (" (watch)" if r["best_edge"] > 0.02 else "")
    print(f"\n#{i}  {r['label']}")
    print(f"    Starters:   {g['home_sp']} (HOME) vs {g['away_sp']} (AWAY)")
    print(f"    Pitcher Scores:  Home {r['h_pitch']:.1f} | Away {r['a_pitch']:.1f}")
    print(f"    Offense Scores:  Home {r['h_off']:.1f}  | Away {r['a_off']:.1f}")
    print(f"    Bullpen Scores:  Home {r['h_bp']:.1f}  | Away {r['a_bp']:.1f}")
    print(f"    Exp Runs:    {g['home_team']} {r['h_runs']:.2f} | {g['away_team']} {r['a_runs']:.2f}  (Total: {r['total_runs']:.2f})")
    print(f"    Model Win%:  {g['home_team']} {r['h_wp']*100:.1f}% | {g['away_team']} {r['a_wp']*100:.1f}%")
    print(f"    Vegas Impl:  {g['home_team']} {r['h_imp']*100:.1f}% | {g['away_team']} {r['a_imp']*100:.1f}%")
    print(f"    Edge:        {g['home_team']} {r['h_edge']*100:+.1f}% | {g['away_team']} {r['a_edge']*100:+.1f}%")
    print(f"    BEST BET:    {r['best_side']}  ML {r['best_ml']:+d}  Edge {r['best_edge']*100:+.1f}%{edge_flag}")

print("\n" + "=" * 65)
winner = results[0]
print(f"  TOP EDGE GAME: {winner['label']}")
print(f"  BET: {winner['best_side']}  ML {winner['best_ml']:+d}  Edge {winner['best_edge']*100:+.1f}%")
print("=" * 65)

# ─────────────────────────────────────────────────────────────────────────────
# FILL MODEL FOR TOP EDGE GAME
# ─────────────────────────────────────────────────────────────────────────────

top = results[0]
g   = top["game"]

ht = g["home_team"].replace(" ", "-").replace("/", "-")
at = g["away_team"].replace(" ", "-").replace("/", "-")
filename = f"{at}-vs-{ht}-5-7-2026.xlsx"
output_path = f"/Users/jameson/Desktop/MLB-Betting-model/{filename}"

TEMPLATE = "/Users/jameson/Desktop/MLB-Betting-model/MLB_Betting_Model.xlsx"
shutil.copy2(TEMPLATE, output_path)

wb = load_workbook(output_path)
gi = wb["Game Input"]

gi["B4"].value  = g["home_team"]
gi["C4"].value  = g["away_team"]
gi["B5"].value  = g["home_sp"]
gi["C5"].value  = g["away_sp"]

gi["B8"].value  = g["home_era"];  gi["C8"].value  = g["away_era"]
gi["B9"].value  = g["home_fip"];  gi["C9"].value  = g["away_fip"]
gi["B10"].value = g["home_xfip"]; gi["C10"].value = g["away_xfip"]
gi["B11"].value = g["home_siera"];gi["C11"].value = g["away_siera"]
gi["B12"].value = g["home_k9"];   gi["C12"].value = g["away_k9"]
gi["B13"].value = g["home_bb9"];  gi["C13"].value = g["away_bb9"]
gi["B14"].value = g["home_babip"];gi["C14"].value = g["away_babip"]
gi["B15"].value = g["home_hrfb"]; gi["C15"].value = g["away_hrfb"]
gi["B16"].value = g["home_gb"];   gi["C16"].value = g["away_gb"]
gi["B17"].value = g["home_form"]; gi["C17"].value = g["away_form"]

gi["B20"].value = g["home_woba"]; gi["C20"].value = g["away_woba"]
gi["B21"].value = g["home_wrc"];  gi["C21"].value = g["away_wrc"]
gi["B22"].value = g["home_ops"];  gi["C22"].value = g["away_ops"]
gi["B23"].value = g["home_tbabip"];gi["C23"].value = g["away_tbabip"]
gi["B24"].value = g["home_barrel"];gi["C24"].value = g["away_barrel"]

gi["B27"].value = g["home_bp_era"];gi["C27"].value = g["away_bp_era"]
gi["B28"].value = g["home_bp_whip"];gi["C28"].value = g["away_bp_whip"]

gi["B32"].value = g["wind_speed"]
gi["B33"].value = g["wind_dir"]
gi["B34"].value = g["temp"]
gi["B35"].value = g["home_rest"]
gi["C36"].value = g["away_rest"]

gi["B39"].value = g["home_ml"]
gi["C39"].value = g["away_ml"]

wb.save(output_path)
print(f"\nModel saved: {filename}")
