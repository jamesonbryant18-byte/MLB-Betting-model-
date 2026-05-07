"""
fill_game.py
Copies MLB_Betting_Model.xlsx and fills in a specific game.
Usage: python3 fill_game.py
"""

import shutil
import os
from openpyxl import load_workbook

# ─────────────────────────────────────────────────────────────────────────────
# GAME DATA — Pirates (Away) at Diamondbacks (Home), May 7 2026
# Sources: FanGraphs, Baseball Savant, FanDuel, covers.com
# ─────────────────────────────────────────────────────────────────────────────

GAME = {
    # ── Teams ──────────────────────────────────────────────────────────────
    "home_team":   "Arizona Diamondbacks",
    "away_team":   "Pittsburgh Pirates",

    # ── Starting Pitchers ──────────────────────────────────────────────────
    "home_sp":     "Zac Gallen",
    "away_sp":     "Mitch Keller",

    # ── Pitcher Stats — Home (Zac Gallen, 2026) ────────────────────────────
    # Source: FanGraphs / Baseball Savant as of 5/7/2026
    "home_era":    4.45,   # ERA
    "home_fip":    3.63,   # FIP
    "home_xfip":   4.46,   # xFIP
    "home_siera":  3.80,   # SIERA (estimated between FIP/xFIP; FG shows ~3.32–3.80 range)
    "home_k9":     5.57,   # K/9
    "home_bb9":    2.78,   # BB/9
    "home_babip":  0.339,  # BABIP
    "home_hrfb":   5.4,    # HR/FB%
    "home_gb":     49.1,   # GB%
    "home_form":   4.45,   # Recent Form — last 5 starts ERA (using season avg)

    # ── Pitcher Stats — Away (Mitch Keller, 2026) ──────────────────────────
    # Source: FanGraphs as of 5/7/2026  (3-1 record, dominant early season)
    "away_era":    2.85,   # ERA
    "away_fip":    3.76,   # FIP
    "away_xfip":   2.89,   # xFIP (BABIP-luck inflated; low HR/FB%)
    "away_siera":  4.49,   # SIERA (suggests some regression likely)
    "away_k9":     6.80,   # K/9
    "away_bb9":    2.63,   # BB/9
    "away_babip":  0.254,  # BABIP (lucky — below avg)
    "away_hrfb":   2.5,    # HR/FB%
    "away_gb":     39.8,   # GB%
    "away_form":   2.85,   # Recent Form

    # ── Team Offense — Home (Arizona Diamondbacks, 2026) ───────────────────
    # Source: FanGraphs Teams / wRC+ leaderboard
    "home_woba":   0.314,  # wOBA
    "home_wrc":    96,     # wRC+ (4% below avg)
    "home_ops":    0.708,  # OPS (OBP .303 + SLG .405)
    "home_tbabip": 0.291,  # Team BABIP
    "home_barrel": 7.5,    # Barrel Rate % (estimated; below-avg power)

    # ── Team Offense — Away (Pittsburgh Pirates, 2026) ─────────────────────
    # Source: FanGraphs Teams
    "away_woba":   0.324,  # wOBA
    "away_wrc":    102,    # wRC+ (slightly above avg)
    "away_ops":    0.718,  # OPS (estimated OBP ~.330 + SLG ~.388)
    "away_tbabip": 0.295,  # Team BABIP (estimated)
    "away_barrel": 8.0,    # Barrel Rate % (near league avg)

    # ── Bullpen — Home (Arizona Diamondbacks) ──────────────────────────────
    # Source: insidethepen.com / season stats as proxy for last 14 days
    "home_bp_era":  4.78,  # Bullpen ERA
    "home_bp_whip": 1.38,  # Bullpen WHIP (estimated)

    # ── Bullpen — Away (Pittsburgh Pirates) ────────────────────────────────
    # Source: bucsdugout.com / season stats as proxy for last 14 days
    "away_bp_era":  3.80,  # Bullpen ERA
    "away_bp_whip": 1.28,  # Bullpen WHIP (estimated)

    # ── Context ────────────────────────────────────────────────────────────
    # Chase Field has a retractable roof — closed in May due to 89-101°F heat
    # Indoor climate-controlled: ~72°F inside, minimal wind effect
    "wind_speed":   3,     # mph (nominal; roof closed)
    "wind_dir":     "Cross",
    "temp":         72,    # °F (indoor Chase Field AC)
    "home_rest":    0,     # Days of rest — 3rd consecutive game of series
    "away_rest":    0,     # Days of rest — 3rd consecutive game of series

    # ── Vegas Lines (FanDuel / covers.com, 5/7/2026) ───────────────────────
    # Home (Arizona Diamondbacks): +112
    # Away (Pittsburgh Pirates):   -132
    "home_ml":  112,   # Arizona +112
    "away_ml": -132,   # Pittsburgh -132
}

# ─────────────────────────────────────────────────────────────────────────────
# COPY TEMPLATE AND FILL
# ─────────────────────────────────────────────────────────────────────────────

TEMPLATE = "/Users/jameson/Desktop/MLB-Betting-model/MLB_Betting_Model.xlsx"
OUTPUT   = "/Users/jameson/Desktop/MLB-Betting-model/Pirates-Diamondbacks 5-7-2026.xlsx"

shutil.copy2(TEMPLATE, OUTPUT)
print(f"Copied template → {OUTPUT}")

wb = load_workbook(OUTPUT)
gi = wb["Game Input"]

# ── Teams ──────────────────────────────────────────────────────────────────
gi["B4"].value = GAME["home_team"]
gi["C4"].value = GAME["away_team"]
gi["B5"].value = GAME["home_sp"]
gi["C5"].value = GAME["away_sp"]

# ── Home pitcher stats ──────────────────────────────────────────────────────
gi["B8"].value  = GAME["home_era"]
gi["B9"].value  = GAME["home_fip"]
gi["B10"].value = GAME["home_xfip"]
gi["B11"].value = GAME["home_siera"]
gi["B12"].value = GAME["home_k9"]
gi["B13"].value = GAME["home_bb9"]
gi["B14"].value = GAME["home_babip"]
gi["B15"].value = GAME["home_hrfb"]
gi["B16"].value = GAME["home_gb"]
gi["B17"].value = GAME["home_form"]

# ── Away pitcher stats ──────────────────────────────────────────────────────
gi["C8"].value  = GAME["away_era"]
gi["C9"].value  = GAME["away_fip"]
gi["C10"].value = GAME["away_xfip"]
gi["C11"].value = GAME["away_siera"]
gi["C12"].value = GAME["away_k9"]
gi["C13"].value = GAME["away_bb9"]
gi["C14"].value = GAME["away_babip"]
gi["C15"].value = GAME["away_hrfb"]
gi["C16"].value = GAME["away_gb"]
gi["C17"].value = GAME["away_form"]

# ── Home offense ────────────────────────────────────────────────────────────
gi["B20"].value = GAME["home_woba"]
gi["B21"].value = GAME["home_wrc"]
gi["B22"].value = GAME["home_ops"]
gi["B23"].value = GAME["home_tbabip"]
gi["B24"].value = GAME["home_barrel"]

# ── Away offense ────────────────────────────────────────────────────────────
gi["C20"].value = GAME["away_woba"]
gi["C21"].value = GAME["away_wrc"]
gi["C22"].value = GAME["away_ops"]
gi["C23"].value = GAME["away_tbabip"]
gi["C24"].value = GAME["away_barrel"]

# ── Bullpen ──────────────────────────────────────────────────────────────────
gi["B27"].value = GAME["home_bp_era"]
gi["B28"].value = GAME["home_bp_whip"]
gi["C27"].value = GAME["away_bp_era"]
gi["C28"].value = GAME["away_bp_whip"]

# ── Context ──────────────────────────────────────────────────────────────────
# B31/C31 are VLOOKUP formulas — leave intact, they'll resolve when Excel opens
gi["B32"].value = GAME["wind_speed"]
gi["B33"].value = GAME["wind_dir"]
gi["B34"].value = GAME["temp"]
gi["B35"].value = GAME["home_rest"]
gi["C36"].value = GAME["away_rest"]

# ── Vegas lines ───────────────────────────────────────────────────────────────
gi["B39"].value = GAME["home_ml"]
gi["C39"].value = GAME["away_ml"]

wb.save(OUTPUT)
print("All game data written.")
print(f"\nGame: {GAME['away_team']} (Away) at {GAME['home_team']} (Home)")
print(f"Starters: {GAME['away_sp']} vs {GAME['home_sp']}")
print(f"Vegas: Home {GAME['home_ml']:+d} | Away {GAME['away_ml']:+d}")
print(f"\nFile saved: {OUTPUT}")
print(f"File size: {os.path.getsize(OUTPUT):,} bytes")
