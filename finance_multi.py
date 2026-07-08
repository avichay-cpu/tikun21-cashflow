"""
חישוב מימון לכל מתחם — מחבר את החילוץ (extract) למנוע המימון.
"""
import sys, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from financing import simulate_financing, FinancingCostBreakdown
from generate_cashflow import (spread_at_permit, spread_month1, spread_linear,
                               spread_partial_permit, spread_start_end,
                               spread_window, spread_rent_prepaid, build_revenue, CFG)
import extract as X

RULEMAP = {
    "at_permit": lambda amt, n, cfg: spread_at_permit(amt, n),
    "month1":    lambda amt, n, cfg: spread_month1(amt, n),
    "linear":    lambda amt, n, cfg: spread_linear(amt, n),
    "start_end": lambda amt, n, cfg: spread_start_end(amt, n),
    "permit_linear": lambda amt, n, cfg: spread_partial_permit(amt, cfg["planning_permit_share"], n),
    "rent":      lambda amt, n, cfg: spread_rent_prepaid(amt, n, cfg["months"]),
    "win_bsmt":  lambda amt, n, cfg: spread_window(amt, *cfg["basement"], n),
    "win_skel":  lambda amt, n, cfg: spread_window(amt, *cfg["skeleton"], n),
    "win_fin":   lambda amt, n, cfg: spread_window(amt, *cfg["finishing"], n),
}


def finance_compound(ex, cfg):
    """מקבל נתוני מתחם מחולצים + הנחות, מחזיר אחוז מימון ופירוט."""
    n = cfg["months"] + 1
    # פריסת הוצאות
    base_expense = [0.0] * n
    for name, amt, rule in ex["cost_lines"]:
        vec = RULEMAP.get(rule, RULEMAP["linear"])(amt, n, cfg)
        for m in range(n):
            base_expense[m] += vec[m]
    total_expense_base = sum(base_expense)

    # הכנסות (עם/בלי מסלול 20%)
    revenue = build_revenue(ex["pidyon"], cfg)

    # עמלות וערבויות
    acc = total_expense_base * cfg["fee_accompaniment"]
    sale = [revenue[m] * cfg["fee_sale_law"] / 12 for m in range(n)]
    rent_total = sum(a for nm, a, r in ex["cost_lines"] if r == "rent")
    rent_g = cfg["fee_rent"] / 12 * (rent_total / cfg["months"] * 12) if rent_total else 0.0
    # ערבות בעלים על שווי דירות הדיירים (עמודה K)
    own_full = cfg["fee_owners"] / 12 * ex["owners"]
    non_util = [0.0] * n; spent = base_expense[0]
    for m in range(1, n):
        spent += base_expense[m]
        non_util[m] = max(total_expense_base - spent, 0.0) * cfg["fee_non_util"] / 12

    guar = [0.0] * n
    guar[1] += acc
    for m in range(1, n):
        guar[m] += sale[m] + rent_g + own_full + non_util[m]
    expense_incl = [base_expense[m] + guar[m] for m in range(n)]

    equity_cap = total_expense_base * cfg["equity_pct"]
    fr = simulate_financing(revenue, expense_incl, equity_cap, cfg["annual_rate"])

    fb = FinancingCostBreakdown(
        interest=fr.total_interest, accompaniment=acc,
        sale_law_guarantee=sum(sale), rent_guarantee=rent_g * cfg["months"],
        owners_guarantee=own_full * cfg["months"], non_utilization=sum(non_util),
    )
    pct = fb.total / ex["total_before_fin"] if ex["total_before_fin"] else 0.0
    return dict(financing=fb.total, pct=pct, breakdown=fb,
                total_cost=ex["total_before_fin"], pidyon=ex["pidyon"])


if __name__ == "__main__":
    path = sys.argv[1]
    cfg = dict(CFG)
    print(f"{'מתחם':22} {'פדיון':>16} {'עלות':>16} {'מימון':>14} {'אחוז':>7}")
    print("-" * 80)
    for sh in X.find_compound_sheets(path):
        ex = X.extract_compound(path, sh)
        r = finance_compound(ex, cfg)
        print(f"{sh:22} {r['pidyon']:>16,.0f} {r['total_cost']:>16,.0f} "
              f"{r['financing']:>14,.0f} {r['pct']:>7.2%}")
