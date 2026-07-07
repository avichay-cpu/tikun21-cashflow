"""
מחולל תזרים הכנסות והוצאות מקובץ תחשיב (התחדשות עירונית)
=========================================================
קורא את גיליון "מאוחד" מקובץ תחשיב, פורס את העלויות וההכנסות על ציר הזמן
לפי שיטת האקסל של הלשכה, מריץ את סימולציית המימון, ומחשב את אחוז המימון.

מיפוי הנתונים מבוסס על מבנה תבנית ה-"מאוחד" של הלשכה (שורות קבועות).
"""
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
try:
    from engine.financing import simulate_financing, FinancingCostBreakdown
except ModuleNotFoundError:
    from financing import simulate_financing, FinancingCostBreakdown

# ---------- הנחות עבודה (ניתן לעריכה לכל פרויקט) ----------
CFG = dict(
    months=42,                 # חודשי בנייה/תזרים — המניע המרכזי
    annual_rate=0.055,
    equity_pct=0.25,
    down_payment=0.20,         # תקבול ראשון בחתימה
    promo=1,                   # תקופת מבצע (חודשי חתימה מוקדמים)
    track20_share=0.20,        # חלק מסלול ה-20% מהפדיון (H4)
    planning_permit_share=0.50,
    # חלונות ביצוע (חודש התחלה, משך) — נגזרים ממבנה הפרויקט
    basement=(2, 9),           # 3 חודשים × 3 קומות מרתף
    skeleton=(10, 12),
    finishing=(20, 23),
    # שיעורי ערבויות ועמלות
    fee_accompaniment=0.005,
    fee_sale_law=0.008,
    fee_rent=0.035,
    fee_owners=0.007,
    fee_non_util=0.004,
)

# ---------- מיפוי שורות "מאוחד" ----------
def read_source(path):
    from openpyxl import load_workbook
    m = load_workbook(path, data_only=True)["מאוחד"]
    def g(r): return m[f"G{r}"].value or 0.0
    direct_constr = g(175)                       # סה"כ עלות בנייה ישירה
    basement_cost = g(170)                       # שטח תת-קרקעי
    src = dict(
        pidyon=m["J151"].value,                  # שווי פדיון ללא מע"מ
        demolition_dev=g(157) + g(159),          # הריסה ופינוי + פיתוח חצר
        planning=g(182),                         # תכנון ויועצים
        heitel=g(207),                           # היטל השבחה
        purchase_tax=g(206),                     # מס רכישה
        agrot=g(179) + g(180) + g(181),          # אגרות והיטלי פיתוח (עילי+תת+קיזוז)
        elec_res=g(183),                         # חיבור חשמל מגורים
        elec_com=g(184),                         # חיבור חשמל מסחרי
        legal=g(187),                            # משפטיות
        moving=g(196),                           # הובלות (העברה מפונים)
        tenant_advisors=g(198) + g(199) + g(200) + g(201),  # עו"ד+מפקח+שמאי+קרן
        rent=g(194),                             # שכ"ד דיור חלוף
        marketing=g(185),                        # שיווק ופרסום
        management=g(186) + g(188),              # תקורה+ניהול + פיקוח
        contingency=g(189),                      # בצ"מ
        basement=basement_cost,
        skeleton=(direct_constr - basement_cost) * 0.6,
        finishing=direct_constr - (direct_constr - basement_cost) * 0.6 - basement_cost,
        total_cost=g(213),                       # סה"כ עלות הקמה (ביקורת)
        # שווי דירות בעלים לערבות בעלים
        owners_value_main=m["I139"].value * m["I88"].value,   # עיקרי בלבד (מחודש 2)
        owners_value_full=m["I139"].value * m["I88"].value
                           + m["I139"].value * ((m["J88"].value or 0) + (m["K88"].value or 0)) * 0.5,  # + מרפסות (חודש 1)
    )
    return src


# ---------- פריסת עלות ----------
def spread_at_permit(amt, n):      # חודש 0
    v = [0.0]*n; v[0] = amt; return v

def spread_month1(amt, n):         # חודש בנייה ראשון
    v = [0.0]*n; v[1] = amt; return v

def spread_linear(amt, n):         # לינארי על חודשי הבנייה
    v = [0.0]*n
    per = amt/(n-1)
    for m in range(1, n): v[m] = per
    return v

def spread_partial_permit(amt, share, n):
    v = spread_linear(amt*(1-share), n); v[0] = amt*share; return v

def spread_start_end(amt, n):
    v = [0.0]*n; v[1] += amt/2; v[n-1] += amt/2; return v

def spread_window(amt, start, dur, n):
    v = [0.0]*n
    end = min(start+dur-1, n-1); span = end-start+1
    per = amt/span
    for m in range(start, end+1): v[m] = per
    return v

def spread_rent_prepaid(amt, n, months):
    """תשלום שכ"ד שנתי מראש: כל 12 חודשים, מתוחם לסך השכ"ד."""
    v = [0.0]*n
    annual = amt/months*12
    paid = 0.0; m = 1
    while m < n and paid < amt - 1:
        pay = min(annual, amt - paid)
        v[m] += pay; paid += pay; m += 12
    return v


def build_expense_lines(src, cfg):
    n = cfg["months"]+1
    lines = {
        "תכנון ויועצים": spread_partial_permit(src["planning"], cfg["planning_permit_share"], n),
        "היטל השבחה": spread_at_permit(src["heitel"], n),
        "מס רכישה": spread_at_permit(src["purchase_tax"], n),
        "אגרות והיטלי פיתוח": spread_at_permit(src["agrot"], n),
        "חיבור חשמל מגורים": spread_month1(src["elec_res"], n),
        "חיבור חשמל מסחרי": spread_month1(src["elec_com"], n),
        "משפטיות": spread_linear(src["legal"], n),
        "הובלות": spread_start_end(src["moving"], n),
        "יועצי דיירים (עו\"ד/מפקח/שמאי/קרן)": spread_linear(src["tenant_advisors"], n),
        "שכ\"ד דיור חלוף": spread_rent_prepaid(src["rent"], n, cfg["months"]),
        "שיווק ופרסום": spread_linear(src["marketing"], n),
        "ניהול תקורה ופיקוח": spread_linear(src["management"], n),
        "בלת\"מ": spread_linear(src["contingency"], n),
        "הריסה ופיתוח": spread_at_permit(src["demolition_dev"], n),
        "מרתפים": spread_window(src["basement"], *cfg["basement"], n),
        "שלד": spread_window(src["skeleton"], *cfg["skeleton"], n),
        "גמר": spread_window(src["finishing"], *cfg["finishing"], n),
    }
    return lines


def build_revenue(pidyon, cfg):
    """
    מודל הפדיון של הלשכה — מדרגה אלכסונית עם בלון במסירה:
      • מסלול 20% (חתימה): 20% בחודש 1, 80% בלון בחודש המסירה.
      • מסלול רגיל (80%): מכירה לינארית באצוות חודשיות (חודשים 2..M).
        כל אצווה משלמת מקדמה 20% בחודש המכירה, והיתרה לינארית עד המסירה.
    """
    months = cfg["months"]; n = months+1
    down = cfg["down_payment"]
    promo = cfg["promo"]
    track20_share = cfg["track20_share"]

    v = [0.0]*n
    I4 = pidyon * track20_share           # מסלול 20%
    I5 = pidyon * (1 - track20_share)     # מסלול רגיל

    # מסלול 20%: מקדמה בחודש 1, בלון במסירה
    v[promo] += I4 * down
    v[months] += I4 * (1 - down)

    # מסלול רגיל: אצוות חודשיות, חודשים (promo+1)..months
    n_batches = months - promo
    batch_value = I5 / n_batches
    batch_down = batch_value * down
    batch_balance = batch_value * (1 - down)
    for k in range(promo + 1, months + 1):
        v[k] += batch_down
        rem_months = months - k
        if rem_months > 0:
            per = batch_balance / rem_months
            for mm in range(k + 1, months + 1):
                v[mm] += per
        else:
            v[k] += batch_balance
    return v


def generate(path, overrides=None):
    cfg = dict(CFG)
    if overrides:
        cfg.update({k: v for k, v in overrides.items() if v is not None})
    n = cfg["months"]+1
    src = read_source(path)

    exp_lines = build_expense_lines(src, cfg)
    base_expense = [sum(exp_lines[k][m] for k in exp_lines) for m in range(n)]
    revenue = build_revenue(src["pidyon"], cfg)

    total_expense_base = sum(base_expense)

    # עמלות וערבויות
    accompaniment = total_expense_base * cfg["fee_accompaniment"]           # חד-פעמי
    sale_law = [revenue[m]*cfg["fee_sale_law"]/12 for m in range(n)]        # חודשי מול פדיון
    rent_guar_month = cfg["fee_rent"]/12*(src["rent"]/cfg["months"]*12)
    # ערבות בעלים: חודש 1 על מלוא השווי (כולל מרפסות), מחודש 2 על העיקרי בלבד
    owners_guar = [0.0]*n
    owners_guar[1] = cfg["fee_owners"]/12*src["owners_value_full"]
    for m in range(2, n):
        owners_guar[m] = cfg["fee_owners"]/12*src["owners_value_main"]
    # אי-ניצול: על היתרה הבלתי-מנוצלת של המסגרת (כולל הוצאות חודש ההיתר)
    non_util = [0.0]*n; spent = base_expense[0]
    for m in range(1, n):
        spent += base_expense[m]
        non_util[m] = max(total_expense_base-spent, 0.0)*cfg["fee_non_util"]/12

    guar_month = [0.0]*n
    guar_month[1] += accompaniment
    for m in range(1, n):
        guar_month[m] += sale_law[m] + rent_guar_month + owners_guar[m] + non_util[m]

    expense_incl = [base_expense[m] + guar_month[m] for m in range(n)]

    equity_cap = total_expense_base * cfg["equity_pct"]   # 25% מסך ההוצאות (כמו E27)
    fr = simulate_financing(revenue, expense_incl, equity_cap, cfg["annual_rate"])

    fb = FinancingCostBreakdown(
        interest=fr.total_interest,
        accompaniment=accompaniment,
        sale_law_guarantee=sum(sale_law),
        rent_guarantee=rent_guar_month*cfg["months"],
        owners_guarantee=sum(owners_guar),
        non_utilization=sum(non_util),
    )
    financing_pct = fb.total / src["total_cost"]

    return dict(src=src, cfg=cfg, exp_lines=exp_lines, base_expense=base_expense,
                revenue=revenue, expense_incl=expense_incl, guar_month=guar_month,
                fin=fr, breakdown=fb, financing_pct=financing_pct,
                total_expense_base=total_expense_base)


if __name__ == "__main__":
    path = sys.argv[1] if len(sys.argv) > 1 else "new.xlsx"
    R = generate(path)
    fb = R["breakdown"]
    print(f"פדיון (ללא מע\"מ):        {R['src']['pidyon']:>16,.0f}")
    print(f"סך עלות הקמה:            {R['src']['total_cost']:>16,.0f}")
    print(f"חודשי בנייה:             {R['cfg']['months']:>16}")
    print("-"*44)
    print(f"ריבית נצברת:             {fb.interest:>16,.0f}")
    print(f"עמלת ליווי:              {fb.accompaniment:>16,.0f}")
    print(f"ערבות חוק מכר:           {fb.sale_law_guarantee:>16,.0f}")
    print(f"ערבות שכ\"ד:              {fb.rent_guarantee:>16,.0f}")
    print(f"ערבות בעלים:             {fb.owners_guarantee:>16,.0f}")
    print(f"עמלת אי-ניצול:           {fb.non_utilization:>16,.0f}")
    print("-"*44)
    print(f"סך מימון:                {fb.total:>16,.0f}")
    print(f"אחוז מימון:              {R['financing_pct']:>15.2%}")
