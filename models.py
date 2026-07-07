"""
מנוע תזרים ומימון חודשי — Monthly Cashflow & Financing Engine
==============================================================
הליבה של תחשיב המימון. במקום מקדם קירוב, המנוע פורס הוצאות והכנסות
על פני חודשי הבנייה ומריץ סימולציית אשראי מתגלגל חודש-בחודש.

השיטה זהה לאקסל הדוגמה של לשכת שמאות לנדאו:
  - הון עצמי נמשך ראשון, לכיסוי ההוצאה החודשית, עד לתקרה מצטברת.
  - היתרה מתגלגלת מחודש לחודש.
  - ריבית נצברת רק בחודשים בהם היתרה שלילית (אשראי בשימוש).

שינוי במספר חודשי הבנייה מעדכן אוטומטית את כל הפריסה ואת הריבית הנצברת.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import List


@dataclass
class FinancingResult:
    """תוצאת סימולציית המימון החודשית."""
    interest: List[float]           # ריבית לכל חודש (שלילי = עלות)
    balance: List[float]            # יתרה לסוף חודש אחרי ריבית
    balance_before_interest: List[float]
    equity_injection: List[float]   # הזרמת הון עצמי לכל חודש
    total_interest: float           # סך הריבית (עלות מימון, ערך חיובי)
    total_equity_used: float        # סך ההון העצמי שנוצל
    peak_debt: float                # שיא האשראי בשימוש (יתרה שלילית מקסימלית)

    @property
    def months(self) -> int:
        return len(self.interest)


def simulate_financing(
    revenue: List[float],
    expense: List[float],
    equity_cap: float,
    annual_rate: float,
    equity_policy: str = "gross_expense",
) -> FinancingResult:
    """
    מריץ את סימולציית המימון החודשית.

    פרמטרים
    --------
    revenue      : הכנסות חודשיות (זרם, אורך = מספר החודשים).
    expense      : הוצאות חודשיות כולל עמלות (אותו אורך).
    equity_cap   : תקרת ההון העצמי המצטבר (= equity_pct * total_costs).
    annual_rate  : ריבית שנתית על האשראי (למשל 0.055).
    equity_policy: "gross_expense" (כמו האקסל — ההון מכסה את מלוא ההוצאה
                   החודשית עד לתקרה) או "net_need" (ההון מכסה רק את הגירעון
                   התזרימי החודשי). ברירת המחדל נאמנה לאקסל.

    מחזיר FinancingResult.
    """
    if len(revenue) != len(expense):
        raise ValueError("revenue ו-expense חייבים להיות באותו אורך")

    n = len(revenue)
    monthly_rate = annual_rate / 12.0

    interest = [0.0] * n
    balance = [0.0] * n
    balance_before = [0.0] * n
    equity_inj = [0.0] * n

    remaining_equity = equity_cap
    prev_after_interest = 0.0
    peak_debt = 0.0

    for m in range(n):
        net_flow = revenue[m] - expense[m]

        if equity_policy == "net_need":
            need = max(-(prev_after_interest + net_flow), 0.0)
        else:  # gross_expense — נאמן לאקסל
            need = expense[m]

        inj = min(max(need, 0.0), remaining_equity)
        remaining_equity -= inj
        equity_inj[m] = inj

        bal_before = prev_after_interest + net_flow + inj
        intr = bal_before * monthly_rate if bal_before < 0 else 0.0
        after = bal_before + intr

        balance_before[m] = bal_before
        interest[m] = intr
        balance[m] = after
        prev_after_interest = after
        peak_debt = min(peak_debt, bal_before)

    return FinancingResult(
        interest=interest,
        balance=balance,
        balance_before_interest=balance_before,
        equity_injection=equity_inj,
        total_interest=-sum(interest),      # עלות חיובית
        total_equity_used=equity_cap - remaining_equity,
        peak_debt=-peak_debt,
    )


@dataclass
class FinancingCostBreakdown:
    """פירוט סך עלות המימון: ריבית + ערבויות ועמלות."""
    interest: float
    accompaniment: float          # עמלת ליווי
    sale_law_guarantee: float     # ערבות חוק מכר
    rent_guarantee: float         # ערבות שכ"ד לבעלים
    owners_guarantee: float       # ערבות בעלים
    non_utilization: float        # עמלת אי-ניצול מסגרת
    round_to: int = -4            # עיגול סך המימון (כמו האקסל: לרבבות)

    @property
    def total_raw(self) -> float:
        return (self.interest + self.accompaniment + self.sale_law_guarantee
                + self.rent_guarantee + self.owners_guarantee + self.non_utilization)

    @property
    def total(self) -> float:
        # עיגול לרבבות כמו ROUND(...,-4) באקסל
        factor = 10 ** (-self.round_to)
        return round(self.total_raw / factor) * factor
