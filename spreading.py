"""
מודל הנתונים — Data Models
==========================
מבני הקלט של תחשיב הכדאיות לפי תקן 21 (פינוי-בינוי).
כל הסכומים בש"ח, שטחים במ"ר.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from enum import Enum
from typing import List, Optional


class SpreadRule(str, Enum):
    """כללי פריסת עלות על פני חודשי הפרויקט."""
    AT_PERMIT = "at_permit"                     # מלוא הסכום בחודש ההיתר (חודש 0)
    LINEAR = "linear"                           # לינארי על חודשי הבנייה
    PARTIAL_PERMIT_LINEAR = "partial_permit_linear"  # X% בהיתר, היתרה לינארית
    START_END = "start_end"                     # מחצית בהתחלה, מחצית בסוף
    WINDOW = "window"                           # לינארי בחלון חודשים (start, duration)
    BASEMENT = "basement"                       # לפי חודשים×קומות מרתף
    RENT_PREPAID = "rent_prepaid"               # שכ"ד, תשלום שנה מראש


@dataclass
class CostLine:
    """שורת עלות בודדת עם כלל הפריסה שלה."""
    name: str
    amount: float
    rule: SpreadRule = SpreadRule.LINEAR
    permit_share: float = 0.0        # ל-PARTIAL_PERMIT_LINEAR
    start_month: int = 1             # ל-WINDOW / BASEMENT / skeleton / finishing
    duration: Optional[int] = None   # ל-WINDOW (מס' חודשים)


@dataclass
class ExistingState:
    existing_units: int
    existing_avg_area: float
    land_area: float
    existing_value_per_unit: Optional[float] = None


@dataclass
class ProposedProgram:
    total_new_units: int
    new_unit_avg_main_area: float
    service_area_per_unit: float = 0.0
    balcony_area_per_unit: float = 0.0
    storage_area_per_unit: float = 0.0
    commercial_area: float = 0.0
    parking_total: int = 0
    underground_area: float = 0.0


@dataclass
class RevenueParams:
    price_res_per_sale_sqm: float
    price_commercial_per_sqm: float = 0.0
    price_per_extra_parking: float = 0.0
    down_payment_pct: float = 0.20          # תקבול ראשון בחתימה
    factor_balcony: float = 0.30
    factor_storage: float = 0.50
    factor_service: float = 0.0
    sellable_res_area_override: Optional[float] = None


@dataclass
class FinancingParams:
    cashflow_months: int                    # חודשי תזרים/בנייה — המניע המרכזי
    promo_months: int = 1                   # תקופת מבצע מכירות מוקדמת
    annual_interest_rate: float = 0.055
    equity_pct: float = 0.25
    equity_policy: str = "gross_expense"    # נאמן לאקסל
    # שיעורי ערבויות ועמלות
    fee_accompaniment: float = 0.005        # עמלת ליווי (% מסך ההוצאות)
    fee_sale_law_guarantee: float = 0.008   # ערבות חוק מכר (% מההכנסות)
    fee_rent_guarantee: float = 0.035
    fee_owners_guarantee: float = 0.007
    fee_non_utilization: float = 0.004


@dataclass
class Compensation:
    mode: str = "detailed"                  # "detailed" | "lump_sum"
    # מצב detailed
    comp_units: int = 0
    comp_unit_main_area: float = 0.0
    rent_monthly_per_unit: float = 0.0
    rent_months: int = 0
    moving_per_unit: float = 0.0
    legal_per_unit: float = 0.0
    guarantees_per_unit: float = 0.0
    other_per_unit: float = 0.0
    # מצב lump_sum
    comp_built_area: float = 0.0
    comp_ancillary_lump: float = 0.0


@dataclass
class Project:
    name: str
    existing: ExistingState
    program: ProposedProgram
    revenue: RevenueParams
    financing: FinancingParams
    compensation: Compensation
    cost_lines: List[CostLine] = field(default_factory=list)
    contingency_pct: float = 0.0
    target_profit_pct: float = 0.15
    profit_base: str = "on_cost"            # "on_cost" | "on_revenue"
