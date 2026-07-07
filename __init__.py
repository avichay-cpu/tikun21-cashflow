"""
מודל הכדאיות המצרפי — Feasibility Model
=======================================
מארגן את שלבי החישוב (5.1–5.10 במפרט):
מצב מוצע → שטחי מכירה → הכנסות → עלויות → פריסה חודשית →
סימולציית מימון → רווח יזמי ומדדי כדאיות.
"""
from __future__ import annotations
from dataclasses import dataclass, field
from typing import Dict, List
from .models import Project, Compensation, SpreadRule, CostLine
from .spreading import (build_expense_stream, build_revenue_stream,
                        sellable_res_area)
from .financing import simulate_financing, FinancingCostBreakdown


@dataclass
class FeasibilityResult:
    units_for_sale: int
    comp_units: int
    sellable_res_area: float
    total_revenue: float
    comp_ancillary: float
    cost_lines_total: float
    contingency: float
    financing_cost: float
    total_costs: float
    developer_profit: float
    profit_rate_on_cost: float
    profit_rate_on_revenue: float
    required_profit: float
    surplus_deficit: float
    is_feasible: bool
    financing_breakdown: FinancingCostBreakdown
    monthly_revenue: List[float] = field(default_factory=list)
    monthly_expense: List[float] = field(default_factory=list)
    monthly_interest: List[float] = field(default_factory=list)
    monthly_balance: List[float] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


def _compensation_ancillary(c: Compensation) -> float:
    if c.mode == "lump_sum":
        return c.comp_ancillary_lump
    return c.comp_units * (
        c.rent_monthly_per_unit * c.rent_months
        + c.moving_per_unit + c.legal_per_unit
        + c.guarantees_per_unit + c.other_per_unit
    )


def evaluate(project: Project) -> FeasibilityResult:
    p = project
    fin = p.financing
    months = fin.cashflow_months
    warnings: List[str] = []

    # 5.2 חלוקת יחידות
    comp_units = (p.compensation.comp_units if p.compensation.mode == "detailed"
                  else p.existing.existing_units)
    units_for_sale = p.program.total_new_units - comp_units
    if units_for_sale <= 0:
        warnings.append("אין יחידות למכירה — הפרויקט אינו כדאי מיסודו.")

    # 5.3 שטח מכירה  5.4 הכנסות
    res_area = sellable_res_area(p.program, p.revenue, max(units_for_sale, 0))
    revenue_res = res_area * p.revenue.price_res_per_sale_sqm
    revenue_commercial = p.program.commercial_area * p.revenue.price_commercial_per_sqm
    extra_parking = max(p.program.parking_total - comp_units, 0)
    revenue_parking = extra_parking * p.revenue.price_per_extra_parking
    total_revenue = revenue_res + revenue_commercial + revenue_parking

    # 5.5–5.7 עלויות (מסופקות כשורות עלות) + תמורה נלווית
    comp_ancillary = _compensation_ancillary(p.compensation)
    cost_lines_total = sum(l.amount for l in p.cost_lines) + comp_ancillary
    contingency = cost_lines_total * p.contingency_pct

    # אזהרת אימות למצב lump_sum
    if p.compensation.mode == "lump_sum" and p.compensation.comp_built_area > 0:
        implied = comp_units * (p.program.new_unit_avg_main_area
                                + p.program.service_area_per_unit)
        if abs(implied - p.compensation.comp_built_area) > 1:
            warnings.append(
                "מצב lump_sum: ודא ששטח דירות התמורה כלול בשורות עלות הבנייה "
                "(אחרת עלות הבנייה תיגרע).")

    # 5.8 פריסה חודשית + סימולציית מימון
    expense_lines = list(p.cost_lines)
    if comp_ancillary > 0:
        expense_lines = expense_lines + [
            CostLine("תמורה נלווית", comp_ancillary, SpreadRule.LINEAR)]

    monthly_expense = build_expense_stream(expense_lines, months)
    monthly_revenue = build_revenue_stream(
        total_revenue, months, fin.promo_months, p.revenue.down_payment_pct)

    # עמלות וערבויות (מזינות גם את הזרם החודשי וגם את דיווח המימון)
    total_costs_pre_fin = cost_lines_total + contingency
    accompaniment = total_costs_pre_fin * fin.fee_accompaniment
    sale_law_guar = total_revenue * fin.fee_sale_law_guarantee
    rent_guar = _rent_guarantee(p, fin)
    owners_guar = total_revenue * fin.fee_owners_guarantee
    non_utilization = _non_utilization(total_costs_pre_fin, monthly_expense, fin)
    fees_total = (accompaniment + sale_law_guar + rent_guar
                  + owners_guar + non_utilization)

    # הזרם החודשי כולל עמלות (פריסה לינארית של סך העמלות)
    fee_stream = _linear_stream(fees_total, months)
    monthly_expense_incl_fees = [monthly_expense[m] + fee_stream[m]
                                 for m in range(len(monthly_expense))]

    equity_cap = (cost_lines_total + contingency + fees_total) * fin.equity_pct
    fr = simulate_financing(monthly_revenue, monthly_expense_incl_fees,
                            equity_cap, fin.annual_interest_rate,
                            fin.equity_policy)

    fb = FinancingCostBreakdown(
        interest=fr.total_interest,
        accompaniment=accompaniment,
        sale_law_guarantee=sale_law_guar,
        rent_guarantee=rent_guar,
        owners_guarantee=owners_guar,
        non_utilization=non_utilization,
    )
    financing_cost = fb.total

    # 5.9 סך עלויות, רווח ומאזן
    total_costs = cost_lines_total + contingency + financing_cost
    developer_profit = total_revenue - total_costs
    profit_rate_on_cost = developer_profit / total_costs if total_costs else 0.0
    profit_rate_on_revenue = developer_profit / total_revenue if total_revenue else 0.0

    base = total_costs if p.profit_base == "on_cost" else total_revenue
    required_profit = base * p.target_profit_pct
    surplus_deficit = developer_profit - required_profit
    is_feasible = surplus_deficit >= 0

    return FeasibilityResult(
        units_for_sale=units_for_sale,
        comp_units=comp_units,
        sellable_res_area=res_area,
        total_revenue=total_revenue,
        comp_ancillary=comp_ancillary,
        cost_lines_total=cost_lines_total,
        contingency=contingency,
        financing_cost=financing_cost,
        total_costs=total_costs,
        developer_profit=developer_profit,
        profit_rate_on_cost=profit_rate_on_cost,
        profit_rate_on_revenue=profit_rate_on_revenue,
        required_profit=required_profit,
        surplus_deficit=surplus_deficit,
        is_feasible=is_feasible,
        financing_breakdown=fb,
        monthly_revenue=monthly_revenue,
        monthly_expense=monthly_expense_incl_fees,
        monthly_interest=fr.interest,
        monthly_balance=fr.balance,
        warnings=warnings,
    )


def _linear_stream(amount: float, months: int) -> List[float]:
    n = months + 1
    v = [0.0] * n
    per = amount / months
    for m in range(1, n):
        v[m] = per
    return v


def _rent_guarantee(p: Project, fin) -> float:
    c = p.compensation
    if c.mode == "detailed":
        annual_rent = c.comp_units * c.rent_monthly_per_unit * 12
        return fin.fee_rent_guarantee * annual_rent
    return fin.fee_rent_guarantee * c.comp_ancillary_lump * 0.0  # אין פירוט שכ"ד ב-lump


def _non_utilization(total_costs_pre_fin: float, monthly_expense: List[float],
                     fin) -> float:
    """עמלת אי-ניצול על היתרה הבלתי-מנוצלת של המסגרת, מצטבר חודשי."""
    monthly_rate = fin.fee_non_utilization / 12
    spent = 0.0
    fee = 0.0
    for m in range(1, len(monthly_expense)):
        spent += monthly_expense[m]
        undrawn = max(total_costs_pre_fin - spent, 0.0)
        fee += undrawn * monthly_rate
    return fee
