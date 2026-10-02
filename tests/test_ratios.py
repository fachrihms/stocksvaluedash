from app.models.financial import FinancialPeriod
from app.utils.ratios import cash_dividend_payout


def test_cash_dividend_payout_uses_same_period_operating_cash_flow():
    """Dividen kas dibandingkan dengan arus kas operasi periode yang sama."""
    period = FinancialPeriod(
        period_label="31 December 2025",
        dividends_paid=-20_508,
        operating_cash_flow=44_694,
        net_income_parent=32_769,
    )

    assert round(cash_dividend_payout(period), 1) == 45.9


def test_cash_dividend_payout_is_unavailable_when_operating_cash_flow_negative():
    period = FinancialPeriod(
        period_label="31 December 2025",
        dividends_paid=-100,
        operating_cash_flow=-50,
    )

    assert cash_dividend_payout(period) is None
