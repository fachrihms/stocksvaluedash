"""
Test penanganan mata uang pelaporan asing.

Bug nyata yang di-guard di sini: POWR (PT Cikarang Listrindo) menyajikan
lapkeu dalam USD. EPS-nya 0.0048 USD dibandingkan langsung dengan harga
saham Rupiah membuat PER muncul 222.826x (seharusnya ~15-25x). Semua
rasio per-share (PER, PBV, dividend yield) WAJIB melalui konversi kurs.

Jalankan: py -m pytest tests/test_currency_handling.py -v
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.parsers.financial_xlsx import _currency_code
from app.models.financial import FinancialPeriod
from app.models.price import PriceSeries, PriceBar
from app.utils.ratios import build_ratio_summary
from datetime import date


POWR_2024 = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "sample_data",
    "FinancialStatement-2024-Tahunan-POWR.xlsx"
)


def test_currency_code_extraction():
    """Kode ISO diekstrak dari nilai metadata template IDX."""
    assert _currency_code("Dollar Amerika / USD") == "USD"
    assert _currency_code("Rupiah / IDR") == "IDR"
    assert _currency_code("Euro / EUR") == "EUR"
    assert _currency_code("Yen Jepang / JPY") == "JPY"
    assert _currency_code("") is None
    print("PASS: ekstraksi kode mata uang benar")


def test_powr_detected_as_usd():
    """File POWR asli terdeteksi ber-mata-uang USD (bukan IDR)."""
    if not os.path.exists(POWR_2024):
        print("SKIP: sample POWR tidak ada")
        return
    from app.parsers.financial_xlsx import parse_financial_xlsx

    result = parse_financial_xlsx(POWR_2024)
    for period in result.periods:
        assert period.reporting_currency == "USD", (
            f"Mata uang laporan POWR harus 'USD', dapat {period.reporting_currency!r}"
        )
    print("PASS: POWR terdeteksi ber-mata-uang USD")


def _powr_scenario():
    """Skenario PER/PBV POWR: file asli + harga & saham realistis."""
    if not os.path.exists(POWR_2024):
        return None
    from app.parsers.financial_xlsx import parse_financial_xlsx

    result = parse_financial_xlsx(POWR_2024)
    current, prior = result.periods[0], result.periods[1]
    prices = PriceSeries(source_filename="test.csv", bars=[
        # Harga POWR di kisaran Rp1.000-1.300 per lembar
        PriceBar(trade_date=date(2026, 9, 18), close=1150)
    ])
    # ~25,3 miliar lembar saham beredar POWR
    return current, prior, prices, 25_303_300_000


def test_per_pbv_without_rate_are_none():
    """Tanpa kurs, PER/PBV/yield TIDAK boleh muncul (None) -- bukan angka
    gila 222.826x seperti bug semula."""
    scenario = _powr_scenario()
    if scenario is None:
        print("SKIP: sample POWR tidak ada")
        return
    current, prior, prices, shares = scenario

    ratios = build_ratio_summary(current, prior, prices, shares, exchange_rate_to_idr=None)
    assert ratios["reporting_currency"] == "USD"
    assert ratios["currency_rate_missing"] is True
    assert ratios.get("per") is None, f"PER harus None tanpa kurs, dapat {ratios.get('per')}"
    assert ratios.get("pbv") is None, f"PBV harus None tanpa kurs, dapat {ratios.get('pbv')}"
    assert ratios.get("dividend_yield_pct") is None
    # Rasio tanpa harga (DER, ROE, dsb) tetap terhitung normal
    assert ratios.get("der") is not None
    assert ratios.get("roe_pct") is not None
    print("PASS: tanpa kurs, PER/PBV/yield aman (None), rasio lain tetap normal")


def test_per_pbv_with_usd_rate_sane():
    """Dengan kurs USD->IDR realistis, PER & PBV masuk rentang wajar."""
    scenario = _powr_scenario()
    if scenario is None:
        print("SKIP: sample POWR tidak ada")
        return
    current, prior, prices, shares = scenario

    # EPS 0.0048 USD x 16.250 = Rp78/lembar; harga Rp1.150 -> PER ~14.7x
    # BVPS: equity parent 708.044.164 USD x 16.250 / 25,3 miliar lembar = ~Rp455 -> PBV ~2.5x
    ratios = build_ratio_summary(current, prior, prices, shares, exchange_rate_to_idr=16_250)
    assert ratios["exchange_rate_applied"] == 16_250
    assert ratios["currency_rate_missing"] is False

    eps_idr = ratios.get("eps_in_idr")
    assert eps_idr is not None and 70 < eps_idr < 90, f"EPS IDR aneh: {eps_idr}"
    per = ratios.get("per")
    assert per is not None and 5 < per < 40, (
        f"PER = {per} -- di luar rentang wajar; bug PER 222.826x mungkin kembali"
    )
    bvps = ratios.get("book_value_per_share")
    assert bvps is not None and 300 < bvps < 600, f"BVPS IDR aneh: {bvps}"
    pbv = ratios.get("pbv")
    assert pbv is not None and 1.0 < pbv < 6.0, f"PBV = {pbv} di luar rentang wajar"
    dy = ratios.get("dividend_yield_pct")
    assert dy is not None and 1 < dy < 25, f"Dividend yield aneh: {dy}"
    print(f"PASS: dengan kurs 16.250 -- EPS Rp{eps_idr:.0f}, PER {per:.1f}x, PBV {pbv:.2f}x, yield {dy:.1f}%")


def test_idr_report_unaffected_by_currency_logic():
    """Laporan IDR tidak terpengaruh logika kurs (tidak butuh kurs,
    nilai per-share tetap dihitung)."""
    period = FinancialPeriod(
        period_label="31 December 2025",
        reporting_currency=None,  # tidak terdeteksi -> diperlakukan IDR
        basic_eps=170.84,
        total_equity_parent=39_509_108,
        reporting_unit_multiplier=1_000_000,
    )
    prices = PriceSeries(source_filename="t.csv", bars=[
        PriceBar(trade_date=date(2026, 9, 18), close=2900)
    ])
    ratios = build_ratio_summary(period, period, prices, 32_250_000_000)
    assert ratios["reporting_currency"] == "IDR"
    assert ratios["currency_rate_missing"] is False
    assert ratios.get("per") is not None and 10 < ratios["per"] < 25, ratios.get("per")
    print("PASS: laporan IDR tidak terpengaruh logika kurs")


if __name__ == "__main__":
    test_currency_code_extraction()
    test_powr_detected_as_usd()
    test_per_pbv_without_rate_are_none()
    test_per_pbv_with_usd_rate_sane()
    test_idr_report_unaffected_by_currency_logic()
    print("\nSemua test mata uang lolos.")
