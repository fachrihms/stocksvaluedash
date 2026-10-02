"""
Test parser XLSX laporan keuangan (financial_xlsx.py).

Dua format workbook diuji:

A. Workbook "gaya lama" (label satu kolom, header tanggal di atas tabel)
   -- dibangun on-the-fly dengan openpyxl.

B. Workbook GAYA ASLI ekspor XBRL IDX (diverifikasi dari file nyata
   POWR): sheet bernama kode XBRL ("1000000" info umum, "3210000" neraca,
   "3311000" laba rugi, "3510000" arus kas, "36xxxxx" catatan, "Context"
   metadata teknis), label Inggris di KANAN angka, header tabel cuma
   berisi nama konteks ("CurrentYearInstant") -- tanggal periode & nama
   emiten WAJIB diambil dari sheet "1000000".

C. File asli POWR di sample_data/ (kalau ada) -- cross-file consistency
   antara laporan 2023 dan 2024.

Jalankan: py -m pytest tests/test_financial_xlsx_parser.py -v
"""

import sys
import os
import tempfile
from datetime import datetime

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook

from app.parsers.financial_xlsx import parse_financial_xlsx, _cell_to_text, _rows_from_sheet


# ============================================================
# A. Workbook gaya lama (header tanggal di atas tabel)
# ============================================================

# Angka dalam satuan JUTA (reporting_unit_multiplier=1_000_000), meniru
# gaya laporan ISAT. Nilai dipilih bebas tapi konsisten (rasio masuk akal).
XLSX_ISAT_LIKE = {
    "total_assets": 118_629_871,
    "total_current_assets": 19_003_874,
    "total_liabilities": 79_120_763,
    "total_current_liabilities": 30_499_266,
    "total_equity": 39_509_108,
    "total_equity_parent": 38_100_000,
    "cash_and_equivalents": 5_075_324,
    "revenue": 56_518_141,
    "profit_before_tax": 7_284_647,
    "net_income_parent": 5_509_703,
    "operating_cash_flow": 21_550_578,
    "investing_cash_flow": -12_097_383,
    "financing_cash_flow": -8_861_591,
    "capex": -12_359_655,
    "dividends_paid": -2_702_400,
    "basic_eps": 170.84,
}


def _wb_from_rows_dict(sheet_rows: dict[str, list[list]]) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)
    for sheet_name, rows in sheet_rows.items():
        ws = wb.create_sheet(title=sheet_name)
        for row in rows:
            ws.append(row)
    return wb


def _build_idx_like_workbook() -> Workbook:
    """
    Workbook meniru ekspor Excel IDX (gaya lama): sheet metadata + neraca
    konsolidasian + laba rugi konsolidasian + arus kas konsolidasian +
    versi "Entitas Saja" dengan angka BERBEDA (untuk menguji prioritas).
    """
    wb = Workbook()
    wb.remove(wb.active)

    # --- Sheet metadata / General information ---
    ws_meta = wb.create_sheet(title="General Information")
    ws_meta.append(["Nama Emiten", "PT Indosat Tbk", "Entity Name"])
    ws_meta.append(["Tanggal akhir periode berjalan", "December 31, 2025"])
    ws_meta.append(["Tanggal akhir periode sebelumnya", "December 31, 2024"])
    ws_meta.append(["Pembulatan yang digunakan", "dalam penyajian jumlah"])
    ws_meta.append([None, "Jutaan", "In Million"])

    # --- Neraca konsolidasian ---
    ws_bs = wb.create_sheet(title="Laporan Posisi Keuangan Konsolidasian")
    ws_bs.append(["Laporan Posisi Keuangan", None, None])
    ws_bs.append(["(dinyatakan dalam jutaan Rupiah)", None, None])
    ws_bs.append(["Keterangan", "31 Desember 2025", "31 Desember 2024"])
    ws_bs.append([2, "Aset lancar", 19_003_874, 17_000_000])
    ws_bs.append([3, "Kas dan setara kas", 5_075_324, 4_500_000])
    ws_bs.append([4, "Jumlah aset lancar", 19_003_874, 17_000_000])
    ws_bs.append([5, "Jumlah aset", 118_629_871, 110_000_000])
    ws_bs.append([6, "Liabilitas jangka pendek", 30_499_266, 28_000_000])
    ws_bs.append([7, "Jumlah liabilitas jangka pendek", 30_499_266, 28_000_000])
    ws_bs.append([8, "Jumlah liabilitas", 79_120_763, 72_000_000])
    ws_bs.append([9, "Ekuitas yang diatribusikan ke entitas induk", 38_100_000, 35_000_000])
    ws_bs.append([10, "Jumlah ekuitas", 39_509_108, 38_000_000])

    # --- Laba rugi konsolidasian ---
    ws_is = wb.create_sheet(title="Laporan Laba Rugi Konsolidasian")
    ws_is.append(["Keterangan", "31 Desember 2025", "31 Desember 2024"])
    ws_is.append(["Penjualan dan pendapatan", 56_518_141, 52_000_000])
    ws_is.append(["Beban bunga dan keuangan", 1_200_000, 1_100_000])
    ws_is.append(["Jumlah laba (rugi) sebelum pajak", 7_284_647, 6_500_000])
    ws_is.append(["Jumlah laba (rugi)", 5_800_000, 5_300_000])
    ws_is.append(["Laba (rugi) yang dapat diatribusikan ke:", None, None])
    ws_is.append(["   - entitas induk", 5_509_703, 5_100_000])
    ws_is.append(["   - non-pengendali", 290_297, 200_000])
    # Frasa lengkap persis seperti template IDX ("dari operasi yang
    # dilanjutkan") -- disambiguasi basic_eps di engine bergantung pada
    # kata "dilanjutkan"/"continuing" di sekitar baris ini.
    ws_is.append(["Laba per saham dasar dari operasi yang dilanjutkan", 170.84, 158.20])

    # --- Arus kas konsolidasian (label varian XLSX "Jumlah kas bersih") ---
    ws_cf = wb.create_sheet(title="Laporan Arus Kas Konsolidasian")
    ws_cf.append(["Keterangan", "31 Desember 2025", "31 Desember 2024"])
    ws_cf.append(["Penerimaan dari penjualan aset tetap", 100_000, 90_000])
    ws_cf.append(["Jumlah kas bersih diterima dari aktivitas operasi", 21_550_578, 19_000_000])
    ws_cf.append(["Perolehan aset tetap", -12_359_655, -11_000_000])
    ws_cf.append(["Jumlah kas bersih diterima dari aktivitas investasi", -12_097_383, -10_500_000])
    ws_cf.append(["Pembayaran dividen", -2_702_400, -2_400_000])
    ws_cf.append(["Jumlah kas bersih diterima dari aktivitas pendanaan", -8_861_591, -7_800_000])

    # --- Versi "Entitas Saja" dengan angka BERBEDA -- kalau parser salah
    #     baca sheet ini, nilai test di atas akan gagal. ---
    entitas_values = {k: (v * 2 if isinstance(v, (int, float)) else v) for k, v in XLSX_ISAT_LIKE.items()}
    ws_bs_e = wb.create_sheet(title="Laporan Posisi Keuangan (Entitas Saja)")
    ws_bs_e.append(["Keterangan", "31 Desember 2025", "31 Desember 2024"])
    ws_bs_e.append(["Jumlah aset", entitas_values["total_assets"], 220_000_000])
    ws_is_e = wb.create_sheet(title="Laporan Laba Rugi (Entitas Saja)")
    ws_is_e.append(["Keterangan", "31 Desember 2025", "31 Desember 2024"])
    ws_is_e.append(["Penjualan dan pendapatan", entitas_values["revenue"], 104_000_000])
    ws_cf_e = wb.create_sheet(title="Laporan Arus Kas (Entitas Saja)")
    ws_cf_e.append(["Keterangan", "31 Desember 2025", "31 Desember 2024"])
    ws_cf_e.append(["Jumlah kas bersih diterima dari aktivitas operasi", entitas_values["operating_cash_flow"], 38_000_000])

    return wb


def _write_tmp_xlsx(wb: Workbook) -> str:
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    wb.save(path)
    return path


def test_cell_to_text_renders_dates_and_floats():
    """Tanggal Excel dirender '31 December 2025' (format yang dikenali
    _detect_period_labels), float bulat tanpa '.0'."""
    assert _cell_to_text(datetime(2025, 12, 31)) == "31 December 2025"
    assert _cell_to_text(56_518_141.0) == "56518141"
    assert _cell_to_text(170.84) == "170.84"
    assert _cell_to_text(None) is None
    assert _cell_to_text("  label  ") == "label"
    print("PASS: konversi sel Excel ke teks benar")


def test_rows_from_sheet_drops_note_reference_numbers():
    """Angka referensi catatan (1-2 digit) di kiri label dibuang, tapi
    angka nilai periode SETELAH label tetap utuh."""
    wb = Workbook()
    ws = wb.create_sheet("BS")
    ws.append([2, "Kas dan setara kas", 5_075_324, 4_500_000])
    rows = _rows_from_sheet(ws)
    assert rows == ["Kas dan setara kas 5075324 4500000"], rows
    print("PASS: kolom referensi catatan dibuang dengan benar")


def test_xlsx_end_to_end_idx_like_workbook():
    """Workbook sintetis format IDX (gaya lama): semua field kunci terbaca
    dengan nilai yang benar, satuan jutaan terdeteksi, is_interim=False,
    dan angka yang terambil berasal dari sheet KONSOLIDASIAN (bukan
    Entitas Saja)."""
    path = _write_tmp_xlsx(_build_idx_like_workbook())
    try:
        result = parse_financial_xlsx(path)
        period = result.periods[0]

        failures = []
        for field, expected in XLSX_ISAT_LIKE.items():
            actual = getattr(period, field)
            if actual != expected:
                failures.append(f"{field}: got {actual}, expected {expected}")
        assert not failures, "Mismatch:\n" + "\n".join(failures)

        # Periode & metadata. Label periode diambil dari HEADER tabel Excel
        # "31 Desember 2025" (Indonesia) -- bukan dari metadata berbahasa
        # Inggris di sheet General Information -- karena baris header tabel
        # muncul lebih dulu saat pemindaian baris. Parser tunggal tidak
        # melewati merge_periods, jadi label mentahnya tetap "31 Desember
        # 2025" persis seperti di header sheet.
        assert result.entity_name == "PT Indosat Tbk"
        assert period.period_label == "31 Desember 2025"
        assert result.periods[1].period_label == "31 Desember 2024"
        assert period.reporting_unit_multiplier == 1_000_000
        assert period.is_interim is False
        assert period.fiscal_year == 2025
        assert result.parse_warnings == [], f"Warning tak terduga: {result.parse_warnings}"
    finally:
        os.unlink(path)
    print("PASS: end-to-end workbook IDX-like terbaca sempurna (konsolidasian menang)")


def test_xlsx_cash_flow_variant_labels():
    """Label XLSX 'Jumlah kas bersih diterima dari aktivitas ...' dan
    'Perolehan aset tetap' terpetakan ke field yang benar -- bukan ketukar-
    tukar antar aktivitas (regression guard bug window PDF)."""
    path = _write_tmp_xlsx(_build_idx_like_workbook())
    try:
        result = parse_financial_xlsx(path)
        period = result.periods[0]
        assert period.operating_cash_flow == 21_550_578
        assert period.investing_cash_flow == -12_097_383
        assert period.financing_cash_flow == -8_861_591
        assert period.capex == -12_359_655
        assert period.dividends_paid == -2_702_400
    finally:
        os.unlink(path)
    print("PASS: label cash flow varian XLSX terpetakan benar")


def test_xlsx_prior_period_filled():
    """Kolom periode kedua (prior year) ikut terisi dari workbook yang sama."""
    path = _write_tmp_xlsx(_build_idx_like_workbook())
    try:
        result = parse_financial_xlsx(path)
        prior = result.periods[1]
        assert prior.revenue == 52_000_000
        assert prior.total_assets == 110_000_000
        assert prior.net_income_parent == 5_100_000
        assert prior.operating_cash_flow == 19_000_000
        assert prior.basic_eps == 158.20
        assert prior.fiscal_year == 2024
    finally:
        os.unlink(path)
    print("PASS: periode prior terisi dari kolom kedua")


# ============================================================
# B. Workbook gaya ASLI ekspor XBRL IDX (sheet kode XBRL)
# ============================================================

def _build_xbrl_code_workbook() -> Workbook:
    """
    Meniru file asli POWR: sheet 'Context' (metadata teknis yang panjang),
    '1000000' (General information: nama entitas, kode, tanggal ISO,
    pembulatan "Satuan Penuh"), laporan inti 3210000/3311000/3510000 dengan
    label Indonesia kiri - angka tengah - label Inggris KANAN, dan sheet
    catatan 3611000 berisi label duplikat (harus TIDAK ikut mengganggu).
    """
    wb = Workbook()
    wb.remove(wb.active)

    # --- Context: ~ratusan baris metadata teknis di file asli; cukup
    #     sampel barisnya di sini untuk memastikan TIDAK mengganggu. ---
    ws_ctx = wb.create_sheet("Context")
    ws_ctx.append(["Context"])
    ws_ctx.append([None])
    ws_ctx.append(["entity"])
    ws_ctx.append(["identifier", "entityCode"])
    ws_ctx.append(["scheme", "http://www.idx.co.id/xbrl"])
    ws_ctx.append(["period"])
    ws_ctx.append(["instant", "2011-09-30 00:00:00"])
    ws_ctx.append(["CurrentYearInstant"])
    ws_ctx.append(["entity"])
    ws_ctx.append(["identifier", "entityCode"])
    ws_ctx.append(["scheme", "http://www.idx.co.id/xbrl"])
    ws_ctx.append(["period"])
    ws_ctx.append(["instant", "2023-12-31 00:00:00"])
    ws_ctx.append(["PriorEndYearInstant"])
    ws_ctx.append(["entity"])
    ws_ctx.append(["identifier", "entityCode"])
    ws_ctx.append(["period"])
    ws_ctx.append(["instant", "2022-12-31 00:00:00"])

    # --- 1000000: General information (SUMBER KEBENARAN metadata) ---
    ws_gi = wb.create_sheet("1000000")
    ws_gi.append(["[1000000] General information"])
    ws_gi.append([None])
    ws_gi.append(["Informasi umum", "General information"])
    ws_gi.append(["Nama entitas", "PT Listrik Perkasa Tbk.", "Entity name"])
    ws_gi.append(["Kode entitas", "LPWR", "Entity code"])
    ws_gi.append(["Industri utama entitas", "Infrastruktur / Infrastructure", "Entity main industry"])
    ws_gi.append(["Periode penyampaian laporan keuangan", "Tahunan / Annual", "Period of financial statements submissions"])
    ws_gi.append(["Tanggal awal periode berjalan", "2023-01-01", "Current period start date"])
    ws_gi.append(["Tanggal akhir periode berjalan", "2023-12-31", "Current period end date"])
    ws_gi.append(["Tanggal akhir tahun sebelumnya", "2022-12-31", "Prior year end date"])
    ws_gi.append(["Mata uang pelaporan", "Rupiah / IDR", "Description of presentation currency"])
    ws_gi.append(["Pembulatan yang digunakan dalam penyajian jumlah", "Satuan Penuh / Full Amount", "Level of rounding used in financial statements"])

    # --- 3210000: Neraca (label ID kiri, angka, label EN KANAN) ---
    ws_bs = wb.create_sheet("3210000")
    ws_bs.append(["[3210000] Statement of financial position"])
    ws_bs.append([None])
    ws_bs.append(["Laporan posisi keuangan", "Statement of financial position"])
    ws_bs.append(["CurrentYearInstant", "PriorEndYearInstant"])
    ws_bs.append(["Kas dan setara kas", 244_291_095, 305_083_705, "Cash and cash equivalents"])
    ws_bs.append(["Jumlah aset lancar", 556_308_492, 542_054_110, "Total current assets"])
    ws_bs.append(["Jumlah aset", 1_324_229_288, 1_361_618_473, "Total assets"])
    ws_bs.append(["Liabilitas jangka pendek", 57_816_392, 54_751_246, "Current liabilities"])
    ws_bs.append(["Jumlah liabilitas jangka pendek", 57_816_392, 54_751_246, "Total current liabilities"])
    ws_bs.append(["Jumlah liabilitas", 620_104_942, 661_857_508, "Total liabilities"])
    ws_bs.append(["Ekuitas yang diatribusikan kepada pemilik entitas induk", "Equity attributable to equity owners of parent entity"])
    ws_bs.append(["Jumlah ekuitas yang diatribusikan kepada pemilik entitas induk", 704_124_346, 699_760_965, "Total equity attributable to equity owners of parent entity"])
    ws_bs.append(["Kepentingan non-pengendali", 5_000_000, 4_000_000, "Non-controlling interests"])
    ws_bs.append(["Jumlah ekuitas", 709_124_346, 703_760_965, "Total equity"])
    ws_bs.append(["Jumlah liabilitas dan ekuitas", 1_324_229_288, 1_361_618_473, "Total liabilities and equity"])

    # --- 3311000: Laba rugi ---
    ws_is = wb.create_sheet("3311000")
    ws_is.append(["[3311000] Statement of profit or loss"])
    ws_is.append([None])
    ws_is.append(["CurrentYearDuration", "PriorYearDuration"])
    ws_is.append(["Penjualan dan pendapatan usaha", 546_079_025, 550_450_870, "Sales and revenue"])
    ws_is.append(["Beban bunga dan keuangan", 27_698_082, 29_108_772, "Interest and finance costs"])
    ws_is.append(["Jumlah laba (rugi) sebelum pajak", 112_281_595, 111_300_660, "Total profit (loss) before tax"])
    ws_is.append(["Laba (rugi) yang dapat diatribusikan", "Profit (loss) attributable to"])
    ws_is.append(["Laba (rugi) yang dapat diatribusikan ke entitas induk", 76_976_795, 72_535_694, "Profit (loss) attributable to parent entity"])
    ws_is.append(["Laba (rugi) yang dapat diatribusikan ke kepentingan non-pengendali", 1_000_000, 900_000, "Profit (loss) attributable to non-controlling interests"])
    ws_is.append(["Jumlah laba (rugi)", 77_976_795, 73_435_694, "Total profit (loss)"])
    ws_is.append(["Laba (rugi) per saham dasar dari operasi yang dilanjutkan", 0.0049, 0.0046, "Basic earnings (loss) per share from continuing operations"])

    # --- 3510000: Arus kas (angka capex/dividen POSITIF persis file asli) ---
    ws_cf = wb.create_sheet("3510000")
    ws_cf.append(["[3510000] Statement of cash flows"])
    ws_cf.append([None])
    ws_cf.append(["CurrentYearInstant", "PriorYearInstant"])
    ws_cf.append(["Penerimaan dari pelanggan", 541_956_178, 550_087_420, "Receipts from customers"])
    ws_cf.append(["Jumlah arus kas bersih yang diperoleh dari (digunakan untuk) aktivitas operasi", 154_865_050, 146_552_733, "Total net cash flows from operating activities"])
    # DECOY regression guard: baris "Pembayaran untuk ..." TANPA "aset
    # tetap" lalu diikuti baris disposal (Penerimaan dari penjualan aset
    # tetap) -- pola persis yang membuat capex POWR salah ambil angka.
    ws_cf.append(["Pembayaran untuk future contracts, forward contracts, option contracts, dan swap contracts", "Payments for derivatives"])
    ws_cf.append(["Penerimaan dari penjualan aset tetap", 672_380, 232_316, "Proceeds from disposal of property and equipment"])
    ws_cf.append(["Pembayaran untuk perolehan aset tetap", 27_852_349, 37_610_366, "Payments for acquisition of property and equipment"])
    ws_cf.append(["Jumlah arus kas bersih yang diperoleh dari (digunakan untuk) aktivitas investasi", -141_153_769, -42_204_228, "Total net cash flows from investing activities"])
    ws_cf.append(["Pembayaran dividen dari aktivitas pendanaan", 74_838_811, 67_873_126, "Dividends paid from financing activities"])
    ws_cf.append(["Jumlah arus kas bersih yang diperoleh dari (digunakan untuk) aktivitas pendanaan", -75_953_819, -69_307_037, "Total net cash flows from financing activities"])

    # --- 3611000: Catatan (label duplikat -- HARUS tidak menang atas
    #     laporan inti karena urutan pemrosesan sheet) ---
    ws_note = wb.create_sheet("3611000")
    ws_note.append(["[3611000] Notes - Property, plant, and equipment"])
    ws_note.append([None])
    ws_note.append(["Aset tetap", "Property, plant, and equipment"])
    ws_note.append(["Kas dan setara kas", 111_111_111, 222_222_222, "Decoy cash in notes"])

    return wb


def test_xlsx_real_idx_xbrl_code_sheet_format():
    """Format ASLI ekspor XBRL IDX (sheet kode): emiten dari sheet 1000000,
    periode ISO dari metadata, satuan 'Satuan Penuh' = multiplier 1,
    is_interim=False, dan nilai dari laporan inti (bukan sheet catatan).
    Regression guard untuk bug 'entitas tidak terdeteksi' + capex salah
    ambil angka disposal."""
    path = _write_tmp_xlsx(_build_xbrl_code_workbook())
    try:
        result = parse_financial_xlsx(path)
        current, prior = result.periods[0], result.periods[1]

        # Metadata eksplisit
        assert result.entity_name == "PT Listrik Perkasa Tbk."
        assert current.entity_code == "LPWR"
        assert current.reporting_unit_multiplier == 1
        assert current.period_label == "2023-12-31"
        assert prior.period_label == "2022-12-31"
        assert current.fiscal_year == 2023 and prior.fiscal_year == 2022
        assert current.is_interim is False and prior.is_interim is False
        assert result.parse_warnings == [], f"Warning: {result.parse_warnings}"

        # Nilai dari laporan inti -- BUKAN dari sheet catatan (decoy 111M)
        assert current.cash_and_equivalents == 244_291_095
        assert current.total_assets == 1_324_229_288
        assert current.total_current_assets == 556_308_492
        assert current.total_liabilities == 620_104_942
        assert current.total_current_liabilities == 57_816_392
        # parent = 704124346, total termasuk NCI = 709124346
        assert current.total_equity_parent == 704_124_346
        assert current.total_equity == 709_124_346
        assert current.revenue == 546_079_025
        assert current.profit_before_tax == 112_281_595
        assert current.net_income_parent == 76_976_795
        assert current.net_income_total == 77_976_795
        assert current.basic_eps == 0.0049
        assert current.interest_expense == 27_698_082
        assert current.operating_cash_flow == 154_865_050
        assert current.investing_cash_flow == -141_153_769
        assert current.financing_cash_flow == -75_953_819
        # Capex HARUS dari baris "Pembayaran untuk perolehan aset tetap"
        # (positif di file -> dinormalisasi negatif), BUKAN angka baris
        # disposal (672380) yang kebetulan di atasnya.
        assert current.capex == -27_852_349
        assert prior.capex == -37_610_366
        assert current.dividends_paid == -74_838_811
        assert prior.dividends_paid == -67_873_126
    finally:
        os.unlink(path)
    print("PASS: format asli XBRL IDX (sheet kode) terbaca lengkap & benar")


# ============================================================
# C. File asli POWR di sample_data/
# ============================================================

POWR_2023 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sample_data", "FinancialStatement-2023-Tahunan-POWR.xlsx")
POWR_2024 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sample_data", "FinancialStatement-2024-Tahunan-POWR.xlsx")


def test_real_powr_files():
    """File asli unduhan IDX (POWR 2023 & 2024): entitas terdeteksi, nilai
    kunci cocok angka laporan asli, dan cross-file konsisten."""
    if not os.path.exists(POWR_2023):
        print("SKIP: sample POWR tidak ada di sample_data/")
        return

    r2023 = parse_financial_xlsx(POWR_2023)
    cur = r2023.periods[0]
    prior = r2023.periods[1]

    assert r2023.entity_name == "PT Cikarang Listrindo Tbk."
    assert cur.entity_code == "POWR"
    assert cur.reporting_unit_multiplier == 1
    assert cur.period_label == "2023-12-31" and prior.period_label == "2022-12-31"
    assert cur.fiscal_year == 2023 and prior.fiscal_year == 2022
    assert cur.is_interim is False and prior.is_interim is False
    assert r2023.parse_warnings == [], f"Warning: {r2023.parse_warnings}"

    # Nilai terverifikasi manual dari file asli (USD, satuan penuh)
    assert cur.total_assets == 1_324_229_288
    assert cur.total_equity == 704_124_346
    assert cur.revenue == 546_079_025
    assert cur.net_income_parent == 76_976_795
    assert cur.basic_eps == 0.0049
    assert cur.operating_cash_flow == 154_865_050
    assert cur.capex == -27_852_349
    assert cur.dividends_paid == -74_838_811
    assert prior.revenue == 550_450_870

    # Cross-file: prior di laporan 2024 == current di laporan 2023
    if os.path.exists(POWR_2024):
        r2024 = parse_financial_xlsx(POWR_2024)
        cur24 = r2024.periods[0]
        assert cur24.period_label == "2024-12-31"
        assert r2024.entity_name == "PT Cikarang Listrindo Tbk."
        assert r2024.parse_warnings == [], f"Warning 2024: {r2024.parse_warnings}"
        prior24 = r2024.periods[1]
        assert prior24.total_assets == cur.total_assets
        assert prior24.revenue == cur.revenue
        assert prior24.net_income_parent == cur.net_income_parent
        assert prior24.operating_cash_flow == cur.operating_cash_flow

        # Merge lintas file -> 3 periode unik terurut
        from app.utils.ratios import merge_periods
        merged = merge_periods(r2023.periods + r2024.periods)
        assert [p.period_label for p in merged] == [
            "31 December 2022", "31 December 2023", "31 December 2024"
        ]
    print("PASS: file asli POWR 2023 & 2024 terbaca benar dan konsisten lintas file")


if __name__ == "__main__":
    test_cell_to_text_renders_dates_and_floats()
    test_rows_from_sheet_drops_note_reference_numbers()
    test_xlsx_end_to_end_idx_like_workbook()
    test_xlsx_cash_flow_variant_labels()
    test_xlsx_prior_period_filled()
    test_xlsx_real_idx_xbrl_code_sheet_format()
    test_real_powr_files()
    print("\nSemua test XLSX lolos.")
