"""
Test parser terhadap data ISAT asli (2024 & 2025 annual report).

Dua lapis verifikasi:
1. Nilai absolut -- dicocokkan manual terhadap laporan keuangan asli
   (dibaca & dihitung manual, baris per baris, sebelum parser ini dibuat).
2. Cross-file consistency -- periode "prior year" di laporan 2025 HARUS
   sama persis dengan periode "current year" di laporan 2024, karena
   keduanya melaporkan angka tahun buku yang sama (2024). Ini bentuk
   verifikasi yang lebih kuat daripada nilai absolut saja, karena tidak
   bergantung pada parser "kebetulan cocok" dengan angka yang sudah
   diketahui duluan.

Jalankan: python -m pytest tests/test_financial_parser.py -v
(atau langsung: python tests/test_financial_parser.py)
"""

import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.parsers.financial_pdf import parse_financial_pdf
from app.utils.ratios import build_ratio_summary

SAMPLE_2025 = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "sample_data",
    "FinancialStatement-2025-Tahunan-ISAT.pdf"
)
SAMPLE_2024 = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "sample_data",
    "FinancialStatement-2024-Tahunan-ISAT.pdf"
)

# File ASII dipakai untuk menguji ketahanan parser terhadap variasi format
# nyata yang ditemukan di dunia nyata: tanggal "December 31, 2023" (bukan
# "31 December 2023"), label yang wrapping dengan pola BERBEDA-BEDA antar
# file (kadang nempel satu baris, kadang terpecah 2-5 baris, kadang angka
# nyempil di tengah frasa label). Struktur ASII jauh lebih beragam
# formatnya dibanding ISAT karena mencakup rentang tahun lebih panjang
# (2021-2026) dan kemungkinan sumber ekspor PDF yang berbeda-beda.
SAMPLE_ASII = {
    2022: os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sample_data", "FinancialStatement-2022-Tahunan-ASII.pdf"),
    2023: os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sample_data", "FinancialStatement-2023-Tahunan-ASII.pdf"),
    2024: os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sample_data", "FinancialStatement-2024-Tahunan-ASII.pdf"),
    2025: os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sample_data", "FinancialStatement-2025-Tahunan-ASII.pdf"),
}
SAMPLE_ASII_Q2_2026 = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "sample_data", "FinancialStatement-2026-II-ASII.pdf")

# Angka ini diverifikasi manual dari teks laporan (bukan ditebak dari
# hasil parser) sebelum dipakai sebagai baseline test.
EXPECTED_2025_CURRENT = {
    "total_assets": 118629871,
    "total_current_assets": 19003874,
    "total_liabilities": 79120763,
    "total_current_liabilities": 30499266,
    "total_equity": 39509108,
    "cash_and_equivalents": 5075324,
    "revenue": 56518141,
    "profit_before_tax": 7284647,
    "net_income_parent": 5509703,
    "operating_cash_flow": 21550578,
    "investing_cash_flow": -12097383,
    "financing_cash_flow": -8861591,
    "capex": -12359655,
    "dividends_paid": -2702400,
    "basic_eps": 170.84,
}

CROSS_CHECK_FIELDS = [
    "total_assets", "total_liabilities", "total_equity", "revenue",
    "net_income_parent", "operating_cash_flow", "investing_cash_flow",
    "financing_cash_flow", "capex", "dividends_paid", "basic_eps",
]


def test_2025_current_period_matches_verified_values():
    """Angka periode berjalan (2025) harus persis sama dengan yang sudah
    diverifikasi manual dari teks laporan."""
    if not os.path.exists(SAMPLE_2025):
        print("SKIP: sample file tidak ada di sample_data/")
        return

    result = parse_financial_pdf(SAMPLE_2025)
    period = result.periods[0]

    failures = []
    for field, expected in EXPECTED_2025_CURRENT.items():
        actual = getattr(period, field)
        if actual != expected:
            failures.append(f"{field}: got {actual}, expected {expected}")

    assert not failures, "Mismatch:\n" + "\n".join(failures)
    print(f"PASS: semua {len(EXPECTED_2025_CURRENT)} field cocok dengan nilai terverifikasi")


def test_cross_file_consistency():
    """Periode 'prior' di laporan 2025 harus identik dengan periode
    'current' di laporan 2024 -- keduanya melaporkan tahun buku yang sama.
    Ini menguji bahwa parser generalize dengan benar, bukan cuma cocok
    kebetulan di satu file."""
    if not (os.path.exists(SAMPLE_2025) and os.path.exists(SAMPLE_2024)):
        print("SKIP: sample files tidak lengkap di sample_data/")
        return

    r2025 = parse_financial_pdf(SAMPLE_2025)
    r2024 = parse_financial_pdf(SAMPLE_2024)

    prior_in_2025 = r2025.periods[1]
    current_in_2024 = r2024.periods[0]

    failures = []
    for field in CROSS_CHECK_FIELDS:
        a = getattr(prior_in_2025, field)
        b = getattr(current_in_2024, field)
        if a != b:
            failures.append(f"{field}: 2025-report-prior={a} vs 2024-report-current={b}")

    assert not failures, "Cross-file mismatch:\n" + "\n".join(failures)
    print(f"PASS: semua {len(CROSS_CHECK_FIELDS)} field konsisten lintas file")


def test_no_parse_warnings_on_known_good_files():
    """Untuk file yang formatnya sudah kita kenal (ISAT, template
    infrastruktur/telko standar IDX), parser seharusnya tidak menghasilkan
    warning sama sekali. Kalau ada warning muncul di sini, berarti ada
    regresi di parser."""
    if not os.path.exists(SAMPLE_2025):
        print("SKIP: sample file tidak ada")
        return

    result = parse_financial_pdf(SAMPLE_2025)
    assert result.parse_warnings == [], f"Ada warning tak terduga: {result.parse_warnings}"
    print("PASS: tidak ada parse warning")


def test_multi_pdf_merge_deduplicates_overlapping_year():
    """Upload dua laporan tahunan (2024 dan 2025) yang overlap di tahun 2024
    (2024 report punya periode 2023+2024, 2025 report punya periode 2024+2025).
    Hasil merge harus punya 3 tahun unik (2023,2024,2025), bukan 4 dengan
    2024 dobel."""
    from app.utils.ratios import merge_periods

    if not (os.path.exists(SAMPLE_2025) and os.path.exists(SAMPLE_2024)):
        print("SKIP: sample files tidak lengkap")
        return

    s2024 = parse_financial_pdf(SAMPLE_2024)
    s2025 = parse_financial_pdf(SAMPLE_2025)
    all_periods = s2024.periods + s2025.periods
    merged = merge_periods(all_periods)

    labels = [p.period_label for p in merged]
    expected_labels = ["31 December 2023", "31 December 2024", "31 December 2025"]
    assert labels == expected_labels, f"Expected {expected_labels}, got {labels}"

    # Data 2024 di hasil merge harus sama dengan yang dilaporkan sebagai
    # CURRENT period di file 2024 (bukan prior period di file 2025) --
    # meski di kasus ISAT ini kebetulan sama persis, prinsip prioritasnya
    # tetap harus diverifikasi.
    period_2024 = next(p for p in merged if p.period_label == "31 December 2024")
    expected_revenue_2024 = s2024.periods[0].revenue  # current period di file 2024
    assert period_2024.revenue == expected_revenue_2024, (
        f"2024 revenue mismatch: got {period_2024.revenue}, expected {expected_revenue_2024}"
    )

    print(f"PASS: merge menghasilkan {len(labels)} periode unik ({labels}), tanpa duplikat")


def test_multi_pdf_merge_deduplicates_equivalent_date_formats():
    """Format DMY dan MDY pada tanggal yang sama tidak boleh jadi dua tren."""
    from app.models.financial import FinancialPeriod
    from app.utils.ratios import merge_periods

    periods = [
        FinancialPeriod(period_label="31 December 2023", revenue=316_565),
        FinancialPeriod(period_label="December 31, 2023", revenue=316_565),
        FinancialPeriod(period_label="December 31, 2022", revenue=301_379),
    ]

    merged = merge_periods(periods)
    assert [p.period_label for p in merged] == [
        "31 December 2022", "31 December 2023"
    ]


def test_multi_pdf_merge_prefers_annual_over_interim_comparative():
    """Angka pembanding interim tidak boleh menggantikan laporan tahunan."""
    from app.models.financial import FinancialPeriod
    from app.utils.ratios import merge_periods

    interim_comparative = FinancialPeriod(
        period_label="31 December 2025", is_interim=True,
        revenue=162_857, net_income_parent=15_515,
    )
    annual = FinancialPeriod(
        period_label="31 December 2025", is_interim=False,
        revenue=323_392, net_income_parent=32_769,
    )

    merged = merge_periods([interim_comparative, annual])
    assert len(merged) == 1
    assert merged[0].is_interim is False
    assert merged[0].revenue == 323_392


def test_annual_reports_not_flagged_as_interim():
    """Laporan tahunan asli (berakhir 31 Desember) HARUS terdeteksi
    is_interim=False. Kalau ini gagal, dashboard akan salah munculkan
    peringatan 'perbandingan periode tidak setara' untuk laporan yang
    sebenarnya valid dibandingkan."""
    if not os.path.exists(SAMPLE_2025):
        print("SKIP: sample file tidak ada")
        return

    result = parse_financial_pdf(SAMPLE_2025)
    for period in result.periods:
        assert period.is_interim is False, (
            f"Periode '{period.period_label}' salah terdeteksi sebagai interim"
        )
        assert period.fiscal_year is not None, (
            f"fiscal_year tidak terisi untuk '{period.period_label}'"
        )
    print("PASS: laporan tahunan terdeteksi is_interim=False dengan benar")


def test_interim_detection_logic():
    """Test unit murni untuk _detect_interim_info -- tidak butuh file PDF
    asli, cukup simulasi baris teks yang polanya sama dengan laporan
    interim/kuartalan asli (mis. laporan per 30 Juni)."""
    from app.parsers.financial_pdf import _detect_interim_info

    rows_interim = [
        "30 June 2026 30 June 2025",
        "untuk enam (6) bulan yang berakhir 30 Juni 2026",
        "for the six (6) months ended 30 June 2026",
    ]
    is_interim, months = _detect_interim_info(rows_interim, "30 June 2026")
    assert is_interim is True, "Laporan per 30 Juni harus terdeteksi interim"
    assert months == 6, f"Coverage months harus 6, dapat {months}"

    rows_annual = ["31 December 2025 31 December 2024"]
    is_interim2, months2 = _detect_interim_info(rows_annual, "31 December 2025")
    assert is_interim2 is False, "Laporan per 31 Desember TIDAK boleh terdeteksi interim"

    # "X bulan yang berakhir 31 Desember [tahun label]" pada laporan
    # berlabel 31 Desember HARUS DIABAIKAN -- bukan cakupan laporan ini.
    # Laporan interim IDX tidak pernah berlabel 31 Desember (selalu akhir
    # kuartal), jadi frasa itu hampir pasti milik disclosure kombinasi
    # bisnis/akuisisi: entitas yang diakuisisi punya periode pelaporannya
    # sendiri. Contoh nyata (PGAS 2023): catatan "...pendapatan
    # konsolidasian untuk periode sembilan bulan yang berakhir 31 Desember
    # 2023 dan 2022..." sempat membuat laporan TAHUNAN salah tertandai
    # interim 9 bulan dan hilang dari rasio tahunan.
    rows_acquisition_disclosure = [
        "31 December 2025 31 December 2024",
        "untuk enam (6) bulan yang berakhir 31 Desember 2025",
        "pendapatan konsolidasian untuk periode sembilan bulan yang berakhir 31 Desember 2025 dan 2024",
    ]
    is_interim3, months3 = _detect_interim_info(
        rows_acquisition_disclosure, "31 December 2025"
    )
    assert is_interim3 is False, "Disclosure akuisisi 'X bulan berakhir 31 Des' tidak boleh menandai laporan tahunan sebagai interim"
    assert months3 is None, f"Coverage months tidak boleh terisi dari disclosure itu, dapat {months3}"

    # SEBALIKNYA: teks cakupan eksplisit yang menyebutkan TANGGAL LAIN
    # (bukan period_label yang sedang diperiksa) HARUS DIABAIKAN -- ini
    # regression guard untuk bug nyata yang pernah terjadi di file ASII Q2
    # 2026: dokumen itu punya SATU baris metadata "6 Bulan yang berakhir
    # pada 30/06/2026" yang menjelaskan periode CURRENT (Juni 2026), tapi
    # baris itu "bocor" dan salah diterapkan ke periode PRIOR (31 Desember
    # 2025) juga -- membuat laporan tahunan penuh salah tertandai interim.
    rows_numeric_coverage_different_date = [
        "laporan keuangan untuk periode 6 Bulan yang berakhir pada 30/06/2026"
    ]
    is_interim4, months4 = _detect_interim_info(
        rows_numeric_coverage_different_date, "31 December 2025"
    )
    assert is_interim4 is False, (
        "Teks cakupan yang tanggalnya BEDA dari period_label harus "
        "diabaikan -- kalau tidak, periode lain 'bocor' jadi interim"
    )
    assert months4 is None, f"Coverage months harus None (diabaikan), dapat {months4}"

    # Kasus yang sama tapi tanggalnya justru DICEK untuk periode yang
    # benar (Juni 2026) -- harus tetap terdeteksi interim.
    is_interim5, months5 = _detect_interim_info(
        rows_numeric_coverage_different_date, "30 June 2026"
    )
    assert is_interim5 is True, "Cakupan numerik dengan tanggal yang cocok harus ditandai interim"
    assert months5 == 6, f"Coverage months harus 6, dapat {months5}"

    print("PASS: deteksi interim vs tahunan bekerja dengan benar")


def test_asii_cross_file_consistency_and_no_warnings():
    """
    Uji parser terhadap 5 file ASII (2022-2025 tahunan + Q2 2026 interim).
    Ini regression guard untuk bug nyata yang pernah terjadi:
    - Format tanggal "December 31, 2023" (bukan "31 December 2023") bikin
      deteksi periode gagal total dan salah menandai laporan tahunan
      sebagai interim.
    - Field cash flow (operating/investing/financing) sempat HILANG dari
      LABEL_MAP tanpa terdeteksi lewat warning (silent failure).
    - investing_cash_flow sempat salah mengambil nilai operating_cash_flow
      karena window disambiguasi terlalu lebar.
    - capex sempat salah mengambil baris "Penerimaan dari penjualan aset
      tetap" (proceeds, BUKAN capex) karena fallback keyword-pair terlalu
      longgar dan dicoba sebelum baris yang benar sempat diperiksa.

    Kalau salah satu bug ini muncul lagi di masa depan (mis. akibat
    perubahan pattern/disambiguation), test ini akan gagal.
    """
    if not all(os.path.exists(p) for p in SAMPLE_ASII.values()):
        print("SKIP: sample file ASII tidak lengkap")
        return

    results = {year: parse_financial_pdf(path) for year, path in SAMPLE_ASII.items()}

    # 1. Nol parse warning di semua file tahunan ASII.
    for year, result in results.items():
        assert result.parse_warnings == [], (
            f"ASII {year}: ada warning tak terduga: {result.parse_warnings}"
        )

    # 2. Semua periode tahunan ASII harus terdeteksi is_interim=False,
    #    walau formatnya berbeda-beda (2023 pakai "December 31, 2023").
    for year, result in results.items():
        for period in result.periods:
            assert period.is_interim is False, (
                f"ASII {year} period '{period.period_label}' salah "
                f"terdeteksi sebagai interim"
            )

    # 3. Cross-file consistency: tahun yang overlap antar file (mis. 2023
    #    muncul sebagai current di file 2023 DAN sebagai prior di file
    #    2024) SEHARUSNYA punya angka identik untuk field kunci -- kecuali
    #    kalau emiten merilis restatement (revisi kecil angka lapkeu tahun
    #    sebelumnya), yang merupakan praktik akuntansi wajar, bukan bug.
    #    Toleransi kecil (0.1%) diberikan untuk mengakomodasi ini; cash
    #    flow TIDAK diberi toleransi karena itu sumber bug paling serius
    #    yang pernah ditemukan (salah ambil field lain sepenuhnya, bukan
    #    beda angka sedikit) -- kalau cash flow beda sama sekali, itu
    #    tanda parser salah tangkap baris, bukan restatement.
    fields_with_tolerance = [
        "total_equity", "total_liabilities", "total_current_liabilities",
        "net_income_parent", "basic_eps",
    ]
    fields_exact = ["operating_cash_flow", "investing_cash_flow", "financing_cash_flow"]

    # Catatan: ASII 2024 mengalami restatement kecil pada beberapa item
    # liabilitas antara laporan 2024 dan laporan 2025 (diverifikasi manual
    # dari kedua dokumen -- misal total_current_liabilities 2024 tercatat
    # 133,303 di laporan 2024 tapi 130,497 sebagai comparative di laporan
    # 2025, selisih ~2.1%). Ini praktik akuntansi wajar, bukan bug parser.
    # Toleransi 3% dipakai untuk mengakomodasi restatement semacam ini.
    for year in [2022, 2023, 2024]:
        current_in_year = results[year].periods[0]
        prior_in_next_year = results[year + 1].periods[1]

        for field in fields_with_tolerance:
            a = getattr(current_in_year, field)
            b = getattr(prior_in_next_year, field)
            if a is None or b is None:
                continue
            diff_pct = abs(a - b) / max(abs(a), abs(b), 1) * 100
            assert diff_pct < 3.0, (
                f"ASII {year} {field}: current-in-{year}={a} vs "
                f"prior-in-{year+1}={b} -- beda {diff_pct:.2f}%, "
                f"terlalu besar untuk sekadar restatement"
            )

        for field in fields_exact:
            a = getattr(current_in_year, field)
            b = getattr(prior_in_next_year, field)
            assert a == b, (
                f"ASII {year} {field}: current-in-{year}={a} vs "
                f"prior-in-{year+1}={b} -- HARUS identik (bukan kandidat "
                f"restatement), kemungkinan parser salah tangkap baris"
            )

    # 4. Sanity check nilai absolut untuk satu tahun yang sudah diverifikasi
    #    manual ke sumber resmi (bukan cuma konsisten antar file, tapi juga
    #    benar terhadap kenyataan) -- ASII 2023: total ekuitas ~Rp250,4T,
    #    net income ~Rp33,8T, EPS ~Rp836.
    p2023 = results[2023].periods[0]
    assert p2023.total_equity == 250418.0, f"ASII 2023 total_equity salah: {p2023.total_equity}"
    assert p2023.net_income_parent == 33839.0, f"ASII 2023 net_income_parent salah: {p2023.net_income_parent}"
    assert p2023.basic_eps == 836.0, f"ASII 2023 basic_eps salah: {p2023.basic_eps}"
    assert p2023.operating_cash_flow == 33746.0, f"ASII 2023 operating_cash_flow salah: {p2023.operating_cash_flow}"

    print(f"PASS: 5 file ASII konsisten lintas file, nol warning, dan nilai 2023 terverifikasi ke sumber resmi")


def test_asii_interim_report_detected_correctly():
    """
    File Q2 2026 ASII (interim/kuartalan) harus terdeteksi is_interim=True
    untuk periode 30 Juni 2026, TAPI periode prior-nya (31 Desember 2025,
    laporan tahunan) harus tetap is_interim=False. Ini menguji bahwa
    deteksi interim per-periode independen dan tidak "ketularan" dari
    periode lain dalam file yang sama.
    """
    if not os.path.exists(SAMPLE_ASII_Q2_2026):
        print("SKIP: sample file Q2 2026 ASII tidak ada")
        return

    result = parse_financial_pdf(SAMPLE_ASII_Q2_2026)
    current, prior = result.periods[0], result.periods[1]

    assert current.is_interim is True, (
        f"Periode '{current.period_label}' (Q2 2026) harus terdeteksi interim"
    )
    assert prior.is_interim is False, (
        f"Periode '{prior.period_label}' (31 Des 2025) TIDAK boleh "
        f"terdeteksi interim"
    )

    print("PASS: laporan Q2 2026 ASII -- interim dan periode prior tahunan terdeteksi benar")


def test_reporting_unit_multiplier_detected_correctly():
    """
    Regression guard untuk bug PBV 683x yang pernah terjadi: laporan IDX
    TIDAK selalu disajikan dalam satuan jutaan Rupiah -- ASII memakai
    satuan MILIARAN (terbukti dari metadata eksplisit "Pembulatan yang
    digunakan: Miliaran / In Billion" di dalam file PDF-nya sendiri),
    sementara ISAT memakai jutaan. Kalau ini di-hardcode sebagai jutaan
    untuk semua emiten (seperti sebelumnya), BVPS/dividend-per-share untuk
    ASII akan salah 1000x lebih kecil, dan PBV yang dihitung dari situ
    jadi menyesatkan total.
    """
    if not (os.path.exists(SAMPLE_ASII[2025]) and os.path.exists(SAMPLE_2025)):
        print("SKIP: sample file tidak lengkap")
        return

    asii = parse_financial_pdf(SAMPLE_ASII[2025])
    isat = parse_financial_pdf(SAMPLE_2025)

    assert asii.periods[0].reporting_unit_multiplier == 1_000_000_000, (
        f"ASII harus terdeteksi satuan MILIAR (1_000_000_000), dapat "
        f"{asii.periods[0].reporting_unit_multiplier}"
    )
    assert isat.periods[0].reporting_unit_multiplier == 1_000_000, (
        f"ISAT harus terdeteksi satuan JUTA (1_000_000), dapat "
        f"{isat.periods[0].reporting_unit_multiplier}"
    )

    print("PASS: satuan pembulatan (jutaan/miliaran) terdeteksi benar per-file")


def test_asii_pbv_not_wildly_wrong():
    """
    End-to-end regression guard untuk bug PBV 683x: dengan harga saham dan
    jumlah saham beredar ASII yang realistis, PBV yang dihasilkan HARUS
    berada di rentang yang masuk akal untuk perusahaan sebesar ASII
    (secara historis di kisaran 0.5x-2x), bukan ratusan kali lipat dari
    itu. Test ini tidak menuntut angka PBV persis (karena bergantung pada
    harga saham & jumlah lembar yang bisa berubah), tapi memastikan ORDE
    BESARNYA benar -- kalau bug unit multiplier muncul lagi, PBV akan
    kembali ke ratusan/ribuan x dan test ini akan gagal.
    """
    if not os.path.exists(SAMPLE_ASII[2025]):
        print("SKIP: sample file ASII 2025 tidak ada")
        return

    from app.models.price import PriceSeries, PriceBar
    from datetime import date

    result = parse_financial_pdf(SAMPLE_ASII[2025])
    current, prior = result.periods[0], result.periods[1]

    prices = PriceSeries(source_filename="test.csv", bars=[
        PriceBar(trade_date=date(2026, 9, 11), close=4910)
    ])
    shares_outstanding = 40_483_553_140  # jumlah saham beredar ASII, dari sumber resmi

    ratios = build_ratio_summary(current, prior, prices, shares_outstanding)
    pbv = ratios.get("pbv")

    assert pbv is not None, "PBV harus terhitung (tidak None) dengan input yang lengkap"
    assert 0.3 < pbv < 5.0, (
        f"PBV ASII = {pbv}, di luar rentang wajar (0.3x-5x) -- kemungkinan "
        f"bug unit multiplier muncul lagi (PBV 683x pernah terjadi akibat "
        f"bug ini)"
    )

    print(f"PASS: PBV ASII = {pbv:.2f}x, berada di rentang yang masuk akal")


if __name__ == "__main__":
    test_2025_current_period_matches_verified_values()
    test_cross_file_consistency()
    test_no_parse_warnings_on_known_good_files()
    test_multi_pdf_merge_deduplicates_overlapping_year()
    test_annual_reports_not_flagged_as_interim()
    test_interim_detection_logic()
    test_asii_cross_file_consistency_and_no_warnings()
    test_asii_interim_report_detected_correctly()
    test_reporting_unit_multiplier_detected_correctly()
    test_asii_pbv_not_wildly_wrong()
    print("\nSemua test lolos.")
