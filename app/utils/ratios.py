"""
Kalkulasi rasio finansial dan valuasi -- menggabungkan data laporan
keuangan (FinancialPeriod) dan harga saham (PriceSeries).

Semua fungsi return None kalau input yang dibutuhkan None, BUKAN 0 atau
exception -- supaya dashboard bisa nampilin "data tidak tersedia" dengan
jujur daripada angka yang salah/menyesatkan.
"""

from app.models.financial import FinancialPeriod
from app.models.price import PriceSeries
from app.utils.period_dates import parse_period_label_date
from datetime import date
import re


# Benchmark umum untuk konteks "angka ini bagus atau nggak" -- ini PATOKAN
# UMUM lintas industri, bukan angka resmi/otoritatif. Rasio yang wajar
# SANGAT bergantung sektor (bank vs telko vs consumer goods punya struktur
# neraca yang beda jauh, jadi DER "sehat" untuk bank bisa jauh lebih tinggi
# daripada untuk manufaktur). Field ini dikirim ke frontend sebagai
# referensi kasar, BUKAN rekomendasi keputusan -- selalu dibandingkan juga
# dengan rata-rata sektornya sendiri, bukan angka absolut ini saja.
RATIO_BENCHMARKS = {
    "revenue_growth_pct": {"good": 10, "ok": 3, "note": "Pertumbuhan >10%/tahun kuat; di bawah 3% tergolong lambat untuk sebagian besar sektor."},
    "net_income_growth_pct": {"good": 10, "ok": 0, "note": "Idealnya tumbuh sejalan atau lebih cepat dari revenue."},
    "net_margin_pct": {"good": 15, "ok": 5, "note": "Di atas 15% kuat; di bawah 5% tipis, rawan tertekan biaya."},
    "der": {"good": 1.0, "ok": 2.0, "note": "Di bawah 1x konservatif; di atas 2x tergolong leverage tinggi (wajar untuk sektor capital-intensive seperti telko/infrastruktur, tapi tetap perlu diwaspadai).", "lower_is_better": True},
    "current_ratio": {"good": 1.5, "ok": 1.0, "note": "Di atas 1.5x nyaman untuk likuiditas jangka pendek; di bawah 1x berarti aset lancar tidak cukup menutup liabilitas jangka pendek."},
    "roe_pct": {"good": 15, "ok": 8, "note": "Di atas 15% tergolong efisien memutar modal pemegang saham."},
    "roa_pct": {"good": 8, "ok": 3, "note": "Di atas 8% kuat; ROA rendah wajar untuk perusahaan padat aset/leverage tinggi."},
    "per": {"good": 15, "ok": 20, "note": "PER di bawah 15x relatif murah; di atas 20x mulai mahal -- tapi SANGAT tergantung rata-rata sektor & ekspektasi pertumbuhan.", "lower_is_better": True},
    "pbv": {"good": 1.5, "ok": 3.0, "note": "PBV di bawah 1.5x relatif murah terhadap nilai buku; di atas 3x mulai mahal.", "lower_is_better": True},
    "dividend_yield_pct": {"good": 4, "ok": 2, "note": "Di atas 4%/tahun tergolong menarik untuk strategi income/dividen."},
    "cash_dividend_payout_pct": {"good": 60, "ok": 80, "note": "Dividen kas dibayar dibanding arus kas operasi pada periode yang sama. Di bawah 60% menyisakan kas untuk investasi, utang, dan kebutuhan operasional.", "lower_is_better": True},
}


def get_benchmark_verdict(field: str, value: float | None) -> str | None:
    """
    Return 'good', 'ok', atau 'watch' berdasarkan benchmark kasar di atas.
    None kalau field tidak punya benchmark terdaftar atau value None.
    """
    if value is None or field not in RATIO_BENCHMARKS:
        return None
    b = RATIO_BENCHMARKS[field]
    lower_is_better = b.get("lower_is_better", False)

    if lower_is_better:
        if value <= b["good"]:
            return "good"
        elif value <= b["ok"]:
            return "ok"
        else:
            return "watch"
    else:
        if value >= b["good"]:
            return "good"
        elif value >= b["ok"]:
            return "ok"
        else:
            return "watch"


def pct_growth(current: float | None, prior: float | None) -> float | None:
    if current is None or prior is None or prior == 0:
        return None
    return (current / prior - 1) * 100


def net_margin(period: FinancialPeriod) -> float | None:
    if period.net_income_parent is None or not period.revenue:
        return None
    return period.net_income_parent / period.revenue * 100


def der(period: FinancialPeriod) -> float | None:
    """Debt-to-Equity: Total liabilitas / Total ekuitas."""
    if period.total_liabilities is None or not period.total_equity:
        return None
    return period.total_liabilities / period.total_equity


def debt_to_assets(period: FinancialPeriod) -> float | None:
    if period.total_liabilities is None or not period.total_assets:
        return None
    return period.total_liabilities / period.total_assets * 100


def current_ratio(period: FinancialPeriod) -> float | None:
    if period.total_current_assets is None or not period.total_current_liabilities:
        return None
    return period.total_current_assets / period.total_current_liabilities


def free_cash_flow(period: FinancialPeriod) -> float | None:
    """OCF - Capex. Capex disimpan sebagai angka negatif (uang keluar),
    jadi OCF + capex (bukan OCF - capex)."""
    if period.operating_cash_flow is None or period.capex is None:
        return None
    return period.operating_cash_flow + period.capex


def roe(period: FinancialPeriod, prior_equity: float | None = None) -> float | None:
    """Return on Equity. Kalau prior_equity tersedia, pakai rata-rata
    (current+prior)/2 -- lebih akurat daripada cuma equity akhir tahun."""
    if period.net_income_parent is None or not period.total_equity:
        return None
    if prior_equity:
        avg_equity = (period.total_equity + prior_equity) / 2
        if avg_equity == 0:
            return None
        return period.net_income_parent / avg_equity * 100
    return period.net_income_parent / period.total_equity * 100


def roa(period: FinancialPeriod, prior_assets: float | None = None) -> float | None:
    if period.net_income_parent is None or not period.total_assets:
        return None
    if prior_assets:
        avg_assets = (period.total_assets + prior_assets) / 2
        if avg_assets == 0:
            return None
        return period.net_income_parent / avg_assets * 100
    return period.net_income_parent / period.total_assets * 100


def cash_dividend_payout(period: FinancialPeriod) -> float | None:
    """Dividen kas dibayar / arus kas operasi pada periode yang sama.

    Berbeda dari dividend payout ratio formal (dividen yang dideklarasikan
    untuk laba tahun tertentu / laba tahun tersebut), angka ``dividends_paid``
    di laporan arus kas dapat berasal dari keputusan dividen tahun sebelumnya.
    Karena itu pembaginya memakai arus kas operasi di periode yang sama,
    sehingga metrik ini mengukur kemampuan kas mendanai dividen tanpa
    mencampur dua tahun laba yang berbeda.
    """
    if period.dividends_paid is None or period.operating_cash_flow is None:
        return None
    if period.operating_cash_flow <= 0:
        return None
    return abs(period.dividends_paid) / period.operating_cash_flow * 100


def get_price_on_or_before(prices: PriceSeries, target_date: date) -> float | None:
    """Cari harga penutupan pada tanggal tertentu, atau tanggal terdekat
    SEBELUM itu kalau nggak ada data persis (misal weekend/libur bursa)."""
    candidates = [b for b in prices.bars if b.trade_date <= target_date]
    if not candidates:
        return None
    closest = max(candidates, key=lambda b: b.trade_date)
    return closest.close


def per(price: float | None, eps: float | None) -> float | None:
    if price is None or not eps or eps <= 0:
        return None
    return price / eps


def pbv(price: float | None, book_value_per_share: float | None) -> float | None:
    if price is None or not book_value_per_share or book_value_per_share <= 0:
        return None
    return price / book_value_per_share


def dividend_yield(dividend_per_share: float | None, price: float | None) -> float | None:
    if dividend_per_share is None or not price:
        return None
    return dividend_per_share / price * 100


def price_volatility(prices: PriceSeries, annualize: bool = True) -> float | None:
    """Volatilitas harga (std deviasi return harian), dalam persen.
    Kalau annualize=True, dikali sqrt(252) hari bursa per tahun."""
    closes = [b.close for b in prices.bars]
    if len(closes) < 2:
        return None

    returns = [(closes[i] / closes[i - 1] - 1) for i in range(1, len(closes))]
    n = len(returns)
    mean = sum(returns) / n
    variance = sum((r - mean) ** 2 for r in returns) / n
    daily_std = variance ** 0.5

    if annualize:
        daily_std *= 252 ** 0.5

    return daily_std * 100


def max_drawdown(prices: PriceSeries) -> float | None:
    """Penurunan terbesar dari puncak ke lembah dalam periode data, persen."""
    closes = [b.close for b in prices.bars]
    if len(closes) < 2:
        return None

    peak = closes[0]
    max_dd = 0.0
    for price in closes:
        if price > peak:
            peak = price
        drawdown = (price - peak) / peak * 100
        if drawdown < max_dd:
            max_dd = drawdown
    return max_dd


def merge_periods(all_periods: list[FinancialPeriod]) -> list[FinancialPeriod]:
    """
    Gabungkan periode dari beberapa file PDF (masing-masing 2 periode:
    current + prior) jadi satu deret waktu tanpa duplikat.

    Kenapa perlu dedup: laporan tahunan format IDX selalu menyandingkan 2
    tahun (current + prior). Kalau user upload 3 laporan (2023, 2024, 2025),
    maka tahun 2024 akan muncul 2x -- sebagai "current" di laporan 2024 DAN
    sebagai "prior" di laporan 2025. Dedup berdasarkan `period_label` (mis.
    "31 December 2024"), dan kalau ada duplikat, pilih versi dari laporan
    yang paling relevan (current lebih diutamakan drpd prior, karena baris
    labelnya biasanya lebih lengkap/pasti benar dibanding baris comparative
    tahun sebelumnya).

    Return: list periode terurut dari yang paling LAMA ke paling BARU
    (kebalikan dari urutan mentah tiap file, yang selalu current-lalu-prior).
    """
    def _parse_period_date(label: str) -> date | None:
        """Parse format tanggal yang dipakai parser menjadi tanggal kalender.

        Laporan IDX tidak konsisten: satu file dapat memakai ``31 December
        2023``, file lain memakai ``December 31, 2023``, dan ekspor Excel
        memakai nama bulan Indonesia (``31 Desember 2023``). Semuanya harus
        dianggap satu periode. Implementasi di-delegasikan ke util bersama
        parse_period_label_date (app/utils/period_dates.py) supaya aturan
        formatnya satu sumber kebenaran dengan parser lapkeu -- jangan
        duplikat daftar format di sini.
        """
        return parse_period_label_date(label)

    def _period_key(period: FinancialPeriod) -> str:
        # Gunakan tanggal terkanonis agar format DMY dan MDY tidak menjadi
        # dua baris/grafik terpisah. Fallback tetap membuat label aneh aman.
        parsed_date = _parse_period_date(period.period_label)
        if parsed_date:
            return parsed_date.isoformat()
        return re.sub(r"\s+", " ", period.period_label.strip()).casefold()

    def _canonical_period_label(period: FinancialPeriod) -> str:
        """Satu format tampilan untuk tabel dan seluruh grafik dashboard."""
        parsed_date = _parse_period_date(period.period_label)
        return parsed_date.strftime("%d %B %Y") if parsed_date else period.period_label

    seen: dict[str, FinancialPeriod] = {}
    # Prioritas: iterasi tiap file, current period (index genap dalam
    # kelompoknya) menang atas prior period kalau label sama.
    for period in all_periods:
        key = _period_key(period)
        if key not in seen:
            seen[key] = period
        else:
            existing = seen[key]
            # Laporan interim sering membawa angka pembanding periode lalu
            # dengan label neraca 31 Desember. Jika laporan tahunan asli
            # untuk tanggal sama juga di-upload, data tahunan harus menang;
            # jangan sampai angka enam-bulan menggantikan 12-bulan hanya
            # karena jumlah field yang berhasil diekstrak kebetulan sama.
            if existing.is_interim and not period.is_interim:
                seen[key] = period
                continue
            if not existing.is_interim and period.is_interim:
                continue

            # Kalau period baru ini punya lebih banyak field terisi
            # (non-None) daripada yang sudah tersimpan, pakai yang baru --
            # heuristik sederhana untuk pilih versi paling lengkap.
            existing_filled = sum(
                1 for v in existing.model_dump(exclude={"raw_matches"}).values() if v is not None
            )
            new_filled = sum(
                1 for v in period.model_dump(exclude={"raw_matches"}).values() if v is not None
            )
            if new_filled > existing_filled:
                seen[key] = period

    # Urutkan berdasarkan tanggal yang ter-parse dari period_label. Kalau
    # gagal parse (label tidak dikenali), taruh di akhir sebagai fallback
    # daripada bikin whole sort gagal.
    def _sort_key(p: FinancialPeriod):
        # date.max menjaga label tak dikenal di bagian akhir tanpa gagal sort.
        return _parse_period_date(p.period_label) or date.max

    merged_periods = sorted(seen.values(), key=_sort_key)
    # Normalisasi dilakukan setelah deduplikasi: format asli tetap cukup untuk
    # mengenali sumber yang sama, sementara output dashboard selalu konsisten.
    for period in merged_periods:
        period.period_label = _canonical_period_label(period)
    return merged_periods


def build_multi_year_summary(periods: list[FinancialPeriod]) -> list[dict]:
    """
    Hitung rasio year-over-year untuk deret waktu yang sudah digabung (lihat
    merge_periods). Setiap entry menyertakan `is_interim` dan
    `coverage_months` dari parser (lihat _detect_interim_info di
    financial_pdf.py) supaya frontend tahu periode mana yang laporan
    tahunan penuh vs interim/kuartalan.

    PENTING: growth rate (revenue_growth_pct, dst) HANYA dihitung antar
    periode TAHUNAN PENUH (is_interim=False) yang berurutan -- bukan
    terhadap periode sebelumnya apa adanya. Ini mencegah perbandingan
    12-bulan vs 6-bulan yang salah kalau user mencampur laporan tahunan
    dengan laporan kuartalan berjalan (mis. laporan Q2 2026 yang baru
    mencakup Januari-Juni, dibandingkan ke laporan tahunan 2025 yang
    mencakup Januari-Desember penuh).

    Periode interim tetap ditampilkan (dengan angka mentahnya dan tanggal
    jelas), tapi growth_pct-nya None -- frontend menampilkan ini sebagai
    info kecil terpisah, bukan bagian dari rantai pertumbuhan tahunan utama.
    """
    results = []
    last_annual_period = None

    for period in periods:
        entry = {
            "period_label": period.period_label,
            "is_interim": period.is_interim,
            "coverage_months": period.coverage_months,
            "revenue": period.revenue,
            "net_income_parent": period.net_income_parent,
            "total_assets": period.total_assets,
            "total_liabilities": period.total_liabilities,
            "total_equity": period.total_equity,
            "operating_cash_flow": period.operating_cash_flow,
            "eps": period.basic_eps,
            "net_margin_pct": net_margin(period),
            "der": der(period),
            "current_ratio": current_ratio(period),
            "free_cash_flow": free_cash_flow(period),
        }

        if not period.is_interim and last_annual_period is not None:
            entry["revenue_growth_pct"] = pct_growth(period.revenue, last_annual_period.revenue)
            entry["net_income_growth_pct"] = pct_growth(period.net_income_parent, last_annual_period.net_income_parent)
            entry["eps_growth_pct"] = pct_growth(period.basic_eps, last_annual_period.basic_eps)
        else:
            # Periode pertama dalam rantai tahunan (belum ada tahunan
            # sebelumnya untuk dibandingkan), ATAU periode interim/kuartalan
            # (sengaja tidak dibandingkan ke periode tahunan penuh).
            entry["revenue_growth_pct"] = None
            entry["net_income_growth_pct"] = None
            entry["eps_growth_pct"] = None

        if not period.is_interim:
            last_annual_period = period

        results.append(entry)
    return results


def build_ratio_summary(
    current: FinancialPeriod,
    prior: FinancialPeriod,
    prices: PriceSeries | None = None,
    shares_outstanding: float | None = None,
    exchange_rate_to_idr: float | None = None,
) -> dict:
    """
    Gabungkan semua rasio jadi satu dict ringkas untuk ditampilkan di
    dashboard. `shares_outstanding` opsional -- kalau user tahu jumlah
    saham beredar (dalam LEMBAR PENUH, bukan jutaan/miliar lembar), PER/PBV
    bisa dihitung; kalau tidak, field itu None dan dashboard menampilkan
    "tidak tersedia" alih-alih menebak.

    PENTING soal satuan: semua angka di FinancialPeriod (total_equity,
    dividends_paid, dst) berasal dari laporan keuangan yang disajikan
    dalam satuan yang BERBEDA-BEDA antar emiten -- bisa jutaan Rupiah
    ATAU miliaran Rupiah, tergantung skala perusahaan (dideteksi otomatis
    dari metadata "Pembulatan yang digunakan" di setiap file, disimpan di
    `current.reporting_unit_multiplier`). Sementara `shares_outstanding`
    yang diinput user adalah jumlah LEMBAR SAHAM PENUH (misal
    32.250.000.000 lembar, bukan "32.250 juta lembar"). Sebelum dibagi,
    total_equity/dividends_paid HARUS dikonversi dulu ke satuan Rupiah
    penuh (dikali reporting_unit_multiplier -- BUKAN hardcode 1 juta)
    supaya hasilnya sebanding dengan harga saham per lembar yang juga
    dalam Rupiah penuh.

    Riwayat bug nyata terkait ini (dua lapis, keduanya sudah diperbaiki):
    1. Awalnya BVPS/DPS tidak dikonversi sama sekali -> PBV muncul 1,7
       juta kali lipat dari seharusnya.
    2. Setelah #1 diperbaiki dengan hardcode "kali 1 juta", ternyata tidak
       semua emiten pakai satuan jutaan -- ASII pakai satuan MILIAR
       (terbukti dari metadata "Pembulatan yang digunakan: Miliaran / In
       Billion" di dalam file PDF-nya sendiri). Hardcode 1 juta membuat
       BVPS ASII salah 1000x lebih kecil, dan PBV yang dihitung dari situ
       muncul 683x (seharusnya ~1.2x). Sekarang faktor konversi diambil
       dari `reporting_unit_multiplier` yang dideteksi otomatis per file,
       bukan diasumsikan seragam.
    """
    unit_multiplier = current.reporting_unit_multiplier

    # --- Konversi mata uang pelaporan -> Rupiah ---
    # Sebagian kecil emiten IDX menyajikan lapkeu dalam mata uang asing
    # (mis. POWR ber-USD -- bug nyata: PER-nya muncul 222.826x karena EPS
    # USD dibagi-dibandingkan langsung dengan harga saham Rupiah). Semua
    # angka yang dibandingkan dengan HARGA SAHAM (per-share: EPS, BVPS,
    # DPS) harus dikonversi dulu dengan kurs; angka agregat (revenue,
    # equity, dst) tidak dibandingkan dengan harga sehingga boleh tetap
    # dalam mata uang asal.
    reporting_currency = (current.reporting_currency or "IDR").upper()
    is_foreign = reporting_currency not in ("IDR",)
    currency_rate_missing = is_foreign and not exchange_rate_to_idr
    fx = exchange_rate_to_idr if (is_foreign and exchange_rate_to_idr) else 1.0

    def to_idr_per_share(value: float | None) -> float | None:
        """Konversi nilai per-saham (sudah dalam satuan penuh mata uang
        laporan) ke Rupiah. None bila kurs asing belum tersedia."""
        if value is None:
            return None
        if currency_rate_missing:
            return None
        return value * fx

    summary = {
        "reporting_currency": reporting_currency,
        "exchange_rate_applied": fx if is_foreign else None,
        "currency_rate_missing": currency_rate_missing,
        "revenue_growth_pct": pct_growth(current.revenue, prior.revenue),
        "net_income_growth_pct": pct_growth(current.net_income_parent, prior.net_income_parent),
        "net_margin_pct": net_margin(current),
        "net_margin_prior_pct": net_margin(prior),
        "der": der(current),
        "der_prior": der(prior),
        "debt_to_assets_pct": debt_to_assets(current),
        "current_ratio": current_ratio(current),
        "current_ratio_prior": current_ratio(prior),
        "operating_cash_flow": current.operating_cash_flow,
        "ocf_growth_pct": pct_growth(current.operating_cash_flow, prior.operating_cash_flow),
        "free_cash_flow": free_cash_flow(current),
        "free_cash_flow_prior": free_cash_flow(prior),
        "roe_pct": roe(current, prior.total_equity),
        "roa_pct": roa(current, prior.total_assets),
        "eps": current.basic_eps,
        "eps_growth_pct": pct_growth(current.basic_eps, prior.basic_eps),
        "cash_dividend_payout_pct": cash_dividend_payout(current),
        "dividends_paid": current.dividends_paid,
    }

    # EPS dalam Rupiah dihitung begitu kurs tersedia -- TIDAK menunggu data
    # harga. Kartu EPS di frontend menampilkan konteks konversi ini (mis.
    # "USD 0.0048 · kurs 16.250 -> ~ Rp78") walau user tidak upload CSV
    # harga; sebelumnya konversi baru muncul kalau ada CSV harga sehingga
    # konteks kurs hilang di kasus umum.
    if current.basic_eps is not None:
        summary["eps_in_idr"] = to_idr_per_share(current.basic_eps)

    if prices and prices.bars:
        latest_price = prices.bars[-1].close
        summary["latest_price"] = latest_price
        summary["latest_price_date"] = str(prices.bars[-1].trade_date)
        summary["price_volatility_annualized_pct"] = price_volatility(prices)
        summary["max_drawdown_pct"] = max_drawdown(prices)

        if current.basic_eps:
            # PER memakai EPS yang sudah dikonversi ke Rupiah (to_idr_per_share
            # mengembalikan None bila kurs asing belum diisi -- PER juga None).
            summary["per"] = per(latest_price, summary.get("eps_in_idr"))

        if shares_outstanding:
            # PENTING: BVPS/PBV HARUS pakai equity yang diatribusikan ke
            # pemilik entitas INDUK saja (total_equity_parent), BUKAN
            # total_equity biasa (yang termasuk kepentingan non-pengendali
            # / NCI di anak-anak perusahaan). Untuk holding company besar
            # dengan banyak anak perusahaan yang tidak dimiliki 100% (mis.
            # ASII), NCI bisa jadi porsi signifikan dari total equity --
            # memakai total_equity biasa untuk BVPS akan menghasilkan PBV
            # yang jauh lebih kecil dari kenyataan (karena mengasumsikan
            # ekuitas NCI juga "milik" pemegang saham publik, padahal
            # bukan). Ini bug nyata yang pernah ditemukan: PBV ASII
            # sempat muncul 683x (seharusnya ~1.2x) karena salah pakai
            # total_equity yang termasuk NCI.
            #
            # Fallback ke total_equity biasa HANYA kalau total_equity_parent
            # gagal ke-parse (None) -- lebih baik kasih angka approximate
            # (yang mungkin sedikit meleset untuk emiten dengan NCI besar)
            # daripada tidak menampilkan PBV sama sekali.
            equity_for_bvps = current.total_equity_parent if current.total_equity_parent is not None else current.total_equity
            if equity_for_bvps:
                equity_full_rupiah = equity_for_bvps * unit_multiplier
                if is_foreign and not currency_rate_missing:
                    equity_full_rupiah *= fx
                if currency_rate_missing:
                    summary["book_value_per_share"] = None
                    summary["pbv"] = None
                else:
                    bvps = equity_full_rupiah / shares_outstanding
                    summary["book_value_per_share"] = bvps
                    summary["pbv"] = pbv(latest_price, bvps)
                summary["bvps_used_parent_equity"] = current.total_equity_parent is not None

        if shares_outstanding and current.dividends_paid:
            dividends_full_rupiah = abs(current.dividends_paid) * unit_multiplier
            if is_foreign and not currency_rate_missing:
                dividends_full_rupiah *= fx
            if currency_rate_missing:
                summary["dividend_per_share"] = None
                summary["dividend_yield_pct"] = None
            else:
                dps = dividends_full_rupiah / shares_outstanding
                summary["dividend_per_share"] = dps
                summary["dividend_yield_pct"] = dividend_yield(dps, latest_price)

    # Tempelkan verdict benchmark ('good'/'ok'/'watch') untuk tiap field
    # yang punya benchmark terdaftar. Dikirim sebagai field terpisah
    # "<nama>_verdict" supaya frontend gampang render badge/warna tanpa
    # perlu logic benchmark lagi di JS.
    for field in list(summary.keys()):
        if field in RATIO_BENCHMARKS:
            summary[f"{field}_verdict"] = get_benchmark_verdict(field, summary[field])

    return summary
