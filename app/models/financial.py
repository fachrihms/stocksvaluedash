"""
Struktur data hasil parsing laporan keuangan.

Semua angka dalam SATUAN ASLI SEPERTI TERCANTUM DI LAPORAN -- PENTING:
satuan ini BERBEDA-BEDA antar emiten! Laporan IDX/OJK punya metadata
eksplisit "Pembulatan yang digunakan" (level of rounding) yang bisa berupa
"Jutaan / In Million" ATAU "Miliaran / In Billion" tergantung skala
perusahaan (emiten besar seperti ASII pakai satuan miliar, emiten seperti
ISAT pakai jutaan). Field `reporting_unit_multiplier` di bawah menyimpan
faktor pengali yang benar (1_000_000 untuk jutaan, 1_000_000_000 untuk
miliar) HASIL DETEKSI OTOMATIS dari metadata itu -- JANGAN PERNAH
asumsikan/hardcode satuan tertentu di kode manapun yang mengonsumsi field
model ini, selalu pakai reporting_unit_multiplier ini sebagai sumber
kebenaran. Field yang nggak ketemu di PDF akan None, bukan 0 -- penting
supaya kalkulasi rasio nggak diam-diam salah karena mengira "kosong" itu
"nol".
"""

from pydantic import BaseModel
from typing import Optional


class FinancialPeriod(BaseModel):
    """Satu periode pelaporan (biasanya satu tahun fiskal)."""

    period_label: str  # contoh: "31 December 2025"
    # Nama PDF asal periode ini. Dipertahankan saat periode lintas-file
    # digabung agar dashboard dapat menjelaskan sumber setiap rasio.
    source_filename: Optional[str] = None
    fiscal_year: Optional[int] = None  # contoh: 2025 -- dipakai untuk dedup/sort saat merge multi-file
    entity_code: Optional[str] = None  # contoh: "ISAT"
    # Faktor pengali untuk mengonversi angka mentah lapkeu (total_equity,
    # revenue, dst) ke Rupiah PENUH. 1_000_000 kalau laporan disajikan
    # dalam jutaan, 1_000_000_000 kalau dalam miliar. Dideteksi otomatis
    # dari metadata "Pembulatan yang digunakan" di setiap file -- default
    # 1_000_000 (jutaan) kalau metadata ini tidak ditemukan, karena itu
    # yang paling umum dipakai emiten IDX.
    reporting_unit_multiplier: int = 1_000_000
    # Kode mata uang pelaporan ("USD", "IDR", ...) dari metadata "Mata
    # uang pelaporan" -- ada di ekspor XLSX IDX. PENTING: sebagian kecil
    # emiten IDX menyajikan lapkeu dalam mata uang ASING (mis. POWR /
    # PT Cikarang Listrindo ber-USD). Kalau EPS/BVPS-nya dihitung langsung
    # melawan harga saham Rupiah, PER/PBV salah ribuan kali lipat (bug
    # nyata: PER POWR muncul 222.826x). None berarti tidak terdeteksi --
    # diperlakukan sebagai Rupiah (semua laporan PDF memang IDR).
    reporting_currency: Optional[str] = None

    # Deteksi laporan interim/kuartalan vs tahunan penuh. PENTING:
    # laporan tahunan IDX/OJK selalu berakhir 31 Desember. Kalau tanggal
    # periode BUKAN 31 Desember (mis. 30 Juni untuk laporan semester),
    # maka baris laba rugi & arus kas adalah angka KUMULATIF SEBAGIAN
    # TAHUN (mis. 6 bulan), BUKAN setahun penuh -- membandingkan ini
    # langsung dengan laporan tahunan (12 bulan) akan menghasilkan
    # growth rate yang menyesatkan (kelihatan "turun drastis" padahal
    # cuma beda cakupan waktu, bukan penurunan bisnis beneran).
    is_interim: bool = False
    coverage_months: Optional[int] = None  # None = tidak terdeteksi/tidak relevan (neraca)

    # --- Laporan Posisi Keuangan (Balance Sheet) ---
    total_assets: Optional[float] = None
    total_current_assets: Optional[float] = None
    total_liabilities: Optional[float] = None
    total_current_liabilities: Optional[float] = None
    total_equity: Optional[float] = None
    # Equity yang diatribusikan ke pemilik entitas INDUK saja (exclude
    # kepentingan non-pengendali/NCI). Ini yang dipakai untuk hitung Book
    # Value per Share / PBV -- BUKAN total_equity (yang termasuk NCI).
    # Bedanya bisa signifikan untuk holding company besar dengan banyak
    # anak perusahaan yang tidak dimiliki 100% (mis. ASII: total_equity
    # ~290T tapi equity_parent ~229T di tahun yang sama -- pakai yang salah
    # bikin PBV keliatan ratusan kali lipat dari yang sebenarnya).
    total_equity_parent: Optional[float] = None
    cash_and_equivalents: Optional[float] = None

    # --- Laporan Laba Rugi (Income Statement) ---
    revenue: Optional[float] = None
    profit_before_tax: Optional[float] = None
    net_income_parent: Optional[float] = None  # laba diatribusikan ke induk
    net_income_total: Optional[float] = None  # termasuk non-controlling interest
    basic_eps: Optional[float] = None
    interest_expense: Optional[float] = None

    # --- Laporan Arus Kas (Cash Flow Statement) ---
    operating_cash_flow: Optional[float] = None
    investing_cash_flow: Optional[float] = None
    financing_cash_flow: Optional[float] = None
    capex: Optional[float] = None  # payments for acquisition of PP&E
    dividends_paid: Optional[float] = None

    # Baris mentah yang berhasil ke-detect, buat debugging/verifikasi manual
    raw_matches: dict = {}


class FinancialStatement(BaseModel):
    """Hasil parsing satu file PDF -- bisa berisi 1-2 periode (current + prior)."""

    source_filename: str
    entity_name: Optional[str] = None
    periods: list[FinancialPeriod] = []
    parse_warnings: list[str] = []  # baris kunci yang GAGAL ditemukan
