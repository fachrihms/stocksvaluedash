"""
Utilitas parse label periode laporan IDX ('31 December 2025',
'31 Desember 2025', 'December 31, 2025', dst) ke objek date.

Dipakai bersama oleh parser lapkeu (deteksi interim di financial_pdf.py)
dan merge_periods (dedup & urutkan tren multi-tahun di ratios.py).

Nama bulan Indonesia DIDUKUNG karena ekspor Excel IDX kerap memakai
header Indonesia ("31 Desember 2025"), sementara strptime %B murni
Inggris di locale default -- tanpa normalisasi ini label Indonesia gagal
diparse, sehingga periode tak terdedup dan urutan tren multi-tahun jadi
salah (label tak dikenali jatuh ke fallback urutan terakhir).
"""

import re
from datetime import date, datetime

# Bulan Indonesia (nama penuh + singkatan umum) -> padanan Inggris yang
# dikenali strptime %B/%b. Januari-April & September-November ejaannya
# mirip Inggris tapi tetap dipetakan untuk keamanan.
_MONTH_MAP = {
    "januari": "January", "februari": "February", "maret": "March",
    "april": "April", "mei": "May", "juni": "June", "juli": "July",
    "agustus": "August", "oktober": "October", "november": "November",
    "desember": "December",
    "jan": "Jan", "feb": "Feb", "mar": "Mar", "apr": "Apr", "jun": "Jun",
    "jul": "Jul", "agu": "Aug", "ags": "Aug", "sep": "Sep", "okt": "Oct",
    "nov": "Nov", "des": "Dec",
}

# Format yang dipakai template IDX / ekspor Excel & PDF-nya. Dua format
# numerik terakhir untuk tanggal gaya "30/06/2026" yang sering muncul di
# metadata laporan interim ("6 Bulan yang berakhir pada 30/06/2026").
# Urutan penting: day-first dicek lebih dulu (konvensi Indonesia); tanggal
# gaya AS seperti 06/30/2026 otomatis jatuh ke format kedua karena 30
# tidak valid sebagai bulan.
_PERIOD_FORMATS = (
    "%d %B %Y", "%d %b %Y",       # 31 December 2025 / 31 Dec 2025
    "%B %d, %Y", "%b %d, %Y",     # December 31, 2025 / Dec 31, 2025
    "%B %d %Y", "%b %d %Y",       # December 31 2025 / Dec 31 2025
    "%d/%m/%Y", "%m/%d/%Y",       # 30/06/2026 / 06/30/2026
    "%Y-%m-%d",                   # 2023-12-31 (ISO, khas ekspor XLSX IDX)
)


def parse_period_label_date(label: str) -> date | None:
    """
    Parse label periode jadi date, mendukung nama bulan Inggris & Indonesia.
    Return None kalau label tidak bisa diparse (pemanggil menentukan
    fallback-nya sendiri -- jangan pernah asumsi label selalu valid).
    """
    if not label:
        return None
    normalized = re.sub(r"\s+", " ", str(label).strip())
    lowered = normalized.lower()
    # Ganti SEMUA kemunculan bulan Indonesia dengan padanan Inggris.
    # (Bulan pertama yang ketemu cukup -- label periode cuma punya satu.)
    for id_month, en_month in _MONTH_MAP.items():
        if re.search(rf"\b{re.escape(id_month)}\b", lowered):
            lowered = re.sub(rf"\b{re.escape(id_month)}\b", en_month, lowered)
            break
    for fmt in _PERIOD_FORMATS:
        try:
            return datetime.strptime(lowered, fmt).date()
        except ValueError:
            continue
    return None
