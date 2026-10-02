"""
Parser laporan keuangan PDF format IDX/OJK (template XBRL-to-PDF).

PENDEKATAN v2 -- coordinate-based row reconstruction.

Kenapa bukan `page.extract_text()` biasa (v1, dibuang)?
Karena pdfplumber's extract_text() menyusun baris berdasarkan urutan objek
teks di dalam PDF, dan untuk tabel finansial yang labelnya panjang/wrapping,
urutan itu SERING memisahkan label dari angkanya ke baris yang berbeda --
padahal secara visual mereka ada di baris yang sama. Contoh nyata dari lapkeu
ISAT: "Penjualan dan pendapatan usaha" ternyata BUKAN satu baris fisik --
"Penjualan dan pendapatan" ada di satu baris (top=195), lalu kata "usaha"
sendirian nyambung di baris bawahnya (top=207) karena wrapping. Untungnya
angkanya (56,518,141 55,886,870) sudah menempel di baris PERTAMA, jadi kita
tetap bisa menangkapnya asal kita mencocokkan label secara PARSIAL terhadap
baris yang direkonstruksi dari posisi (x, y) kata per kata -- bukan dari
urutan extract_text().

Cara kerja:
1. Ambil semua kata dengan `page.extract_words()` -- ini punya koordinat x0,
   top per kata.
2. Kelompokkan kata-kata itu berdasarkan `top` (posisi vertikal) yang sama
   (dengan toleransi kecil) -- ini merekonstruksi "baris visual" yang benar,
   sesuai yang terlihat mata di PDF.
3. Urutkan kata dalam satu baris berdasarkan `x0` (posisi horizontal) supaya
   urutannya benar kiri-ke-kanan.
4. Cari baris yang cocok dengan label yang kita incar, ambil angka dari baris
   itu.

Setiap ekstraksi disimpan di `raw_matches` untuk audit manual. JANGAN percaya
hasil parser 100% tanpa spot-check -- terutama untuk keputusan finansial
nyata.
"""

import re
import os
import pdfplumber
from app.models.financial import FinancialStatement, FinancialPeriod
from app.utils.period_dates import parse_period_label_date

ROW_TOLERANCE = 2.0

LABEL_MAP = [
    ("total_current_assets", r"^Jumlah aset lancar\b"),
    # "Jumlah aset" polos muncul 3x: "... lancar", "... tidak lancar" (kata
    # "lancar" di baris SETELAHNYA karena wrapping -- lihat EXCLUSION_CONTEXT),
    # dan baris totalnya sendiri (yang kita mau, TANPA kata lanjutan apapun).
    ("total_assets", r"^Jumlah aset\s+[\d(]"),
    ("total_current_liabilities", r"^Jumlah liabilitas\b"),  # disambiguasi lebih lanjut: lihat DISAMBIGUATION_CONTEXT, angka via lookahead
    ("total_liabilities", r"^Jumlah liabilitas\s+[\d(]"),  # disambiguasi: exclude "jangka"
    ("total_equity", r"^Jumlah ekuitas\s+[\d(]"),  # disambiguasi: exclude "yang" (dapat total TERMASUK NCI)
    # total_equity_parent: baris "Jumlah ekuitas" diikuti "yang [dapat]
    # diatribusikan ke entitas induk" -- TAPI kata "yang" ini sering
    # nyambung di baris SETELAHNYA (bukan di baris yang sama), akibat
    # wrapping. Pattern match ke baris "Jumlah ekuitas [angka]" yang sama
    # seperti total_equity biasa; PEMBEDA-nya ada di required_context
    # ("entitas induk") yang dicek di window setelahnya, bukan di pattern
    # itu sendiri. Ini kebalikan dari total_equity: total_equity exclude
    # "yang", sementara ini justru REQUIRE "entitas induk" muncul di window.
    # Varian label template Excel/IDX: "Ekuitas yang diatribusikan ke
    # entitas induk" (TIDAK diawali "Jumlah") -- alternatif kedua menangkap
    # baris ini TAPI WAJIB ada angka di baris yang sama ([^\d]*\d), karena
    # template XLSX juga punya baris HEADER seksi "Ekuitas yang diatribusikan
    # ..." TANPA angka -- tanpa syarat itu, header kosong itu akan mencuri
    # angka komponen ekuitas (saham biasa, dst) dari baris-baris di
    # bawahnya lewat forward-search. Pemisah utama tetap required_context
    # ("entitas induk"), yang di varian Excel sudah ada di baris itu sendiri.
    ("total_equity_parent", r"^Jumlah ekuitas\b|^Ekuitas yang diatribusikan\b[^\d]*\d"),
    ("cash_and_equivalents", r"^Kas dan setara kas\b(?!\s*arus)"),
    ("revenue", r"^Penjualan dan pendapatan\b"),
    ("profit_before_tax", r"^Jumlah laba \(rugi\) sebelum\b"),
    # net_income_parent: baris "Laba (rugi) yang dapat" muncul 2x -- satu
    # untuk "entitas induk" (yang kita mau), satu untuk "non-pengendali"
    # (non-controlling interest, BUKAN yang kita mau). Disambiguasi wajib.
    ("net_income_parent", r"^Laba \(rugi\) yang dapat\b"),
    # net_income_total: baris "Jumlah laba (rugi)" polos muncul beberapa kali
    # (before-tax, from-continuing-ops, total keseluruhan). Exclude varian
    # "sebelum" dan "dari" supaya dapat baris totalnya saja.
    ("net_income_total", r"^Jumlah laba \(rugi\)\s+[\d(]"),
    # basic_eps: label "Laba per saham dasar" muncul untuk continuing ops,
    # discontinued ops, dan diluted -- kita mau yang "dilanjutkan"
    # (continuing operations), disambiguasi via window.
    # Pattern punya 2 alternatif (OR) karena baris tempat angka EPS muncul
    # bisa dalam 2 bentuk berbeda tergantung file:
    #   (a) label lengkap + angka di baris yang sama: "Laba per saham
    #       dasar dari operasi yang dilanjutkan 836 715 ..."
    #   (b) label KEPOTONG oleh row-wrapping sehingga baris yang punya
    #       angka justru DIMULAI dari pertengahan frasa: "dasar dari
    #       operasi yang 715 499 share from continuing" -- di sini kata
    #       "Laba (rugi) per saham" ada di baris SEBELUMNYA, bukan di
    #       baris ini. Alternatif kedua menangkap pola ini langsung.
    ("basic_eps", r"(^Laba(?:\s*\(rugi\))? per saham(?:\s*dasar)?\b.{0,60}?\s+[\d(.])|"
                  r"(^dasar dari operasi yang\s+[\d(.])"),
    ("interest_expense", r"^Beban bunga dan keuangan\b"),
    # Ketiga field cash flow ini pakai label Indonesia YANG SAMA PERSIS
    # ("Jumlah arus kas bersih") -- satu-satunya pembeda adalah teks Inggris
    # "operating/investing/financing activities" yang muncul beberapa baris
    # setelahnya (lihat DISAMBIGUATION_CONTEXT). Tanpa disambiguasi ini
    # parser akan selalu ambil baris operating (yang pertama muncul) untuk
    # ketiganya.
    # "(?:arus )?" menangkap 2 varian template IDX: "Jumlah arus kas bersih
    # dari aktivitas ..." (umum di PDF/ISAT) dan "Jumlah kas bersih diterima
    # dari aktivitas ..." (umum di ekspor XLSX/Excel).
    ("operating_cash_flow", r"^Jumlah (?:arus )?kas bersih\b"),
    ("investing_cash_flow", r"^Jumlah (?:arus )?kas bersih\b"),
    ("financing_cash_flow", r"^Jumlah (?:arus )?kas bersih\b"),
    # XLSX/Excel memakai label Indonesia "Perolehan aset tetap" untuk baris
    # capex; PDF memakai "Pembayaran untuk ...". Pattern WAJIB memuat
    # "aset tetap" di baris yang sama: template XLSX punya BANYAK baris
    # "Pembayaran untuk ..." lain (properti investasi, entitas anak, future
    # contracts, dst) yang tanpa syarat ini bisa menjadi kandidat pertama
    # -- lalu forward-search angkanya menyedot angka baris LAIN di
    # sekitarnya (bug nyata POWR: capex malah ambil angka baris "Penerimaan
    # dari penjualan aset tetap" -- disposal, bukan acquisition).
    ("capex", r"^(?:Pembayaran untuk|Perolehan)\b.*\baset tetap\b"),  # diverifikasi via window
    # dividends_paid: "Pembayaran dividen dari" muncul di aktivitas OPERASI
    # (dividen diterima dari entitas asosiasi, biasanya 0/kosong) DAN di
    # aktivitas PENDANAAN (dividen ke pemegang saham -- yang kita mau).
    # Tanpa "dari" di ujung: ekspor XLSX memakai "Pembayaran dividen" polos.
    # Pembeda operasi vs pendanaan tetap lewat DISAMBIGUATION_CONTEXT /
    # EXCLUSION_CONTEXT (window), bukan pattern-nya.
    ("dividends_paid", r"^Pembayaran dividen\b"),
]

# Untuk label yang match pattern-nya sama tapi mengacu ke item berbeda, kita
# verifikasi baris kandidat dengan mengecek kata kunci ini muncul di window
# SEKITAR baris yang match. Kalau tidak muncul, baris itu dilewati dan
# parser lanjut cari kandidat berikutnya yang match pattern.
DISAMBIGUATION_CONTEXT = {
    "capex": ["perolehan aset tetap", "acquisition of property"],  # dicek juga via keyword-pair fallback, lihat CAPEX_KEYWORD_PAIRS
    "net_income_parent": ["entitas induk", "parent entity"],
    "basic_eps": ["dilanjutkan", "continuing"],
    # Kata kunci bilingual: PDF ISAT pakai teks Inggris, ekspor XLSX banyak
    # yang pakai label Indonesia ("... dari aktivitas operasi/pendanaan").
    "operating_cash_flow": ["operating activities", "aktivitas operasi"],
    "investing_cash_flow": ["investing activities", "aktivitas investasi"],
    "financing_cash_flow": ["financing activities", "aktivitas pendanaan", "aktivitas pendbiayaan"],
    "dividends_paid": ["aktivitas pendanaan", "financing activities"],
    "total_current_liabilities": ["jangka pendek", "current liabilities"],
    "total_equity_parent": ["entitas induk", "attributable to", "owners of parent"],
}
# Kata kunci yang, kalau muncul di window, berarti baris ini BUKAN kandidat
# yang benar meski pattern utama cocok (exclude list, bukan require list).
# PENTING: window untuk exclusion di sini HANYA dicek 1 baris ke depan
# (bukan WINDOW_AFTER penuh) supaya nggak salah exclude baris total yang
# valid gara-gara kata "lancar"/"jangka" muncul di baris JAUH sebelumnya
# akibat window overlap antar item balance sheet yang berdekatan.
EXCLUSION_CONTEXT = {
    "net_income_parent": ["non-pengendali", "non-controlling"],
    "dividends_paid": ["aktivitas operasi", "operating activities"],
    "total_assets": ["lancar", "current assets", "current asset"],
    "total_liabilities": ["jangka", "current liabilities", "non-current liabilities"],
    "total_equity": ["yang", "attributable to"],
}
# Untuk field di EXCLUSION_CONTEXT ini, exclusion hanya dicek di N baris
# TEPAT SETELAH baris match (bukan window penuh) -- karena kata pemisahnya
# selalu nyambung persis di baris berikutnya akibat wrapping, dan window
# penuh terlalu lebar sehingga bisa nyerempet baris item lain yang tidak
# terkait.
NARROW_EXCLUSION_LOOKAHEAD = {"total_assets", "total_liabilities", "total_equity", "net_income_parent"}

# Untuk capex, frasa "perolehan aset tetap" kadang terpecah oleh angka di
# antaranya (mis. "...perolehan payments for acquisition of ( 9,846 )
# ( 4,687 ) aset tetap property, plant and equipment..." -- kata
# "perolehan" dan "aset tetap" nyata-nyata ada tapi TIDAK bersebelahan
# sebagai substring karena angka ikut nyempil di antaranya). Untuk kasus
# ini, exact-phrase substring match di DISAMBIGUATION_CONTEXT akan selalu
# gagal walau isinya benar secara semantik. Fallback: cek KEDUA kata kunci
# ini muncul di window secara independen (urutan/jarak bebas), bukan
# sebagai satu frasa utuh.
CAPEX_KEYWORD_PAIRS = ["perolehan", "aset tetap"]
NARROW_LOOKAHEAD_ROWS = 1


def _exclusion_keyword_wins(
    text: str,
    required_context: list[str] | None,
    excluded_context: list[str] | None,
) -> bool:
    """
    True = baris kandidat harus DIEXCLUDE meski pattern-nya cocok.

    Default: cukup ada satu kata exclusion di teks -> exclude (perilaku
    lama). PENGECALIAN: kalau konteks wajib (required_context) justru
    muncul LEBIH DULU daripada kata exclusion dalam teks yang sama, jangan
    exclude. Kenapa: template ekspor Excel IDX meletakkan baris
    "kepentingan non-pengendali" TEPAT SETELAH baris "Laba (rugi) yang
    dapat ... entitas induk", keduanya diawali frasa yang sama -- dengan
    pemeriksaan posisi, baris entitas induk tetap diterima (kata "entitas
    induk" ada lebih dulu), sedangkan baris NCI tetap ter-exclude (kata
    "non-pengendali" ada di barisnya sendiri, sebelum konteks wajib
    manapun). Exclusion lama yang buta-posisi malah men-exclude baris
    entitas induk yang BENAR (bug nyata di file POWR).
    """
    if not excluded_context:
        return False
    text_lower = text.lower()
    exc_positions = [
        text_lower.find(c.lower()) for c in excluded_context if c.lower() in text_lower
    ]
    if not exc_positions:
        return False
    if not required_context:
        return True
    req_positions = [
        text_lower.find(c.lower()) for c in required_context if c.lower() in text_lower
    ]
    if not req_positions:
        return True
    return min(exc_positions) < min(req_positions)

# Window scan: berapa baris ke BELAKANG dan ke DEPAN yang dicek untuk
# menyatukan label yang terpecah dari angkanya, dan untuk disambiguasi
# konteks (lihat docstring modul).
WINDOW_BEFORE = 2
WINDOW_AFTER = 4

NUMBER_PATTERN = re.compile(r"\(?\s*-?[\d,]+(?:\.\d+)?\s*\)?")


def _parse_number(raw: str) -> float | None:
    raw = raw.strip()
    if not raw:
        return None
    is_negative = raw.startswith("(") and raw.endswith(")")
    cleaned = raw.strip("()").strip().replace(",", "")
    if not cleaned or not re.match(r"^-?\d+(\.\d+)?$", cleaned):
        return None
    value = float(cleaned)
    return -value if is_negative else value


def _extract_numbers(text: str) -> list[float]:
    numbers = []
    for m in NUMBER_PATTERN.findall(text):
        val = _parse_number(m)
        if val is not None:
            numbers.append(val)
    return numbers


def _reconstruct_rows(words: list[dict]) -> list[str]:
    """
    Kelompokkan kata-kata jadi baris visual berdasarkan koordinat 'top',
    lalu urutkan tiap baris berdasarkan 'x0'. Return list of row strings,
    terurut dari atas ke bawah halaman.
    """
    if not words:
        return []

    sorted_words = sorted(words, key=lambda w: w["top"])

    rows: list[list[dict]] = []
    current_row: list[dict] = [sorted_words[0]]
    current_top = sorted_words[0]["top"]

    for w in sorted_words[1:]:
        if abs(w["top"] - current_top) <= ROW_TOLERANCE:
            current_row.append(w)
        else:
            rows.append(current_row)
            current_row = [w]
            current_top = w["top"]
    rows.append(current_row)

    row_strings = []
    for row in rows:
        row_sorted = sorted(row, key=lambda w: w["x0"])
        row_strings.append(" ".join(w["text"] for w in row_sorted))
    return row_strings


def _get_all_rows(filepath: str) -> list[str]:
    all_rows: list[str] = []
    with pdfplumber.open(filepath) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            all_rows.extend(_reconstruct_rows(words))
    return all_rows


def _detect_period_labels(rows: list[str]) -> list[str]:
    """
    Cari label periode current & prior year, dengan DUA strategi berlapis
    karena format tanggal berbeda-beda antar file/sumber ekspor IDX:

    Strategi 1 -- baris tabel 2-kolom seperti '31 December 2025 31 December
    2024' di Laporan Posisi Keuangan. Ini format paling umum. Dua pola
    tanggal didukung: '31 December 2025' (tanggal-bulan-tahun) DAN
    'December 31, 2025' (bulan-tanggal,-tahun) -- beberapa file ternyata
    pakai format kedua (ditemukan pertama kali di file ASII 2023, yang
    filenya sepertinya hasil ekspor dari tool berbeda dari file ASII
    tahun lain -- tapi ini bukan sesuatu yang unik ke satu emiten, jadi
    kedua pola tetap dicek untuk SEMUA file).

    Strategi 2 (fallback) -- kalau strategi 1 gagal, cari baris metadata
    XBRL eksplisit "Tanggal akhir periode berjalan" (current) dan "Tanggal
    akhir periode sebelumnya" atau "Tanggal akhir tahun sebelumnya" (prior)
    yang biasanya muncul di bagian "General information" / cover laporan.
    Baris ini vs baris tabel adalah SUMBER BERBEDA di dokumen yang sama,
    jadi kalau satu gagal karena format aneh, yang lain kemungkinan besar
    tetap terbaca.

    Scan dilakukan di SELURUH dokumen (bukan cuma N baris pertama), karena
    beberapa file (mis. laporan grup holding besar dengan banyak anak
    perusahaan) punya halaman pembuka yang sangat panjang sebelum masuk ke
    laporan keuangan inti -- baris tanggal bisa ada di baris ke-1000+.
    """
    # Strategi 1a: '31 December 2025' (day month year)
    date_pattern_dmy = re.compile(r"\b\d{1,2} \w+ \d{4}\b")
    for row in rows:
        matches = date_pattern_dmy.findall(row)
        # Baris yang MENGULANG tanggal yang sama (khas catatan/audit report
        # di ekspor XLSX) bukan header periode -- dua periode yang sah
        # selalu berbeda tanggalnya. Lewati, jangan jadikan label.
        if len(matches) >= 2 and matches[0] != matches[1]:
            return matches

    # Strategi 1b: 'December 31, 2025' (month day, year)
    date_pattern_mdy = re.compile(r"\b\w+ \d{1,2},? \d{4}\b")
    for row in rows:
        matches = date_pattern_mdy.findall(row)
        # Filter: pastikan match diawali nama bulan asli, bukan kata acak
        # yang kebetulan diikuti angka+tahun (mengurangi false positive).
        valid = [m for m in matches if re.match(
            r"^(January|February|March|April|May|June|July|August|"
            r"September|October|November|December|Januari|Februari|Maret|"
            r"April|Mei|Juni|Juli|Agustus|September|Oktober|November|"
            r"Desember)\b", m, re.IGNORECASE
        )]
        if len(valid) >= 2 and valid[0] != valid[1]:
            return valid

    # Strategi 2 (fallback): metadata XBRL eksplisit "Tanggal akhir ..."
    current_label = None
    prior_label = None
    label_date_pattern = re.compile(r"\b\w+ \d{1,2},? \d{4}\b")
    for row in rows:
        if re.search(r"Tanggal akhir periode berjalan|Current period end date", row, re.IGNORECASE):
            m = label_date_pattern.search(row)
            if m:
                current_label = m.group().strip()
        elif re.search(
            r"Tanggal akhir (periode|tahun) sebelumnya|Prior (period|year) end date",
            row, re.IGNORECASE
        ):
            m = label_date_pattern.search(row)
            if m and prior_label is None:  # ambil yang pertama ketemu saja
                prior_label = m.group().strip()

    if current_label and prior_label:
        return [current_label, prior_label]

    return []


def _is_december_31(period_label: str) -> bool:
    """
    True kalau period_label menunjukkan tanggal 31 Desember, dalam SALAH
    SATU format berikut (beda file/sumber ekspor IDX pakai format berbeda,
    lihat _detect_period_labels untuk konteks lengkap):
    - '31 December 2025' / '31 Desember 2025'  (tanggal-bulan-tahun)
    - 'December 31, 2025' / 'Desember 31, 2025' (bulan-tanggal,-tahun)
    """
    label = period_label.strip().lower()
    if label.startswith("31 december") or label.startswith("31 desember"):
        return True
    if re.match(r"^(december|desember) 31\b", label):
        return True
    # Format lain -- termasuk ISO "2023-12-31" khas ekspor XLSX IDX --
    # dicocokkan lewat tanggal yang ter-parse: laporan tahunan IDX selalu
    # berakhir 31 Desember, jadi cukup cek bulan & tanggal-nya. Tanpa ini,
    # laporan tahunan dari XLSX salah tertandai interim (bug nyata POWR).
    parsed = parse_period_label_date(period_label)
    return parsed is not None and parsed.month == 12 and parsed.day == 31


def _detect_interim_info(rows: list[str], period_label: str) -> tuple[bool, int | None]:
    """
    Deteksi apakah laporan ini interim/kuartalan (bukan tahunan penuh),
    dan berapa bulan cakupannya, KHUSUS untuk satu period_label tertentu.

    Dua sinyal dipakai:
    1. Tanggal periode BUKAN 31 Desember -- laporan tahunan IDX/OJK selalu
       berakhir 31 Desember, jadi tanggal lain (30 Juni, 31 Maret, 30
       September) berarti ini laporan interim. Lihat _is_december_31 untuk
       format tanggal yang didukung.
    2. Teks eksplisit "X bulan yang berakhir [TANGGAL]" / "X months ended
       [DATE]" -- dipakai HANYA kalau [TANGGAL] pada baris itu cocok
       dengan period_label yang sedang diperiksa. Ini PENTING: dokumen yang
       sama dipindai untuk period_current DAN period_prior secara terpisah
       (lihat pemanggilan fungsi ini), tapi teks "X bulan yang berakhir"
       biasanya cuma muncul SEKALI di metadata halaman awal dan menjelaskan
       periode CURRENT saja (mis. "6 Bulan yang berakhir pada 30/06/2026").
       Tanpa pencocokan tanggal ini, match yang sama akan "bocor" dan
       salah diterapkan ke period_prior (mis. 31 Desember 2025 ikut
       ke-flag coverage_months=6, padahal itu laporan tahunan penuh) --
       bug nyata yang pernah terjadi dan menyebabkan laporan tahunan
       salah tertandai sebagai interim.
    """
    is_dec_31 = _is_december_31(period_label)

    # Ambil komponen tanggal dari period_label untuk pencocokan silang
    # dengan tanggal yang disebut di baris "X bulan yang berakhir [tanggal]".
    # period_label bisa dalam beberapa format ("31 December 2025",
    # "31 Desember 2025", "December 31, 2025"), jadi cocokkan lewat objek
    # date yang sudah diparse, bukan string mentah. Parser-nya util bersama
    # (period_dates) yang juga mendukung nama bulan Indonesia -- penting
    # untuk ekspor Excel IDX yang header-nya "31 Desember 2025".
    period_date = parse_period_label_date(period_label)

    coverage_months = None
    month_word_map = {
        "satu": 1, "dua": 2, "tiga": 3, "empat": 4, "lima": 5, "enam": 6,
        "tujuh": 7, "delapan": 8, "sembilan": 9, "sepuluh": 10,
        "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
        "seven": 7, "eight": 8, "nine": 9, "ten": 10,
    }
    # Pattern juga menangkap tanggal opsional setelah "yang berakhir pada" /
    # "ended" untuk verifikasi bahwa baris ini benar-benar menjelaskan
    # period_label yang sedang dicek. PENTING soal urutan alternasi tanggal:
    # regex mencoba alternasi KIRI-KE-KANAN -- kalau pola angka/slash bebas
    # ([\d/]+) dicantumkan lebih dulu, "31 Desember 2023" akan match "31"
    # saja (regex berhenti di match valid pertama) dan tanggal sebenarnya
    # hilang. Bug nyata: verifikasi tanggal gagal parse "31" -> disclosure
    # "...untuk periode sembilan bulan yang berakhir 31 Desember 2023..."
    # di catatan PGAS lolos dan salah menandai laporan TAHUNAN sebagai
    # interim 9 bulan. Alternatif tanggal paling spesifik HARUS dicoba
    # duluan: "31 Desember 2023", lalu "December 31, 2023", baru DD/MM/YYYY.
    date_alt = r"(\d{1,2} \w+ \d{4}|\w+ \d{1,2},? \d{4}|\d{1,2}/\d{1,2}/\d{2,4})"
    pattern = re.compile(
        r"\b(satu|dua|tiga|empat|lima|enam|tujuh|delapan|sembilan|sepuluh|"
        r"one|two|three|four|five|six|seven|eight|nine|ten)\b\s*"
        r"(?:\(\w+\)\s*)?(bulan|months?)\s+(?:yang berakhir|ended)"
        r"(?:\s+(?:pada|on))?\s*" + date_alt + r"?|"
        r"\b([1-9]|10|11)\s+(bulan|months?)\s+(?:yang berakhir|ended)"
        r"(?:\s+(?:pada|on))?\s*" + date_alt + r"?",
        re.IGNORECASE,
    )
    for row in rows:
        match = pattern.search(row)
        if not match:
            continue
        coverage_value = match.group(1) or match.group(4)
        date_str = match.group(3) or match.group(6)

        candidate_months = month_word_map.get(coverage_value.lower()) if coverage_value else None
        if candidate_months is None and coverage_value and coverage_value.isdigit():
            candidate_months = int(coverage_value)

        if candidate_months is None:
            continue

        # Kalau baris ini menyebutkan tanggal, WAJIB cocok dengan
        # period_label yang sedang diperiksa sebelum dipakai -- ini yang
        # mencegah "bocor" ke periode lain dalam dokumen yang sama.
        if date_str and period_date:
            row_date = parse_period_label_date(date_str)
            if row_date is not None and row_date != period_date:
                continue  # tanggal di baris ini untuk periode LAIN, lewati
            # "X bulan yang berakhir 31 Desember [tahun label]" itu
            # kontradiktif untuk laporan tahunan: laporan interim IDX selalu
            # berakhir 31 Maret/Juni/September, jadi frasa berakhir tepat di
            # tanggal label 31 Desember hampir pasti milik disclosure
            # kombinasi bisnis/akuisisi (entitas yang diakuisisi punya
            # periode pelaporannya sendiri), BUKAN cakupan laporan ini.
            # Contoh nyata: catatan PGAS "...pendapatan konsolidasian untuk
            # periode sembilan bulan yang berakhir 31 Desember 2023 dan
            # 2022..." sempat membuat laporan TAHUNAN tertandai interim.
            if row_date is not None and _is_december_31(period_label):
                continue

        coverage_months = candidate_months
        break

    # Tanggal selain 31 Desember sudah cukup untuk menandai interim. Selain
    # itu, teks cakupan eksplisit (yang SUDAH divalidasi cocok dengan
    # period_label ini, lihat di atas) menang atas label tanggal: beberapa
    # PDF meletakkan header neraca 31 Desember lebih dulu, sementara angka
    # laba-rugi/arus-kasnya ternyata baru mencakup 3/6/9 bulan. Menganggap
    # angka tersebut sebagai satu tahun penuh akan membuat growth dan DPR
    # sangat menyesatkan.
    is_interim = not is_dec_31 or (coverage_months is not None and coverage_months < 12)

    return is_interim, coverage_months


def _extract_entity_name(rows: list[str]) -> str | None:
    for row in rows[:60]:
        if "Nama Emiten" in row or "Nama entitas" in row:
            parts = re.split(r"Nama Emiten|Nama entitas", row)
            if len(parts) > 1 and parts[1].strip():
                candidate = parts[1].strip()
                candidate = re.split(r"\s{2,}|Entity Name", candidate)[0]
                return candidate.strip()
    return None


def _detect_reporting_unit_multiplier(rows: list[str]) -> int:
    """
    Deteksi satuan pembulatan yang dipakai laporan ini, dari metadata XBRL
    eksplisit "Pembulatan yang digunakan dalam penyajian jumlah" yang
    SELALU ada di setiap laporan IDX/OJK (biasanya di halaman awal, bagian
    "General information"). Nilainya "Jutaan / In Million" ATAU "Miliaran
    / In Billion" tergantung skala emiten -- BUKAN seragam untuk semua
    emiten. Contoh nyata yang jadi alasan fungsi ini dibuat: ISAT pakai
    "Jutaan", ASII pakai "Miliaran" -- kalau diasumsikan sama-sama jutaan
    (asumsi lama sebelum fungsi ini ada), semua angka per-saham (BVPS,
    dividend per share) untuk ASII akan salah faktor 1000x, dan PBV/valuasi
    yang dihitung dari situ jadi menyesatkan total (bug nyata yang pernah
    terjadi: PBV ASII muncul 683x, seharusnya ~1.2x).

    Return 1_000_000 (jutaan) sebagai default kalau metadata ini tidak
    ditemukan -- itu satuan paling umum dipakai emiten IDX, jadi asumsi
    paling aman ketika tidak ada info eksplisit.
    """
    for idx, row in enumerate(rows[:200]):  # metadata ini selalu di halaman awal
        row_lower = row.lower()
        if "pembulatan" in row_lower or "level of rounding" in row_lower:
            # Nilai satuannya bisa di baris yang SAMA dengan labelnya, ATAU
            # di baris terpisah SESUDAHNYA (umum di XLSX: label dan nilai
            # ada di sel berbeda yang dirender sebagai baris berurutan).
            for candidate in [row_lower] + [r.lower() for r in rows[idx + 1 : idx + 4]]:
                if "miliar" in candidate or "billion" in candidate:
                    return 1_000_000_000
                if "juta" in candidate or "million" in candidate:
                    return 1_000_000
    return 1_000_000  # default: jutaan (paling umum)


def _extract_fiscal_year(period_label: str) -> int | None:
    """Ambil tahun dari label periode ('31 December 2025' -> 2025)."""
    match = re.search(r"\b(19|20)\d{2}\b", period_label)
    return int(match.group()) if match else None


def parse_financial_statement_rows(filename: str, rows: list[str]) -> FinancialStatement:
    """
    Engine bersama PDF & XLSX: parse laporan keuangan IDX/OJK dari daftar
    BARIS TEKS biasa (satu string per baris visual/Excel), return
    FinancialStatement dengan 2 periode (current + prior).

    Parser PDF memasok baris hasil rekonstruksi koordinat kata
    (lihat _reconstruct_rows); parser XLSX (financial_xlsx.py) memasok
    baris hasil join sel per baris spreadsheet. Semua logika
    label-matching, disambiguasi, deteksi periode/interim/satuan ada DI
    SINI supaya kedua format berperilaku identik dan tidak ada yang
    ke-bugfix sepihak.
    """

    period_labels = _detect_period_labels(rows)
    label_current = period_labels[0] if period_labels else "current"
    label_prior = period_labels[1] if len(period_labels) > 1 else "prior"
    period_current = FinancialPeriod(period_label=label_current, source_filename=filename)
    period_prior = FinancialPeriod(period_label=label_prior, source_filename=filename)

    warnings: list[str] = []

    for field_name, pattern in LABEL_MAP:
        regex = re.compile(pattern, re.IGNORECASE)
        required_context = DISAMBIGUATION_CONTEXT.get(field_name)
        excluded_context = EXCLUSION_CONTEXT.get(field_name)
        found = False

        for i, row in enumerate(rows):
            if not regex.search(row):
                continue

            window_start = max(0, i - WINDOW_BEFORE)
            window_end = min(len(rows), i + WINDOW_AFTER + 1)
            window_rows = rows[window_start:window_end]
            window_text = " ".join(window_rows)

            # Kalau field ini butuh disambiguasi (contoh: capex punya beberapa
            # baris "Pembayaran untuk ..." yang mirip), pastikan konteks
            # yang diharapkan benar-benar muncul sebelum diterima.
            # Field-field berikut pakai NARROW lookahead (bukan window
            # lebar) karena baris berikutnya sering berisi HEADER SECTION
            # BARU yang juga cocok dengan salah satu kata kunci disambiguasi
            # -- window lebar bikin baris "Jumlah arus kas bersih ...
            # operasi" (yang cuma valid untuk operating_cash_flow) ikut
            # lolos required_context check untuk investing_cash_flow, karena
            # window-nya nyerempet header "Cash flows from investing
            # activities" yang muncul TEPAT SETELAHNYA (bug nyata yang
            # ditemukan di file ASII 2023: investing_cash_flow salah
            # mengambil nilai operating_cash_flow).
            if field_name == "total_current_liabilities":
                narrow_text = " ".join(
                    rows[i : min(len(rows), i + 3)]
                )
                if not any(ctx.lower() in narrow_text.lower() for ctx in required_context):
                    continue
            elif field_name in ("operating_cash_flow", "investing_cash_flow", "financing_cash_flow"):
                # Prioritas 1: cek baris itu SENDIRI dulu. Beberapa file
                # (mis. ASII) punya label lengkap dalam SATU baris utuh
                # ("Jumlah arus kas bersih ... aktivitas operasi ...
                # 33,746 ..."), jadi baris itu sendiri sudah cukup untuk
                # menentukan field mana yang cocok -- TIDAK PERLU DAN TIDAK
                # BOLEH melihat baris berikutnya, karena baris berikutnya
                # sering berisi header section CASH FLOW LAIN yang juga
                # mengandung salah satu kata kunci (mis. investing_cash_flow
                # bisa salah mengambil baris operating_cash_flow kalau ikut
                # melihat baris +1 yang kebetulan berisi header "Cash flows
                # from investing activities").
                #
                # Prioritas 2 (fallback, HANYA kalau baris itu sendiri tidak
                # mengandung SATU PUN dari ketiga kata kunci): baris itu
                # netral/polos (mis. ISAT: "Jumlah arus kas bersih" tanpa
                # keterangan tambahan), kata kuncinya baru muncul beberapa
                # baris kemudian akibat wrapping -- di sini AMAN untuk
                # melihat ke depan karena baris ini sendiri belum "mengklaim"
                # jadi field manapun.
                all_cf_keywords = ["operating activities", "investing activities", "financing activities", "aktivitas operasi", "aktivitas investasi", "aktivitas pendanaan", "aktivitas pendbiayaan"]
                row_lower = row.lower()
                row_has_any_keyword = any(kw in row_lower for kw in all_cf_keywords)

                if row_has_any_keyword:
                    if not any(ctx.lower() in row_lower for ctx in required_context):
                        continue
                else:
                    narrow_text = " ".join(rows[i : min(len(rows), i + 4)]).lower()
                    # Pastikan TIDAK ADA keyword field lain yang nyelip di
                    # window fallback ini sebelum keyword yang benar muncul
                    # -- kalau keyword field lain muncul LEBIH DULU (index
                    # lebih kecil) daripada keyword yang kita cari, berarti
                    # baris ini sebenarnya milik field lain, bukan field ini.
                    own_positions = [narrow_text.find(ctx.lower()) for ctx in required_context if ctx.lower() in narrow_text]
                    other_keywords = [kw for kw in all_cf_keywords if kw not in [c.lower() for c in required_context]]
                    other_positions = [narrow_text.find(kw) for kw in other_keywords if kw in narrow_text]
                    if not own_positions:
                        continue
                    own_first = min(own_positions)
                    if other_positions and min(other_positions) < own_first:
                        continue
            elif field_name == "capex":
                # HANYA exact-phrase match di sini (mode ketat). Fallback
                # keyword-pair independen (CAPEX_KEYWORD_PAIRS) TIDAK
                # dicoba per-baris di sini -- itu dilakukan di PASS KEDUA
                # terpisah setelah seluruh LABEL_MAP diproses sekali,
                # HANYA jika field capex masih None (lihat kode setelah
                # loop utama). Alasan dipisah jadi 2 pass: kalau fallback
                # longgar dicoba di baris pertama yang gagal exact-phrase,
                # dia bisa "mencuri" match pada baris yang SALAH sebelum
                # sempat mencoba baris-baris lain yang justru cocok exact
                # (bug nyata yang ditemukan: capex ISAT balik salah ambil
                # baris "Penerimaan dari penjualan aset tetap" karena kedua
                # kata kunci fallback kebetulan ada di window baris itu).
                if not any(ctx.lower() in window_text.lower() for ctx in required_context):
                    continue
            elif required_context and not any(
                ctx.lower() in window_text.lower() for ctx in required_context
            ):
                continue

            # Exclusion: untuk field yang rawan salah tangkap gara-gara kata
            # pemisah wrapping ada persis di baris berikutnya (mis. "Jumlah
            # aset" polos vs "Jumlah aset [tidak] lancar"), cek HANYA
            # beberapa baris tepat setelahnya -- bukan window penuh -- supaya
            # tidak salah exclude baris total yang valid karena kata serupa
            # kebetulan ada di item lain yang jauh dalam window.
            if field_name in NARROW_EXCLUSION_LOOKAHEAD:
                lookahead_text = " ".join(
                    rows[i : min(len(rows), i + NARROW_LOOKAHEAD_ROWS + 1)]
                )
                if _exclusion_keyword_wins(
                    lookahead_text, required_context, excluded_context
                ):
                    continue
            elif _exclusion_keyword_wins(
                window_text, required_context, excluded_context
            ):
                continue

            # Cari angka: coba baris itu sendiri dulu, baru cari ke DEPAN
            # saja (bukan window dua arah) kalau baris itu sendiri nggak
            # punya 2 angka -- kasus label duluan, angka nyusul beberapa
            # baris kemudian (mis. operating_cash_flow, capex).
            # PENTING: sengaja tidak ikut baris SEBELUM (WINDOW_BEFORE) di
            # sini, karena baris sebelumnya sering berisi angka milik item
            # lain yang tidak terkait (contoh nyata: baris "Laba (rugi) yang
            # dapat" tanpa angka, kalau ikut baris sebelumnya bisa salah
            # ambil angka comprehensive income dari item yang berbeda).
            numbers = _extract_numbers(row)
            source_text = row
            if len(numbers) < 2:
                forward_text = " ".join(rows[i : window_end])
                forward_numbers = _extract_numbers(forward_text)
                if len(forward_numbers) >= 2:
                    numbers = forward_numbers
                    source_text = forward_text

            if len(numbers) >= 2:
                setattr(period_current, field_name, numbers[0])
                setattr(period_prior, field_name, numbers[1])
                period_current.raw_matches[field_name] = source_text.strip()
                found = True
                break
            elif len(numbers) == 1:
                setattr(period_current, field_name, numbers[0])
                period_current.raw_matches[field_name] = source_text.strip()
                warnings.append(
                    f"'{field_name}': cuma ketemu 1 angka (bukan 2 -- current & prior), "
                    f"cek manual: \"{source_text.strip()}\""
                )
                found = True
                break

        if not found:
            warnings.append(
                f"'{field_name}': label tidak ditemukan atau tidak ada angka valid "
                f"di sekitar baris manapun (pattern: {pattern})"
            )

    # PASS KEDUA (fallback longgar) khusus untuk capex, HANYA dijalankan
    # kalau pass pertama (exact-phrase, mode ketat) di atas gagal total.
    # Ini menangani kasus di mana frasa "perolehan aset tetap" terpecah
    # oleh angka yang menyempil di tengahnya akibat wrapping tabel (mis.
    # "...perolehan payments for acquisition of ( 9,846 ) ( 4,687 ) aset
    # tetap property, plant and equipment..." -- kata "perolehan" dan
    # "aset tetap" keduanya ADA tapi tidak bersebelahan sebagai substring).
    # Dijalankan sebagai pass TERPISAH (bukan dicoba per-baris di pass
    # pertama) supaya tidak salah mengambil baris pertama yang kebetulan
    # punya kedua kata kunci ini secara independen tapi maknanya beda
    # (bug nyata yang sempat terjadi: capex ISAT salah ambil baris
    # "Penerimaan dari penjualan aset tetap" karena fallback dicoba
    # terlalu dini, sebelum baris capex yang benar sempat dicek).
    if period_current.capex is None:
        capex_pattern_idx = next(i for i, (fn, _) in enumerate(LABEL_MAP) if fn == "capex")
        capex_regex = re.compile(LABEL_MAP[capex_pattern_idx][1], re.IGNORECASE)
        for i, row in enumerate(rows):
            if not capex_regex.search(row):
                continue
            window_start = max(0, i - WINDOW_BEFORE)
            window_end = min(len(rows), i + WINDOW_AFTER + 1)
            window_text = " ".join(rows[window_start:window_end])
            if not all(kw.lower() in window_text.lower() for kw in CAPEX_KEYWORD_PAIRS):
                continue
            numbers = _extract_numbers(row)
            source_text = row
            if len(numbers) < 2:
                forward_text = " ".join(rows[i:window_end])
                forward_numbers = _extract_numbers(forward_text)
                if len(forward_numbers) >= 2:
                    numbers = forward_numbers
                    source_text = forward_text
            if len(numbers) >= 2:
                period_current.capex = numbers[0]
                period_prior.capex = numbers[1]
                period_current.raw_matches["capex"] = source_text.strip()
                # Hapus warning capex yang sempat ditambahkan di pass
                # pertama, karena sekarang berhasil ditemukan di pass kedua.
                warnings[:] = [w for w in warnings if not w.startswith("'capex'")]
                break

    # Konvensi tanda: baris pembayaran (capex, dividen) adalah ARUS KELUAR.
    # Template PDF menuliskannya dalam kurung (sudah negatif setelah parse),
    # tapi template Excel IDX menyimpannya sebagai angka POSITIF -- samakan
    # jadi negatif di sini supaya free_cash_flow (OCF + capex) dan metrik
    # lain konsisten di kedua format. Nilai None/0/negatif dibiarkan.
    for period in (period_current, period_prior):
        for field in ("capex", "dividends_paid"):
            value = getattr(period, field)
            if value is not None and value > 0:
                setattr(period, field, -value)

    entity_name = _extract_entity_name(rows)
    reporting_unit_multiplier = _detect_reporting_unit_multiplier(rows)

    # Deteksi interim & isi fiscal_year untuk kedua periode. Prior period
    # BIASANYA juga interim kalau current-nya interim (laporan Q2 2026
    # akan membandingkan ke Q2 2025, bukan ke tahunan 2025) -- tapi kita
    # deteksi dari label masing-masing periode secara independen supaya
    # tetap benar untuk kasus yang tidak biasa.
    for period in (period_current, period_prior):
        is_interim, coverage_months = _detect_interim_info(rows, period.period_label)
        period.is_interim = is_interim
        period.coverage_months = coverage_months
        period.fiscal_year = _extract_fiscal_year(period.period_label)
        # Satuan pembulatan berlaku untuk SELURUH dokumen (current & prior
        # period disajikan dalam satuan yang sama dalam satu file), jadi
        # nilai yang sama dipasang ke kedua periode.
        period.reporting_unit_multiplier = reporting_unit_multiplier

    return FinancialStatement(
        source_filename=filename,
        entity_name=entity_name,
        periods=[period_current, period_prior],
        parse_warnings=warnings,
    )


def parse_financial_pdf(filepath: str) -> FinancialStatement:
    """
    Parse satu file PDF laporan keuangan tahunan format IDX/OJK.
    Return FinancialStatement dengan 2 periode (current year + prior year).
    """
    rows = _get_all_rows(filepath)
    return parse_financial_statement_rows(os.path.basename(filepath), rows)
