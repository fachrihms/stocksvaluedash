"""
Parser laporan keuangan XLSX format IDX -- file Excel hasil unduhan dari
portal Financial Statements IDX (ekspor instance XBRL ke Excel).

STRUKTUR FILE ASLI (diverifikasi dari file nyata, mis. POWR):
Workbook berisi sheet bernama KODE XBRL, bukan nama laporan:

  - "Context"      : metadata XBRL (kode entitas, tanggal periode)
  - "InlineXBRL"   : peta file HTML -> sheet
  - "1000000"      : General information -- NAMA EMITEN, kode entitas,
                     tanggal periode, dan "Level of rounding" ada di sini
  - "3210000"      : Laporan posisi keuangan (neraca)
  - "3311000"      : Laporan laba rugi & komprehensif
  - "3410000"      : Laporan perubahan ekuitas (+ varian "3410000PY")
  - "3510000"      : Laporan arus kas
  - "36xxxxx"      : Catatan atas laporan keuangan (notes)
  - "hidden"/"Token": data dropdown/internal -- abaikan

Layout baris: label Indonesia di kiri, angka periode (current, prior)
di tengah, label Inggris di KANAN angka. Header periode HANYA berisi
nama konteks ("CurrentYearInstant"/"PriorEndYearInstant") -- tanggal
sebenarnya diambil dari sheet metadata "1000000" (format ISO
"2023-12-31") atau blok periode di sheet "Context".

Karena label & angka bersebelahan di baris yang sama (tanpa wrapping),
pola label `^label` di engine langsung cocok. Yang khas XLSX dan
ditangani di sini:

  1. EMITEN & PERIODE diambil dari sheet "1000000" (General information)
     secara EKSLISIT -- sheet "Context" berisi ~60 baris metadata teknis
     sebelum baris manapun yang bisa dicocokkan, dan "Nama entitas" ada
     di baris ke-6 sheet "1000000" tapi urutan sheet menyisipkan ratusan
     baris lain di depannya, sehingga scan jendela 60-baris engine pasti
     melewatkannya (bug nyata: emiten tak terdeteksi).
  2. SISIRAN CATATAN: semua sheet notes ("36xxxxx") dan laporan
     perubahan ekuitas ("3410000") BERISI baris label yang sama persis
     dengan laporan utama (kas dan setara kas, laba (rugi), dst). Kalau
     ikut disisir, field bisa salah ambil angka dari catatan. Sheet
     yang diproses dibatasi ke laporan utama + info umum.
  3. SATUAN: "Satuan Penuh / Full Amount" = multiplier 1 -- selain
     "Jutaan" dan "Miliaran" yang sudah dikenali engine.
  4. Sheet lama tanpa kode XBRL (label satu kolom, header tanggal di
     atas) tetap didukung lewat fallback konversi baris sederhana.
"""

import os
import re
import zipfile
from datetime import date, datetime

from openpyxl import load_workbook

from app.models.financial import FinancialStatement
from app.parsers.financial_pdf import (
    parse_financial_statement_rows,
    _detect_interim_info,
    _extract_fiscal_year,
)


# Sheet laporan inti berdasarkan kode elemen XBRL IDX (prefix nama sheet).
# Hanya sheet ini yang disisir -- sheet catatan (36xxxxx) dan perubahan
# ekuitas (3410000) sengaja DIKECUALIKAN karena berisi label duplikat
# yang bisa menangkap angka yang salah (lihat docstring modul poin 2).
CORE_SHEET_PREFIXES = ("3210000", "3311000", "3510000")

GENERAL_INFO_SHEET = "1000000"
CONTEXT_SHEET = "Context"

_NumericCell = re.compile(r"^[()\-+\d,.\s]+$")
_HasLetter = re.compile(r"[A-Za-z]")


def _cell_to_text(value) -> str | None:
    """Ubah satu sel Excel jadi teks siap join. None kalau sel kosong."""
    if value is None:
        return None
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, (datetime, date)):
        return value.strftime("%d %B %Y")
    if isinstance(value, float) and value == int(value) and abs(value) < 1e15:
        return str(int(value))
    text = str(value).strip()
    return text if text else None


def _rows_from_sheet(ws) -> list[str]:
    """
    Konversi worksheet jadi baris teks (join sel per baris, urut kolom).
    Angka referensi catatan pendek di kiri label dibuang supaya tidak
    tertangkap sebagai nilai periode.
    """
    rows: list[str] = []
    for excel_row in ws.iter_rows(values_only=True):
        cells = [t for t in (_cell_to_text(v) for v in excel_row) if t is not None]
        if not cells:
            continue

        first_label_idx = next(
            (i for i, c in enumerate(cells) if _HasLetter.search(c)), None
        )
        if first_label_idx is None:
            rows.append(" ".join(cells))
            continue

        while first_label_idx > 0 and _NumericCell.match(cells[0]) and len(cells[0]) <= 4:
            cells.pop(0)
            first_label_idx -= 1

        rows.append(" ".join(cells))
    return rows


def _normalize_meta_value(value) -> str | None:
    """Normalisasi sel metadata: tanggal -> ISO date string, lainnya str."""
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        return value.strftime("%Y-%m-%d")
    text = str(value).strip()
    return text or None


def _currency_code(text: str) -> str | None:
    """
    Ekstrak kode mata uang dari nilai metadata 'Mata uang pelaporan',
    format template IDX: 'Dollar Amerika / USD', 'Rupiah / IDR'. Kode
    ISO (huruf besar 3) diambil dari sisi '/' terakhir; kalau tidak ada,
    kata berhuruf besar 3 huruf di teks dipakai. Return None bila tidak
    teridentifikasi.
    """
    if not text:
        return None
    candidate = text.split("/")[-1].strip().upper()
    if re.fullmatch(r"[A-Z]{3}", candidate):
        return candidate
    match = re.search(r"\b[A-Z]{3}\b", text.upper())
    return match.group() if match else None


def _extract_xlsx_metadata(wb, sheet_name: str) -> dict:
    """
    Metadata eksplisit dari sheet General information (1000000):
    entity_name, entity_code, current/prior period end (ISO), rounding.
    Baris metadata berbentuk "<label> | <nilai> | <label Inggris>" --
    nilai diambil dari sel KEDUA baris yang sel pertamanya cocok label.
    Semua opsional -- pemanggil punya fallback bila tidak ketemu.
    """
    meta: dict = {}
    if sheet_name not in wb.sheetnames:
        return meta
    ws = wb[sheet_name]

    def find(label_regex: str):
        pattern = re.compile(label_regex, re.IGNORECASE)
        for row in ws.iter_rows(values_only=True):
            cells = [c for c in row if c is not None and str(c).strip() != ""]
            if len(cells) < 2:
                continue
            first = str(cells[0]).strip()
            if pattern.match(first):
                return _normalize_meta_value(cells[1])
        return None

    name = find(r"^Nama entitas\b")
    if name:
        meta["entity_name"] = name
    code = find(r"^Kode entitas\b")
    if code:
        meta["entity_code"] = code.upper()
    current = find(r"^Tanggal akhir periode berjalan\b")
    if current:
        meta["current_period"] = current
    prior = find(r"^Tanggal akhir (periode|tahun) sebelumnya\b")
    if prior:
        meta["prior_period"] = prior
    rounding = find(r"^Pembulatan yang digunakan\b")
    if rounding:
        meta["rounding"] = rounding
    currency_raw = find(r"^Mata uang pelaporan\b")
    if currency_raw:
        code = _currency_code(currency_raw)
        if code:
            meta["reporting_currency"] = code
    return meta


def _extract_context_period_ends(wb) -> dict:
    """
    Fallback tanggal periode dari sheet "Context" (metadata XBRL mentah):
    blok "CurrentYearDuration"/"CurrentYearInstant" -> endDate/instant
    pertama = akhir periode berjalan; blok "PriorEndYear*" -> akhir
    periode sebelumnya. Return dict {"current_period": ..., "prior_period": ...}
    dalam format ISO, hanya untuk kunci yang berhasil ditemukan.
    """
    out: dict = {}
    if CONTEXT_SHEET not in wb.sheetnames:
        return out
    ws = wb[CONTEXT_SHEET]

    block = None
    seen: dict[str, str] = {}  # block -> tanggal pertama yang relevan
    for row in ws.iter_rows(values_only=True):
        cells = [c for c in row if c is not None and str(c).strip() != ""]
        if not cells:
            continue
        key = str(cells[0]).strip()
        if re.match(r"^\w+(Duration|Instant)$", key) and len(cells) == 1:
            block = key
            continue
        if block is None:
            continue
        value = _normalize_meta_value(cells[1]) if len(cells) > 1 else None
        if value is None:
            continue
        # Potong bagian waktu bila ada ("2023-12-31 00:00:00" -> "2023-12-31").
        iso = value.split(" ")[0]
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", iso):
            continue
        if key == "endDate" and block.endswith("Duration") and block not in seen:
            seen[block] = iso
        elif key == "instant" and block.endswith("Instant") and block not in seen:
            seen[block] = iso

    current = seen.get("CurrentYearDuration") or seen.get("CurrentYearInstant")
    prior = seen.get("PriorEndYearDuration") or seen.get("PriorEndYearInstant")
    if current:
        out["current_period"] = current
    if prior:
        out["prior_period"] = prior
    return out


def _rounding_multiplier(text: str) -> int:
    """'Satuan Penuh / Full Amount' -> 1; 'Jutaan' -> 1_000_000; 'Miliaran' -> 1_000_000_000."""
    t = text.lower()
    if "penuh" in t or "full amount" in t:
        return 1
    if "miliar" in t or "billion" in t:
        return 1_000_000_000
    return 1_000_000  # jutaan (default paling umum)


def _order_sheets(sheetnames: list[str]) -> list[str]:
    """
    Urutan pemrosesan sheet:
    1) Sheet General information (1000000) lebih dulu.
    2) Laporan inti (3210000/3311000/3510000).
    3) Sheet LAIN (Context/InlineXBRL/notes/sheet lama tanpa kode XBRL)
       sebagai fallback terakhir -- menangani workbook dengan struktur
       berbeda, dengan risiko sisiran catatan yang tetap lebih baik
       daripada gagal total.
    """
    def rank(name: str) -> int:
        if name.startswith(GENERAL_INFO_SHEET):
            return 0
        if name.startswith(CORE_SHEET_PREFIXES):
            return 1
        return 2

    return sorted(sheetnames, key=lambda n: (rank(n), sheetnames.index(n)))


def parse_financial_xlsx(filepath: str) -> FinancialStatement:
    """
    Parse satu file XLSX laporan keuangan format IDX (ekspor Excel dari
    portal XBRL/Financial Statements IDX). Return FinancialStatement dengan
    2 periode (current + prior).
    """
    # File .xlsx yang valid adalah arsip ZIP; download IDX yang gagal/terpotong
    # bisa menyisakan file non-ZIP (contoh nyata: file 165 byte berisi sampah
    # teks "Administrator") -- beri pesan yang jelas, bukan traceback zipfile.
    try:
        wb = load_workbook(filepath, read_only=True, data_only=True, keep_links=False)
    except zipfile.BadZipFile as e:
        size = os.path.getsize(filepath)
        raise ValueError(
            f"File bukan Excel yang valid (kemungkinan download dari IDX gagal/terpotong, "
            f"ukuran hanya {size:,} byte). Coba unduh ulang dari IDX -- dan hapus file rusak "
            f"tersebut dari pilihan upload."
        ) from e
    try:
        # Metadata eksplisit dari sheet General information -- SEBELUM
        # menyisir laporan (bug "emiten tidak terdeteksi": tanpa ini,
        # ratusan baris sheet Context/InlineXBRL menggeser baris
        # "Nama entitas" keluar dari jangkauan scan engine).
        meta = _extract_xlsx_metadata(wb, GENERAL_INFO_SHEET)
        if "current_period" not in meta or "prior_period" not in meta:
            for key, value in _extract_context_period_ends(wb).items():
                meta.setdefault(key, value)

        rows: list[str] = []
        for sheet_name in _order_sheets(wb.sheetnames):
            rows.extend(_rows_from_sheet(wb[sheet_name]))
    finally:
        wb.close()

    result = parse_financial_statement_rows(os.path.basename(filepath), rows)

    # --- Terapkan metadata eksplisit XLSX di atas hasil engine ---
    if meta.get("entity_name"):
        result.entity_name = meta["entity_name"]

    # Label periode XLSX diambil PAKSA dari metadata General information /
    # sheet Context (ISO "2023-12-31"). Dua alasan:
    # 1. Header tabel XLSX asli cuma berisi nama konteks XBRL
    #    ("CurrentYearInstant") -- engine pasti gagal deteksi label tanggal.
    # 2. Deteksi engine yang "berhasil" pun SERING SALAH: teks catatan
    #    audit/direksi di sheet notes menyebut tanggal periode dua kali
    #    (mis. "... 31 Desember 2023 ... 31 Desember 2023") dan cocok
    #    dengan pola header 2-tanggal -- membuat KEDUA periode berlabel
    #    sama (bug nyata POWR). Metadata ini sumber kebenarannya.
    # Sheet lama yang header-nya benar hanya memakai label engine kalau
    # metadata tidak ditemukan sama sekali.
    if meta.get("current_period") and result.periods:
        result.periods[0].period_label = meta["current_period"]
    if meta.get("prior_period") and len(result.periods) > 1:
        result.periods[1].period_label = meta["prior_period"]

    # is_interim/coverage/fiscal_year dihitung engine dari label DETEKSINYA
    # sendiri -- yang barusan kita OVERRIDE. Hitung ulang dari label final
    # supaya laporan tahunan (berakhir 31 Des) tidak salah tertandai
    # interim hanya karena label aslinya tidak dikenali engine.
    for period in result.periods:
        is_interim, coverage_months = _detect_interim_info(rows, period.period_label)
        period.is_interim = is_interim
        period.coverage_months = coverage_months
        period.fiscal_year = _extract_fiscal_year(period.period_label)

    for period in result.periods:
        if meta.get("entity_code"):
            period.entity_code = meta["entity_code"]
        if meta.get("rounding"):
            period.reporting_unit_multiplier = _rounding_multiplier(meta["rounding"])
        if meta.get("reporting_currency"):
            period.reporting_currency = meta["reporting_currency"]

    return result
