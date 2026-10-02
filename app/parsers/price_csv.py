"""
Parser CSV harga saham historis format Indonesia (kayak investing.com/IDX).

Format yang didukung, kolom (urutan bebas, dicocokkan by header name):
    Tanggal, Terakhir, Pembukaan, Tertinggi, Terendah, Vol., Perubahan%

Contoh baris data:
    "11/09/2026","2,140","2,130","2,150","2,120","15,23M","0,47%"

Angka format Indonesia butuh normalisasi:
    - "2,140"   -> pemisah ribuan pakai koma       -> 2140
    - "0,47%"   -> koma sebagai desimal DI SINI     -> 0.47
    - "15,23M"  -> "M" = Juta (Million), koma desimal -> 15,230,000

Ambiguitas ini yang bikin parsing CSV harga Indonesia trickier dari yang
kelihatan -- koma bisa berarti pemisah ribuan ATAU desimal tergantung
konteks kolom. Parser ini menangani per-kolom secara eksplisit, bukan
generic number parser, supaya nggak salah interpretasi.
"""

import re
import csv
import io
from datetime import datetime, date
from app.models.price import PriceSeries, PriceBar

# Kemungkinan nama kolom (case-insensitive), untuk fleksibilitas source CSV
COLUMN_ALIASES = {
    "trade_date": ["tanggal", "date"],
    "close": ["terakhir", "close", "price", "harga"],
    "open": ["pembukaan", "open"],
    "high": ["tertinggi", "high"],
    "low": ["terendah", "low"],
    "volume": ["vol.", "vol", "volume"],
    "change_pct": ["perubahan%", "perubahan", "change%", "change"],
}

DATE_FORMATS = ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%d-%m-%Y"]


def _parse_id_date(raw: str) -> date | None:
    raw = raw.strip().strip('"')
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(raw, fmt).date()
        except ValueError:
            continue
    return None


def _parse_id_price(raw: str) -> float | None:
    """Angka harga/nilai biasa: koma = pemisah ribuan. '2,140' -> 2140.0"""
    raw = raw.strip().strip('"')
    if not raw:
        return None
    cleaned = raw.replace(".", "").replace(",", "")
    # Kalau ada titik desimal di format lain (jarang, tapi jaga-jaga)
    try:
        return float(cleaned)
    except ValueError:
        return None


def _parse_id_percent(raw: str) -> float | None:
    """Kolom persen: koma = desimal. '0,47%' -> 0.47. '-1,2%' -> -1.2"""
    raw = raw.strip().strip('"').replace("%", "").strip()
    if not raw:
        return None
    raw = raw.replace(",", ".")
    try:
        return float(raw)
    except ValueError:
        return None


def _parse_id_volume(raw: str) -> float | None:
    """
    Kolom volume: bisa ada suffix Juta (M) atau Ribu (K/rb), koma = desimal
    kalau ada suffix, atau koma = ribuan kalau angka polos.
    '15,23M' -> 15,230,000    '1,5K' -> 1500    '1.234.567' -> 1234567
    """
    raw = raw.strip().strip('"')
    if not raw:
        return None
    multiplier = 1
    if raw[-1].upper() == "M":
        multiplier = 1_000_000
        raw = raw[:-1]
    elif raw[-1].upper() in ("K",):
        multiplier = 1_000
        raw = raw[:-1]
    elif raw.lower().endswith("rb"):
        multiplier = 1_000
        raw = raw[:-2]

    if multiplier > 1:
        # Ada suffix -> koma di sini adalah desimal
        raw = raw.replace(",", ".")
        try:
            return float(raw) * multiplier
        except ValueError:
            return None
    else:
        # Tanpa suffix -> anggap format ribuan biasa
        return _parse_id_price(raw)


def _detect_column_map(header: list[str]) -> dict[str, int]:
    """Cocokkan header CSV ke field kanonis kita, return {field: index}."""
    mapping = {}
    normalized_header = [h.strip().lower() for h in header]
    for field, aliases in COLUMN_ALIASES.items():
        for alias in aliases:
            if alias in normalized_header:
                mapping[field] = normalized_header.index(alias)
                break
    return mapping


def parse_price_csv(filepath_or_content: str, is_content: bool = False) -> PriceSeries:
    """
    Parse CSV harga saham. `filepath_or_content` bisa path file, atau isi
    CSV langsung (string) kalau `is_content=True` (berguna untuk upload
    lewat web -- data datang sebagai string, bukan file di disk).
    """
    import os

    if is_content:
        raw_text = filepath_or_content
        filename = "uploaded.csv"
    else:
        filename = os.path.basename(filepath_or_content)
        with open(filepath_or_content, "r", encoding="utf-8-sig") as f:
            raw_text = f.read()

    reader = csv.reader(io.StringIO(raw_text))
    rows = list(reader)

    warnings: list[str] = []

    if not rows:
        return PriceSeries(source_filename=filename, bars=[], parse_warnings=["File CSV kosong"])

    header = rows[0]
    col_map = _detect_column_map(header)

    if "trade_date" not in col_map or "close" not in col_map:
        warnings.append(
            f"Kolom wajib ('Tanggal' dan 'Terakhir'/'Close') tidak terdeteksi. "
            f"Header terbaca: {header}"
        )
        return PriceSeries(source_filename=filename, bars=[], parse_warnings=warnings)

    bars: list[PriceBar] = []

    for row_num, row in enumerate(rows[1:], start=2):
        if not row or all(not c.strip() for c in row):
            continue  # skip baris kosong

        try:
            trade_date = _parse_id_date(row[col_map["trade_date"]])
            close = _parse_id_price(row[col_map["close"]])

            if trade_date is None or close is None:
                warnings.append(f"Baris {row_num}: tanggal/harga tidak valid, dilewati: {row}")
                continue

            bar = PriceBar(
                trade_date=trade_date,
                close=close,
                open=_parse_id_price(row[col_map["open"]]) if "open" in col_map and col_map["open"] < len(row) else None,
                high=_parse_id_price(row[col_map["high"]]) if "high" in col_map and col_map["high"] < len(row) else None,
                low=_parse_id_price(row[col_map["low"]]) if "low" in col_map and col_map["low"] < len(row) else None,
                volume=_parse_id_volume(row[col_map["volume"]]) if "volume" in col_map and col_map["volume"] < len(row) else None,
                change_pct=_parse_id_percent(row[col_map["change_pct"]]) if "change_pct" in col_map and col_map["change_pct"] < len(row) else None,
            )
            bars.append(bar)
        except (IndexError, ValueError) as e:
            warnings.append(f"Baris {row_num}: gagal diparse ({e}), dilewati: {row}")
            continue

    # Urutkan dari tanggal paling lama ke paling baru -- CSV dari
    # investing.com/IDX biasanya justru terbalik (terbaru duluan)
    bars.sort(key=lambda b: b.trade_date)

    return PriceSeries(source_filename=filename, bars=bars, parse_warnings=warnings)
