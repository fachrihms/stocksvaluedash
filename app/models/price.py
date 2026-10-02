"""Struktur data harga saham harian dari CSV upload."""

from pydantic import BaseModel
from datetime import date
from typing import Optional


class PriceBar(BaseModel):
    """Satu baris harga harian."""

    trade_date: date
    close: float  # "Terakhir"
    open: Optional[float] = None  # "Pembukaan"
    high: Optional[float] = None  # "Tertinggi"
    low: Optional[float] = None  # "Terendah"
    volume: Optional[float] = None  # "Vol."
    change_pct: Optional[float] = None  # "Perubahan%"


class PriceSeries(BaseModel):
    """Kumpulan harga harian dari satu file CSV, terurut dari lama ke baru."""

    source_filename: str
    bars: list[PriceBar] = []
    parse_warnings: list[str] = []
