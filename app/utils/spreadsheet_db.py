"""
Database lokal + sinkronisasi spreadsheet opsional untuk hasil analisis.

Tujuan (request user): menyimpan seluruh hasil analisis di spreadsheet/
database supaya:
1. Analisis berikutnya LEBIH CEPAT -- periode tahun lama tidak perlu
   di-parse ulang dari PDF/XLSX, cukup diambil dari database/spreadsheet.
2. Bisa UPDATE inkremental -- upload 1 laporan terbaru saja, tahun-tahun
   lama otomatis digabung dari database.

Desain:
- Database lokal: SQLite (stdlib, tanpa dependensi baru) di ./data/analysis.db
  - tabel `periods`: satu baris per (entity, period_label) berisi seluruh
    field FinancialPeriod sebagai JSON.
  - tabel `files`: cache hasil parse per sha256 file, supaya upload file
    yang SAMA persis tidak di-parse ulang (parse PDF = bagian paling lambat).
- Spreadsheet: format CSV/XLSX kanonis (lihat SPREADSHEET_COLUMNS) yang bisa
  diimpor/diekspor. User tempel link Google Sheets (share "anyone with link")
  sebagai `spreadsheet_url` opsional -- backend fetch CSV-nya TANPA auth
  (gviz export) dan menjadikannya sumber periode tambahan.
  Menulis LANGSUNG ke Google Sheet user butuh OAuth, jadi arah tulisnya
  lewat tombol unduh CSV/XLSX di dashboard (paste/import sekali ke Sheet).
  Roundtrip ini tetap memenuhi "disimpen di spreadsheet + bisa update".

Format kolom spreadsheet = field FinancialPeriod yang relevan (angka masih
dalam SATUAN LAPORAN ASLI, sama seperti model -- konversi satuan tetap
ditangani ratios.py via reporting_unit_multiplier).
"""

import csv
import io
import json
import os
import re
import sqlite3
import time
import urllib.request
from typing import Optional

from app.models.financial import FinancialPeriod

DB_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "data", "analysis.db")

SPREADSHEET_COLUMNS = [
    "entity_name",
    "period_label",
    "fiscal_year",
    "is_interim",
    "coverage_months",
    "source_filename",
    "reporting_unit_multiplier",
    "reporting_currency",
    "total_assets",
    "total_current_assets",
    "total_liabilities",
    "total_current_liabilities",
    "total_equity",
    "total_equity_parent",
    "cash_and_equivalents",
    "revenue",
    "profit_before_tax",
    "net_income_parent",
    "net_income_total",
    "basic_eps",
    "interest_expense",
    "operating_cash_flow",
    "investing_cash_flow",
    "financing_cash_flow",
    "capex",
    "dividends_paid",
]

_FLOAT_FIELDS = set(SPREADSHEET_COLUMNS) - {
    "entity_name", "period_label", "source_filename", "reporting_currency",
}
_INT_FIELDS = {"fiscal_year", "coverage_months", "reporting_unit_multiplier"}


# ---------- SQLite ----------

def _connect() -> sqlite3.Connection:
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        """CREATE TABLE IF NOT EXISTS periods (
            entity_name TEXT NOT NULL,
            period_label TEXT NOT NULL,
            is_interim INTEGER NOT NULL DEFAULT 0,
            data_json TEXT NOT NULL,
            updated_at REAL NOT NULL,
            PRIMARY KEY (entity_name, period_label, is_interim)
        )"""
    )
    conn.execute(
        """CREATE TABLE IF NOT EXISTS files (
            file_hash TEXT PRIMARY KEY,
            entity_name TEXT,
            data_json TEXT NOT NULL,
            created_at REAL NOT NULL
        )"""
    )
    conn.commit()
    return conn


def upsert_periods(entity_name: str, periods: list[FinancialPeriod]) -> int:
    """Simpan/UPDATE periode ke DB lokal. Return jumlah baris ditulis."""
    if not entity_name or not periods:
        return 0
    conn = _connect()
    now = time.time()
    count = 0
    try:
        for p in periods:
            conn.execute(
                "INSERT OR REPLACE INTO periods "
                "(entity_name, period_label, is_interim, data_json, updated_at) "
                "VALUES (?, ?, ?, ?, ?)",
                (
                    entity_name,
                    p.period_label,
                    1 if p.is_interim else 0,
                    p.model_dump_json(),
                    now,
                ),
            )
            count += 1
        conn.commit()
    finally:
        conn.close()
    return count


def get_stored_periods(entity_name: Optional[str] = None) -> list[FinancialPeriod]:
    """Ambil periode tersimpan (filter entitas opsional). Gagal parse -> skip."""
    conn = _connect()
    try:
        if entity_name:
            rows = conn.execute(
                "SELECT data_json FROM periods WHERE entity_name = ?", (entity_name,)
            ).fetchall()
        else:
            rows = conn.execute("SELECT data_json FROM periods").fetchall()
    finally:
        conn.close()
    out = []
    for (blob,) in rows:
        try:
            out.append(FinancialPeriod.model_validate_json(blob))
        except Exception:
            continue
    return out


def get_cached_file(file_hash: str):
    """Return (entity_name, [FinancialPeriod]) kalau file pernah di-parse."""
    conn = _connect()
    try:
        row = conn.execute(
            "SELECT entity_name, data_json FROM files WHERE file_hash = ?",
            (file_hash,),
        ).fetchone()
    finally:
        conn.close()
    if not row:
        return None
    entity_name, blob = row
    try:
        items = json.loads(blob)
        return entity_name, [FinancialPeriod.model_validate(d) for d in items]
    except Exception:
        return None


def put_cached_file(file_hash: str, entity_name: Optional[str],
                     periods: list[FinancialPeriod]) -> None:
    conn = _connect()
    try:
        conn.execute(
            "INSERT OR REPLACE INTO files (file_hash, entity_name, data_json, created_at)"
            " VALUES (?, ?, ?, ?)",
            (
                file_hash,
                entity_name,
                json.dumps([p.model_dump(mode="json") for p in periods]),
                time.time(),
            ),
        )
        conn.commit()
    finally:
        conn.close()


def database_stats(entity_name: Optional[str] = None) -> dict:
    conn = _connect()
    try:
        if entity_name:
            n_periods = conn.execute(
                "SELECT COUNT(*) FROM periods WHERE entity_name = ?", (entity_name,)
            ).fetchone()[0]
        else:
            n_periods = conn.execute("SELECT COUNT(*) FROM periods").fetchone()[0]
        n_files = conn.execute("SELECT COUNT(*) FROM files").fetchone()[0]
        entities = [r[0] for r in conn.execute(
            "SELECT DISTINCT entity_name FROM periods").fetchall()]
    finally:
        conn.close()
    return {
        "stored_periods": n_periods,
        "cached_files": n_files,
        "entities": entities,
    }


def clear_database(entity_name: Optional[str] = None) -> dict:
    conn = _connect()
    try:
        if entity_name:
            p = conn.execute(
                "DELETE FROM periods WHERE entity_name = ?", (entity_name,)).rowcount
            f = conn.execute(
                "DELETE FROM files WHERE entity_name = ?", (entity_name,)).rowcount
        else:
            p = conn.execute("DELETE FROM periods").rowcount
            f = conn.execute("DELETE FROM files").rowcount
        conn.commit()
    finally:
        conn.close()
    return {"deleted_periods": p, "deleted_file_caches": f}


# ---------- Spreadsheet (CSV/XLSX kanonis) ----------

def periods_to_rows(periods: list[FinancialPeriod],
                    entity_name: Optional[str] = None) -> list[dict]:
    rows = []
    for p in periods:
        d = p.model_dump(mode="json")
        row = {}
        for col in SPREADSHEET_COLUMNS:
            if col == "entity_name":
                row[col] = entity_name or d.get("entity_code") or ""
            elif col == "is_interim":
                row[col] = "TRUE" if d.get("is_interim") else "FALSE"
            else:
                v = d.get(col)
                row[col] = "" if v is None else v
        rows.append(row)
    return rows


def periods_to_csv(periods: list[FinancialPeriod],
                   entity_name: Optional[str] = None) -> str:
    buf = io.StringIO()
    w = csv.DictWriter(buf, fieldnames=SPREADSHEET_COLUMNS)
    w.writeheader()
    for row in periods_to_rows(periods, entity_name):
        w.writerow(row)
    return buf.getvalue()


def periods_to_xlsx_bytes(periods: list[FinancialPeriod],
                           entity_name: Optional[str] = None) -> bytes:
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "analysis_db"
    ws.append(SPREADSHEET_COLUMNS)
    for row in periods_to_rows(periods, entity_name):
        ws.append([row[c] for c in SPREADSHEET_COLUMNS])
    # Baris pertama = header panduan: user boleh edit angka lalu import ulang.
    ws.freeze_panes = "A2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _to_float(v) -> Optional[float]:
    if v is None or v == "":
        return None
    try:
        return float(str(v).replace(" ", "").replace(",", ""))
    except ValueError:
        return None


def _to_int(v) -> Optional[int]:
    f = _to_float(v)
    return None if f is None else int(f)


def rows_to_periods(rows: list[dict],
                    default_entity: Optional[str] = None) -> list[FinancialPeriod]:
    periods = []
    for r in rows:
        # Header case-insensitive + toleran kolom tambahan dari edit manual.
        norm = {(k or "").strip().lower(): v for k, v in r.items()}
        label = str(norm.get("period_label", "") or "").strip()
        if not label:
            continue
        is_interim = str(norm.get("is_interim", "") or "").strip().upper() in (
            "TRUE", "1", "YES", "YA")
        data: dict = {
            "period_label": label,
            "is_interim": is_interim,
            "source_filename": str(norm.get("source_filename", "") or "") or None,
            "reporting_currency": str(norm.get("reporting_currency", "") or "") or None,
        }
        for col in _FLOAT_FIELDS - {"is_interim"}:
            if col in norm and str(norm[col]).strip() != "":
                data[col] = _to_float(norm[col])
        for col in _INT_FIELDS:
            if col in norm and str(norm[col]).strip() != "":
                data[col] = _to_int(norm[col])
        if not data.get("reporting_unit_multiplier"):
            data["reporting_unit_multiplier"] = 1_000_000
        try:
            periods.append(FinancialPeriod.model_validate(data))
        except Exception:
            continue
    return periods


def csv_text_to_periods(text: str) -> list[FinancialPeriod]:
    reader = csv.DictReader(io.StringIO(text))
    return rows_to_periods(list(reader or []))


# ---------- Google Sheets link ----------

def extract_sheet_id(link: str) -> Optional[str]:
    """Terima ID polos, URL /d/{id}/, URL export, atau URL publish."""
    if not link:
        return None
    link = link.strip()
    m = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", link)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9-_]+)", link)
    if m:
        return m.group(1)
    if re.fullmatch(r"[a-zA-Z0-9-_]{20,}", link):
        return link
    return None


def spreadsheet_url_to_csv_urls(link: str) -> list[str]:
    """Kandidat URL CSV untuk sebuah link spreadsheet (dicoba berurutan)."""
    link = (link or "").strip()
    if not link:
        return []
    # Kalau user sudah memberi link CSV/export langsung, pakai itu dulu.
    if "tqx=out:csv" in link or "format=csv" in link or link.endswith(".csv"):
        sheet_id = extract_sheet_id(link)
        urls = [link]
    else:
        sheet_id = extract_sheet_id(link)
        urls = []
    if sheet_id:
        urls.append(
            f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv")
        urls.append(
            f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv")
    # Bukan link Google Sheets (mis. raw CSV di hosting lain): coba langsung.
    if not urls and link.startswith("http"):
        urls.append(link)
    return urls


def fetch_spreadsheet_periods(link: str, timeout: int = 15) -> tuple[list, Optional[str]]:
    """Fetch periode dari link spreadsheet. Return (periods, error_message)."""
    urls = spreadsheet_url_to_csv_urls(link)
    if not urls:
        return [], "Link spreadsheet tidak dikenali -- pakai link share Google Sheets atau URL CSV langsung."
    last_err = None
    for url in urls:
        try:
            req = urllib.request.Request(
                url, headers={"User-Agent": "stock-dashboard/1.0"})
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                raw = resp.read().decode("utf-8-sig")
            periods = csv_text_to_periods(raw)
            if periods:
                return periods, None
            last_err = "CSV terbaca tapi tidak ada baris periode valid (kolom period_label kosong?)."
        except Exception as e:  # noqa: BLE001 -- pesan error diteruskan ke user
            last_err = f"{type(e).__name__}: {e}"
            continue
    hint = ("Pastikan sheet di-share 'Anyone with the link (Viewer)' "
            "atau pakai File > Share > Publish to web (CSV).")
    return [], f"Gagal mengambil spreadsheet. {last_err or ''} {hint}"
