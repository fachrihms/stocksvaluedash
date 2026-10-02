"""
Entry point FastAPI. Serve API upload (lapkeu PDF/XLSX + CSV harga) dan
static frontend (index.html) dalam satu server -- paling praktis untuk
development lokal.

Jalankan: uvicorn app.main:app --reload
Buka: http://localhost:8000
"""

import hashlib
import tempfile
import os
from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import HTMLResponse, JSONResponse, Response
from fastapi.middleware.cors import CORSMiddleware

from app.parsers.financial_pdf import parse_financial_pdf
from app.parsers.financial_xlsx import parse_financial_xlsx
from app.parsers.price_csv import parse_price_csv
from app.utils.ratios import build_ratio_summary, merge_periods, build_multi_year_summary, RATIO_BENCHMARKS
from app.utils import spreadsheet_db
from app.utils import prediction as prediction_utils
from app.models.financial import FinancialPeriod

app = FastAPI(title="Stock Valuation Dashboard API")

# CORS dibuka lebar untuk development lokal. Kalau nanti frontend di-deploy
# terpisah (misal Netlify) dan backend di server lain (Railway/Render),
# ganti allow_origins ke domain frontend yang spesifik -- jangan biarkan "*"
# di production karena endpoint ini menerima upload file.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", response_class=HTMLResponse)
async def serve_frontend():
    frontend_path = os.path.join(os.path.dirname(__file__), "..", "static", "index.html")
    with open(frontend_path, "r", encoding="utf-8") as f:
        return f.read()


@app.get("/rasio", response_class=HTMLResponse)
async def serve_ratio_glossary():
    """Halaman terpisah berisi kamus rasio (rumus + sumber baris lapkeu +
    cara baca), dipindah dari dashboard agar angka/grafik tidak bercampur
    dengan materi penjelasan."""
    glossary_path = os.path.join(os.path.dirname(__file__), "..", "static", "rasio.html")
    with open(glossary_path, "r", encoding="utf-8") as f:
        return f.read()


FINANCIAL_EXTENSIONS = (".pdf", ".xlsx")


def parse_financial_statement_file(filepath: str, filename: str):
    """
    Dispatch parser laporan keuangan sesuai ekstensi file:
    .pdf -> parser PDF (pdfplumber), .xlsx -> parser Excel (openpyxl).
    Keduanya return FinancialStatement dengan perilaku identik (engine
    label-matching yang sama, lihat parse_financial_statement_rows).
    """
    lower = filename.lower()
    if lower.endswith(".pdf"):
        return parse_financial_pdf(filepath)
    if lower.endswith(".xlsx"):
        return parse_financial_xlsx(filepath)
    raise ValueError(f"Format file tidak didukung: {filename}")


def _validate_financial_filename(filename: str) -> bool:
    return filename.lower().endswith(FINANCIAL_EXTENSIONS)


@app.post("/api/parse-financial-pdf")
async def api_parse_financial_pdf(file: UploadFile = File(...)):
    """Upload satu PDF laporan keuangan, return angka-angka terstruktur."""
    if not file.filename.lower().endswith(".pdf"):
        raise HTTPException(400, "File harus berformat PDF")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        result = parse_financial_pdf(tmp_path)
    except Exception as e:
        raise HTTPException(500, f"Gagal parsing PDF: {e}")
    finally:
        os.unlink(tmp_path)

    return JSONResponse(result.model_dump())


@app.post("/api/parse-financial-xlsx")
async def api_parse_financial_xlsx(file: UploadFile = File(...)):
    """
    Upload satu XLSX laporan keuangan (unduhan Excel dari portal Financial
    Statements IDX), return angka-angka terstruktur -- respons identik dengan
    /api/parse-financial-pdf.
    """
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(400, "File harus berformat XLSX (Excel)")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        result = parse_financial_xlsx(tmp_path)
    except Exception as e:
        raise HTTPException(500, f"Gagal parsing XLSX: {e}")
    finally:
        os.unlink(tmp_path)

    return JSONResponse(result.model_dump())


@app.post("/api/parse-financial-statement")
async def api_parse_financial_statement(file: UploadFile = File(...)):
    """
    Upload satu file laporan keuangan (PDF ATAU XLSX), auto-deteksi format
    dari ekstensi. Endpoint gabungan yang direkomendasikan untuk integrasi
    baru; endpoint lama per-format tetap ada untuk kompatibilitas.
    """
    if not _validate_financial_filename(file.filename):
        raise HTTPException(400, "File harus berformat PDF atau XLSX (Excel)")

    suffix = ".pdf" if file.filename.lower().endswith(".pdf") else ".xlsx"
    with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        result = parse_financial_statement_file(tmp_path, file.filename)
    except Exception as e:
        raise HTTPException(500, f"Gagal parsing laporan keuangan: {e}")
    finally:
        os.unlink(tmp_path)

    return JSONResponse(result.model_dump())


@app.post("/api/parse-price-csv")
async def api_parse_price_csv(file: UploadFile = File(...)):
    """Upload satu CSV harga saham historis, return time series terstruktur."""
    if not file.filename.lower().endswith(".csv"):
        raise HTTPException(400, "File harus berformat CSV")

    content = await file.read()
    text = content.decode("utf-8-sig")

    try:
        result = parse_price_csv(text, is_content=True)
    except Exception as e:
        raise HTTPException(500, f"Gagal parsing CSV: {e}")

    return JSONResponse(result.model_dump(mode="json"))


@app.post("/api/analyze")
async def api_analyze(
    financial_pdfs: list[UploadFile] = File(...),
    price_csv: UploadFile | None = File(None),
    shares_outstanding: float | None = None,
    exchange_rate_to_idr: float | None = None,
    spreadsheet_url: str | None = Query(
        default=None,
        description="Link Google Sheets/CSV opsional sebagai database tambahan. "
                    "Periode lama diambil dari sini supaya tidak perlu upload ulang.",
    ),
    use_database: bool = Query(
        default=True,
        description="False untuk mematikan cache database lokal (analisis murni dari file upload).",
    ),
    save_to_database: bool = Query(
        default=True,
        description="False agar hasil TIDAK disimpan ke database lokal.",
    ),
):
    """
    Endpoint utama dashboard: upload SATU ATAU LEBIH file lapkeu (PDF atau
    XLSX, boleh campur; + opsional CSV harga saham), return hasil parse
    LENGKAP dengan semua rasio terhitung, termasuk tren multi-tahun kalau
    ada >1 file.

    Kenapa terima banyak file: satu laporan (tahunan) cuma punya 2 periode
    (current + prior). Untuk tren yang lebih panjang (3+ tahun), user perlu
    upload beberapa laporan sekaligus (mis. laporan 2023, 2024, 2025) --
    periode yang overlap antar file di-dedup otomatis (lihat merge_periods
    di app/utils/ratios.py).

    `shares_outstanding` opsional (jumlah saham beredar, dalam LEMBAR
    PENUH bukan jutaan) -- dibutuhkan untuk hitung PER/PBV/dividend yield.

    `exchange_rate_to_idr` opsional (kurs mata uang laporan -> Rupiah,
    mis. 16250 untuk USD). WAJIB kalau laporan disajikan dalam mata uang
    selain Rupiah -- sebagian kecil emiten IDX (mis. POWR) ber-USD, dan
    EPS/BVPS-nya tidak boleh dibandingkan langsung dengan harga saham
    Rupiah (bug nyata: PER 222.826x). Guard di bawah menolak request
    tanpa kurs dengan pesan jelas, alih-alih menghasilkan angka tameng.
    """
    for f in financial_pdfs:
        if not _validate_financial_filename(f.filename):
            raise HTTPException(400, f"'{f.filename}' bukan file PDF atau XLSX")

    all_periods = []
    per_file_warnings: dict[str, list[str]] = {}
    entity_name = None
    cache_hits = 0

    for f in financial_pdfs:
        suffix = ".pdf" if f.filename.lower().endswith(".pdf") else ".xlsx"
        content = await f.read()
        file_hash = hashlib.sha256(content).hexdigest() if use_database else None

        # Cache file: upload file yang SAMA persis tidak di-parse ulang
        # (parse PDF = tahap paling lambat). Cache disimpan di SQLite lokal.
        if file_hash:
            cached = spreadsheet_db.get_cached_file(file_hash)
            if cached is not None:
                cached_entity, cached_periods = cached
                all_periods.extend(cached_periods)
                if entity_name is None and cached_entity:
                    entity_name = cached_entity
                cache_hits += 1
                continue

        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            tmp.write(content)
            tmp_path = tmp.name

        try:
            result = parse_financial_statement_file(tmp_path, f.filename)
        except Exception as e:
            raise HTTPException(500, f"Gagal parsing '{f.filename}': {e}")
        finally:
            os.unlink(tmp_path)

        # Tandai sumber file per periode (penting untuk panel "cara membaca").
        for p in result.periods:
            if not p.source_filename:
                p.source_filename = f.filename
        all_periods.extend(result.periods)
        if file_hash:
            spreadsheet_db.put_cached_file(file_hash, result.entity_name, result.periods)
        if result.parse_warnings:
            per_file_warnings[f.filename] = result.parse_warnings
        if entity_name is None and result.entity_name:
            entity_name = result.entity_name

    if not all_periods:
        raise HTTPException(422, "Tidak ada periode yang berhasil di-parse dari file yang diupload")

    # --- Sumber periode tambahan: database lokal + spreadsheet opsional ---
    # Ini yang bikin analisis "lebih cepat + bisa update": periode tahun lama
    # yang sudah tersimpan tidak perlu di-upload ulang. Cukup upload 1
    # laporan terbaru, sisanya digabung otomatis dari sini lalu di-dedup
    # via merge_periods (periode dari file upload selalu menang kalau ada
    # duplikat tanggal -- lihat prioritas di merge_periods).
    db_periods_used = 0
    spreadsheet_periods_used = 0
    spreadsheet_error: str | None = None
    if use_database and entity_name:
        try:
            stored = spreadsheet_db.get_stored_periods(entity_name)
            db_periods_used = len(stored)
            all_periods = stored + all_periods
        except Exception:
            pass
    if spreadsheet_url:
        sheet_periods, sheet_err = spreadsheet_db.fetch_spreadsheet_periods(spreadsheet_url)
        if sheet_err:
            spreadsheet_error = sheet_err
        elif sheet_periods:
            # Kalau sheet berisi banyak emiten, filter ke entitas ini bila
            # namanya cocok; kalau sheet tidak mencantumkan nama entitas
            # (kolom kosong), tetap dipakai sebagai fallback.
            spreadsheet_periods_used = len(sheet_periods)
            all_periods = sheet_periods + all_periods

    merged_periods = merge_periods(all_periods)

    price_result = None
    if price_csv is not None:
        if not price_csv.filename.lower().endswith(".csv"):
            raise HTTPException(400, "price_csv harus berformat CSV")
        price_content = await price_csv.read()
        price_text = price_content.decode("utf-8-sig")
        try:
            price_result = parse_price_csv(price_text, is_content=True)
        except Exception as e:
            raise HTTPException(500, f"Gagal parsing CSV: {e}")

    # Rasio "utama" (dashboard atas) dihitung dari 2 periode TAHUNAN PENUH
    # paling baru dalam deret gabungan -- BUKAN sekadar 2 elemen terakhir
    # array, karena elemen terakhir bisa jadi periode interim/kuartalan
    # (mis. laporan Q2 2026 yang baru mencakup 6 bulan). Membandingkan
    # interim vs tahunan penuh langsung akan menghasilkan growth rate yang
    # menyesatkan (12 bulan vs 6 bulan bukan perbandingan yang adil).
    annual_periods = [p for p in merged_periods if not p.is_interim]
    interim_periods = [p for p in merged_periods if p.is_interim]

    if len(annual_periods) < 2:
        raise HTTPException(
            422,
            "Minimal butuh 2 periode TAHUNAN PENUH (31 Desember) untuk hitung "
            "rasio pertumbuhan utama. Periode interim/kuartalan saja tidak cukup."
        )

    current, prior = annual_periods[-1], annual_periods[-2]

    # Guard mata uang: laporan non-IDR butuh kurs sebelum rasio per-share
    # (PER/PBV/yield) boleh dihitung melawan harga saham Rupiah.
    reporting_currency = (current.reporting_currency or "IDR").upper()
    if reporting_currency != "IDR" and not exchange_rate_to_idr:
        raise HTTPException(
            422,
            f"Laporan ini disajikan dalam {reporting_currency} (bukan Rupiah). "
            f"Isi kurs {reporting_currency}->IDR (mis. 16250) agar PER/PBV/dividend "
            f"yield bisa dihitung terhadap harga saham Rupiah."
        )

    ratios = build_ratio_summary(
        current, prior, price_result, shares_outstanding, exchange_rate_to_idr
    )
    multi_year = build_multi_year_summary(merged_periods)

    # Simpan hasil gabungan (termasuk periode terbaru) ke database lokal
    # supaya analisis berikutnya tinggal upload file baru saja.
    stored_now = 0
    if save_to_database and entity_name:
        try:
            stored_now = spreadsheet_db.upsert_periods(entity_name, merged_periods)
        except Exception:
            stored_now = 0

    # Info periode interim/kuartalan terbaru (kalau ada) ditampilkan
    # terpisah di frontend sebagai angka kecil dengan tanggal jelas --
    # BUKAN dimasukkan ke rantai growth rate utama. Ini beda dari
    # `multi_year` yang menampilkan seluruh deret (termasuk interim, tapi
    # dengan growth_pct=None untuk entry interim -- lihat build_multi_year_summary).
    latest_interim = None
    if interim_periods:
        latest = interim_periods[-1]
        latest_interim = {
            "period_label": latest.period_label,
            "coverage_months": latest.coverage_months,
            "revenue": latest.revenue,
            "net_income_parent": latest.net_income_parent,
            "eps": latest.basic_eps,
        }

    return JSONResponse({
        "entity_name": entity_name,
        "period_labels": [p.period_label for p in merged_periods],
        "multi_year": multi_year,
        "latest_interim": latest_interim,
        "analysis_context": {
            "current_period": current.period_label,
            "current_source_filename": current.source_filename,
            "prior_period": prior.period_label,
            "prior_source_filename": prior.source_filename,
            "trend_period_count": len(merged_periods),
        },
        "price": price_result.model_dump(mode="json") if price_result else None,
        "ratios": ratios,
        "benchmarks": RATIO_BENCHMARKS,
        "parse_warnings_by_file": per_file_warnings,
        "database": {
            "cache_hits": cache_hits,
            "db_periods_used": db_periods_used,
            "spreadsheet_periods_used": spreadsheet_periods_used,
            "spreadsheet_error": spreadsheet_error,
            "stored_now": stored_now,
            "from_database": db_periods_used + spreadsheet_periods_used,
        },
    })


@app.get("/api/benchmarks")
async def api_benchmarks():
    """Patokan benchmark rasio (sumber tunggal untuk view Legenda)."""
    return JSONResponse(RATIO_BENCHMARKS)


@app.post("/api/predict")
async def api_predict(
    price_csv: UploadFile = File(...),
    horizon_days: int = Query(
        default=30, ge=5, le=120,
        description="Jumlah hari bursa ke depan yang diprediksi (5-120).",
    ),
):
    """Prediksi harga saham sederhana dari CSV historis.

    Membandingkan 4 model baseline (regresi linier, kuadratik, SMA-20,
    exponential smoothing) dengan evaluasi holdout yang transparan.
    Ini BUKAN nasihat keuangan -- akurasi model deret waktu sederhana
    untuk harga saham umumnya rendah; lihat skor holdout (RMSE/MAPE)
    sebelum mempercayai forecast mana pun.
    """
    if not price_csv.filename.lower().endswith(".csv"):
        raise HTTPException(400, "price_csv harus berformat CSV")
    content = await price_csv.read()
    try:
        prices = parse_price_csv(content.decode("utf-8-sig"), is_content=True)
    except Exception as e:
        raise HTTPException(500, f"Gagal parsing CSV: {e}")

    bars = prices.bars
    if len(bars) < 30:
        raise HTTPException(
            422,
            f"Butuh minimal 30 baris data historis untuk prediksi "
            f"(dapat {len(bars)}). Unduh rentang tanggal yang lebih panjang.",
        )

    closes = [b.close for b in bars]
    result = prediction_utils.evaluate_and_forecast(closes, horizon_days)

    # Kirim histori secukupnya untuk grafik (250 titik terakhir -- CSV
    # investing.com multi-tahun bisa ribuan baris, tidak perlu semua).
    tail = bars[-250:]
    future_dates = prediction_utils.future_business_dates(
        bars[-1].trade_date, horizon_days
    )
    last_close = closes[-1]
    summary = {}
    for model, fc in result["forecasts"].items():
        pct = (fc[-1] / last_close - 1) * 100 if last_close else None
        summary[model] = {
            "last_forecast": fc[-1],
            "change_pct_to_horizon": pct,
        }

    return JSONResponse({
        "source_filename": prices.source_filename,
        "history": {
            "dates": [str(b.trade_date) for b in tail],
            "closes": [b.close for b in tail],
            "shown_points": len(tail),
            "total_points": len(bars),
            "last_close": last_close,
            "last_date": str(bars[-1].trade_date),
        },
        "future_dates": future_dates,
        "horizon_days": horizon_days,
        "models": prediction_utils.MODEL_META,
        "forecasts": result["forecasts"],
        "forecast_summary": summary,
        "holdout_days": result["holdout_days"],
        "metrics": result["metrics"],
        "best_model": result["best_model"],
    })


@app.get("/api/database/status")
async def api_database_status(entity_name: str | None = None):
    """Statistik database lokal (jumlah periode/file tersimpan per emiten)."""
    return JSONResponse(spreadsheet_db.database_stats(entity_name))


@app.get("/api/database/export")
async def api_database_export(
    format: str = Query(default="csv", description="csv, xlsx, atau json"),
    entity_name: str | None = Query(default=None),
):
    """Unduh seluruh hasil analisis tersimpan untuk ditempel/diimpor ke
    spreadsheet (Google Sheets: File > Import > Upload). Ini arah tulis
    database -> spreadsheet."""
    fmt = (format or "csv").lower()
    periods = spreadsheet_db.get_stored_periods(entity_name)
    # Urutkan kronologis supaya enak dibaca di sheet.
    try:
        from app.utils.period_dates import parse_period_label_date
        periods.sort(key=lambda p: parse_period_label_date(p.period_label) or __import__("datetime").date.max)
    except Exception:
        pass

    if fmt == "json":
        return JSONResponse([
            p.model_dump(mode="json") | ({"entity_name": entity_name} if entity_name else {})
            for p in periods
        ])
    if fmt == "xlsx":
        content = spreadsheet_db.periods_to_xlsx_bytes(periods, entity_name)
        return Response(
            content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": "attachment; filename=analysis_db.xlsx"},
        )
    content = spreadsheet_db.periods_to_csv(periods, entity_name)
    return Response(
        content,
        media_type="text/csv; charset=utf-8",
        headers={"Content-Disposition": "attachment; filename=analysis_db.csv"},
    )


@app.post("/api/database/import-spreadsheet")
async def api_database_import_spreadsheet(payload: dict):
    """Ambil periode dari link spreadsheet lalu simpan ke database lokal.
    Body: {"spreadsheet_url": "...", "entity_name": "OPSIONAL"}.
    Return jumlah periode yang berhasil disimpan."""
    url = (payload or {}).get("spreadsheet_url", "")
    entity_name = (payload or {}).get("entity_name") or None
    if not url:
        raise HTTPException(400, "spreadsheet_url wajib diisi")
    periods, err = spreadsheet_db.fetch_spreadsheet_periods(url)
    if err:
        raise HTTPException(422, err)
    if entity_name:
        stored = spreadsheet_db.upsert_periods(entity_name, periods)
    else:
        # Tanpa nama entitas: kelompokkan per entity_code bila ada,
        # sisanya masuk bucket "UNKNOWN".
        stored = spreadsheet_db.upsert_periods("UNKNOWN", periods)
    return JSONResponse({
        "imported_periods": len(periods),
        "stored_now": stored,
        "entity_name": entity_name or "UNKNOWN",
    })


@app.delete("/api/database/clear")
async def api_database_clear(entity_name: str | None = None):
    """Hapus cache database lokal (semua atau per emiten)."""
    return JSONResponse(spreadsheet_db.clear_database(entity_name))
