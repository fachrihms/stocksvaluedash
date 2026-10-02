"""
Prediksi harga saham sederhana dari data historis (CSV harga).

Bukan ramalan akurat -- ini baseline edukatif supaya user bisa
membandingkan beberapa model time-series klasik dalam satu grafik.
Semua model diimplementasikan murni dengan numpy (tanpa dependensi ML
baru) agar instalasi tetap ringan:

- linear: regresi linier (OLS) tren waktu vs harga penutupan.
- quadratic: regresi polinomial derajat 2 (menangkap kelengkungan tren).
- sma: Simple Moving Average 20 hari, forecast datar di level SMA terakhir.
- exp_smooth: Simple Exponential Smoothing (alpha=0.35), forecast datar
  di level terakhir.

Evaluasi: tiap model diuji dengan holdout -- N titik terakhir disisihkan
sebagai "masa depan" pura-pura, model dilatih di sisanya, lalu diukur
MAE/RMSE/MAPE. Model dengan RMSE terkecil jadi `best_model`. Ini
transparan: user melihat sendiri seberapa (tidak) akurat tiap model di
data holdout sebelum percaya forecast-nya.
"""

from datetime import date, timedelta

import numpy as np

MODEL_META = {
    "linear": {
        "label": "Regresi Linier",
        "desc": "Garis tren lurus (OLS) dari seluruh histori -- bagus kalau tren stabil, gagal kalau tren berbelok.",
    },
    "quadratic": {
        "label": "Regresi Kuadratik",
        "desc": "Kurva parabola (derajat 2) -- menangkap akselerasi/perlambatan tren, tapi rawan overfit di ujung.",
    },
    "sma": {
        "label": "Rata-rata Bergerak (SMA-20)",
        "desc": "Rata-rata 20 hari terakhir, forecast datar -- konservatif, cocok untuk harga sideways.",
    },
    "exp_smooth": {
        "label": "Exponential Smoothing",
        "desc": "Rata-rata berbobot eksponensial (alpha=0.35, bobot lebih ke data baru), forecast datar di level terakhir.",
    },
}

SMA_WINDOW = 20
SES_ALPHA = 0.35


def _fit_linear(x: np.ndarray, y: np.ndarray, degree: int) -> np.ndarray:
    """Koefisien polinomial derajat `degree` via least-squares (aman untuk
    input pendek -- numpy polyfit bisa gagal/RankWarning kalau overfit)."""
    if len(x) <= degree:
        degree = max(len(x) - 1, 0)
    coef = np.polyfit(x, y, degree)
    return coef


def _forecast_from_coef(coef: np.ndarray, start_idx: int, horizon: int) -> list[float]:
    future_x = np.arange(start_idx, start_idx + horizon, dtype=float)
    return [float(v) for v in np.polyval(coef, future_x)]


def _sma_level(closes: np.ndarray, window: int = SMA_WINDOW) -> float:
    w = min(window, len(closes))
    return float(np.mean(closes[-w:]))


def _ses_level(closes: np.ndarray, alpha: float = SES_ALPHA) -> float:
    level = float(closes[0])
    for v in closes[1:]:
        level = alpha * float(v) + (1 - alpha) * level
    return level


def forecast_all(closes: list[float], horizon: int) -> dict[str, list[float]]:
    """Forecast horizon titik ke depan untuk semua model (tanpa evaluasi)."""
    arr = np.asarray(closes, dtype=float)
    n = len(arr)
    x = np.arange(n, dtype=float)
    out = {
        "linear": _forecast_from_coef(_fit_linear(x, arr, 1), n, horizon),
        "quadratic": _forecast_from_coef(_fit_linear(x, arr, 2), n, horizon),
    }
    sma_lv = _sma_level(arr)
    ses_lv = _ses_level(arr)
    out["sma"] = [sma_lv] * horizon
    out["exp_smooth"] = [ses_lv] * horizon
    return out


def _metrics(actual: np.ndarray, predicted: list[float]) -> dict:
    actual = np.asarray(actual, dtype=float)
    pred = np.asarray(predicted, dtype=float)
    err = pred - actual
    mae = float(np.mean(np.abs(err)))
    rmse = float(np.sqrt(np.mean(err ** 2)))
    denom = np.where(actual == 0, np.nan, actual)
    with np.errstate(divide="ignore", invalid="ignore"):
        mape = float(np.nanmean(np.abs(err / denom)) * 100)
    if np.isnan(mape):
        mape = None
    return {"mae": mae, "rmse": rmse, "mape": mape}


def evaluate_and_forecast(
    closes: list[float], horizon: int, holdout: int | None = None
) -> dict:
    """Latih di histori minus holdout, ukur error di holdout, lalu latih
    ulang di SEMUA data untuk forecast final (praktik standar)."""
    n = len(closes)
    if holdout is None:
        holdout = min(30, max(10, n // 5))
    holdout = max(5, min(holdout, n - 10))
    train = closes[: n - holdout]
    test = closes[n - holdout:]

    # Forecast "pura-pura" dari model yang hanya melihat data train.
    trial = forecast_all(train, holdout)
    metrics = {m: _metrics(np.asarray(test), trial[m]) for m in trial}

    # Forecast final dari model yang melihat semua data.
    final = forecast_all(closes, horizon)
    best_model = min(metrics, key=lambda m: metrics[m]["rmse"])

    return {
        "holdout_days": holdout,
        "metrics": metrics,
        "best_model": best_model,
        "forecasts": final,
    }


def future_business_dates(last: date, horizon: int) -> list[str]:
    """Tanggal hari bursa ke depan (lewati Sabtu/Minggu) sebagai label."""
    out = []
    d = last
    while len(out) < horizon:
        d += timedelta(days=1)
        if d.weekday() < 5:  # Senin-Jumat
            out.append(d.isoformat())
    return out
