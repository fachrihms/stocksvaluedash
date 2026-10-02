FROM python:3.12-slim

WORKDIR /code

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY app/ ./app/
COPY static/ ./static/
COPY tests/ ./tests/
COPY README.md .

# Database SQLite dibuat otomatis saat runtime (jangan bake file lokal)
CMD uvicorn app.main:app --host 0.0.0.0 --port ${PORT:-8000}
