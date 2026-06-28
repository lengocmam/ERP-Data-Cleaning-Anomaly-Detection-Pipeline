FROM python:3.11-slim

LABEL maintainer="portfolio-project"
LABEL description="ERP Data Cleaning & Anomaly Detection Pipeline"
LABEL version="2.0"

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

WORKDIR /app

RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY src/ ./src/

RUN mkdir -p /app/data /app/output

ENV DATA_DIR=/app/data \
    OUTPUT_DIR=/app/output \
    IQR_MULTIPLIER=1.5 \
    ZSCORE_THRESHOLD=3.0 \
    ISO_CONTAMINATION=0.03 \
    ANOMALY_EXPORT_LIMIT=10000 \
    LOG_LEVEL=INFO

ENTRYPOINT ["python", "-u", "src/pipeline.py"]
