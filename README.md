# ERP Data Cleaning & Anomaly Detection Pipeline

> **Portfolio project** – Data Analyst Intern application  
> Demonstrates: Docker, Python ETL, IQR/Z-Score/Isolation Forest anomaly detection, Excel reporting

---

## Project Structure

```
erp_pipeline/
├── Dockerfile               # Single-stage Python 3.11-slim image
├── docker-compose.yml       # Pipeline + Jupyter services
├── requirements.txt         # Pinned Python dependencies
├── .dockerignore
├── src/
│   └── pipeline.py          # Full ETL + anomaly detection + Excel export
├── data/
│   └── online_retail.xlsx   # ← PUT INPUT FILE HERE
├── output/                  # ← All outputs land here (auto-created)
└── notebooks/               # Jupyter notebooks (optional EDA)
```

---

## Quick Start

### 1. Add input data
Download **Online Retail dataset** from UCI:  
https://archive.ics.uci.edu/dataset/352/online+retail  
Place `online_retail.xlsx` in the `./data/` folder.

### 2. Build & run
```bash
# Build image + run pipeline
docker compose up --build

# Run in background
docker compose up --build -d

# View live logs
docker compose logs -f pipeline
```

### 3. Outputs
After the run, `./output/` will contain:

| File | Description |
|------|-------------|
| `ERP_Anomaly_Report_<ts>.xlsx` | 8-sheet Excel report with KPIs & anomaly breakdown |
| `erp_cleaned_<ts>.csv` | Cleaned transaction data (anomalies excluded) |
| `anomalies_only_<ts>.csv` | All flagged anomaly records |
| `pipeline.log` | Full debug log of the run |

---

## Configuration (Environment Variables)

Override defaults in `docker-compose.yml` or pass with `-e`:

| Variable | Default | Description |
|----------|---------|-------------|
| `IQR_MULTIPLIER` | `1.5` | Fence width for IQR method |
| `ZSCORE_THRESHOLD` | `3.0` | Sigma cutoff for Z-score |
| `ISO_CONTAMINATION` | `0.03` | Expected anomaly rate for Isolation Forest |
| `ANOMALY_EXPORT_LIMIT` | `10000` | Max rows in Anomaly_Records sheet |
| `LOG_LEVEL` | `INFO` | Loguru level: DEBUG / INFO / WARNING |

Example — stricter detection:
```bash
docker run --rm \
  -v $(pwd)/data:/app/data \
  -v $(pwd)/output:/app/output \
  -e IQR_MULTIPLIER=1.0 \
  -e ZSCORE_THRESHOLD=2.5 \
  -e ISO_CONTAMINATION=0.05 \
  erp-pipeline:2.0
```

---

## Pipeline Steps

```
online_retail.xlsx
       │
       ▼
 [1] LOAD          pandas read_excel → 541,909 rows
       │
       ▼
 [2] CLEAN
       ├─ Drop exact duplicates          (-5,268)
       ├─ Tag missing CustomerID → GUEST (-135,037 tagged, kept)
       ├─ Remove cancellations (C-prefix) (-9,251)
       ├─ Correct negative quantity (abs) (1,336 fixed)
       ├─ Flag UnitPrice < 0.01           (2,516 flagged)
       └─ Add: Revenue, Quarter, Month, DayOfWeek, Hour
       │
       ▼
 [3] FEATURE ENGINEERING
       ├─ OrderSize  (total qty per invoice)
       └─ RFM proxy  (Recency, Frequency, Monetary per customer)
       │
       ▼
 [4] ANOMALY DETECTION
       ├─ IQR  × 3 columns  (Quantity, UnitPrice, Revenue)
       ├─ Z-Score × 3 columns
       ├─ Isolation Forest  (contamination = 0.03)
       └─ Composite Is_Anomaly flag + Anomaly_Score (0–7)
       │
       ▼
 [5] EXPORT
       ├─ erp_cleaned_<ts>.csv
       ├─ anomalies_only_<ts>.csv
       └─ ERP_Anomaly_Report_<ts>.xlsx  (8 sheets)
```

---

## Excel Report Sheets

| Sheet | Contents |
|-------|----------|
| `Executive_Summary` | KPI tiles, cleaning log, anomaly breakdown table |
| `Monthly_KPI` | Revenue, invoices, customers by month + color scale |
| `Country_KPI` | Revenue & transactions by country |
| `Top_50_Products` | Top products by revenue |
| `Anomaly_By_Month` | Anomaly count & rate per month (red gradient) |
| `Anomaly_Records` | Flagged rows color-coded by Anomaly_Score (0–7) |
| `Cleaned_Data_Sample` | 2,000 sample rows of clean data |

---

## Optional: Jupyter EDA

```bash
# Launch notebook server (http://localhost:8888)
docker compose --profile notebook up jupyter
```

---

## CV Bullet Point

> *"Xây dựng Dockerized ERP Data Cleaning & Anomaly Detection Pipeline xử lý 541K+ transaction records; áp dụng IQR, Z-Score và Isolation Forest để phát hiện ~15.6% anomaly records; tự động export báo cáo Excel 8-sheet với KPI dashboard và conditional formatting — toàn bộ chạy qua Docker Compose với cấu hình môi trường linh hoạt."*
