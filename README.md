# Order Book Data Processing System

A responsive web application for uploading, processing, reviewing, filtering, and downloading Excel order-book data. The latest successful upload is stored centrally so everyone using the shared link sees the same data until another valid file is uploaded.

## Links

- **Live application:** https://order-book-system.onrender.com/
- **GitHub repository:** https://github.com/84548123/order-book-system
- **Health check:** https://order-book-system.onrender.com/health
- **API documentation:** https://order-book-system.onrender.com/docs

## Main features

- Uploads `.xlsx` workbooks up to the configured size and row limits.
- Validates and processes the configured columns from the first worksheet.
- Uses `ItemPoNo` when `PoNo` is missing.
- Treats matching text consistently regardless of uppercase or lowercase.
- Calculates Gap Days, showing remaining days as positive and overdue days as negative.
- Provides Excel-style searchable, multi-select checkbox filters on every column.
- Provides numeric conditions for **Processed Data → Gap Days**, including greater than or equal, less than or equal, and between.
- Sorts order dates chronologically and keeps table headings and totals visible while scrolling.
- Supports responsive desktop and mobile viewing.
- Creates and downloads a formatted Excel workbook with filters and frozen headings.
- Keeps the latest processed data available to every visitor through persistent Supabase storage.

## Workbook views

### 1. Processed Data

Detailed order records with these columns:

`OrderDate`, `ClientCode`, `StyleCode`, `BagQty`, `PoNo`, `ExpDiaDlvDate`, `Order Book`, `Expected Dia Qly`, `ExpectedDiaWt`, `Total Value`, `Gap Days`, `Notes`, and `PPC`.

`PPC` is populated from the source `Manufacturer` value. Totals are shown for Bag Quantity, Expected Diamond Weight, and Total Value.

### 2. Detail summary

PO-level summary containing:

`Date`, `Order Book`, `Client name`, `Bag Qty`, `PO No.`, and `Gap Days`.

Rows sharing the same PO number are grouped and their Bag Quantity is totaled.

### 3. Factory Summary

Factory-level summary containing:

`Factory Name`, `Total No. of Bag Qty`, `Total Sum of Expected Dia Wt`, and `Total Sum Value`.

Summary values are displayed without decimals.

## Architecture

```text
Desktop / Mobile Browser
          |
          v
Render-hosted FastAPI application
          |
          +--> Excel validation and processing (openpyxl)
          |
          +--> Supabase PostgreSQL persistent latest-data storage
          |
          +--> Browser dashboard and generated Excel download

GitHub repository --> Automatic Render deployment
```

## Technology stack

| Area | Technology | Purpose |
|---|---|---|
| Frontend | HTML, CSS, JavaScript | Responsive dashboard, tabs, search, sorting, and filters |
| Backend | Python, FastAPI | Upload handling, processing, APIs, and file downloads |
| Excel | openpyxl | Workbook reading, calculations, formatting, and generation |
| Server | Uvicorn | Runs the FastAPI application |
| Persistent storage | Supabase PostgreSQL | Stores the latest dataset and generated workbook |
| Hosting | Render | Hosts the publicly accessible web service |
| Source control | Git and GitHub | Version history and automatic deployment source |

## Data flow

1. A user uploads an Excel workbook.
2. The backend validates and processes the workbook.
3. The application creates Processed Data, Detail summary, and Factory Summary.
4. The latest successful result and generated workbook are saved in Supabase.
5. Everyone opening the shared link sees the same saved result.
6. A later successful upload replaces the previously displayed dataset.

## Run locally

Python 3.11 or later is required.

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open http://localhost:8000/.

## Run with Docker

```bash
docker compose up --build
```

## API endpoints

- `GET /health` — service health check
- `GET /api/latest` — latest shared processed data
- `POST /api/process` — process a multipart `.xlsx` upload using the `file` field
- `GET /api/download/{file_id}` — download the generated workbook

## Configuration

Configure these environment variables in Render or a local `.env` file:

- `MAX_UPLOAD_MB` — maximum upload size; default `15`
- `MAX_ROWS` — maximum accepted worksheet rows; default `100000`
- `FILE_TTL_MINUTES` — temporary file lifetime; default `30`
- `SUPABASE_URL` — Supabase project URL
- `SUPABASE_SERVICE_ROLE_KEY` — private server-side Supabase service key

Never commit the Supabase service-role key to GitHub.

## Deployment notes

Render automatically deploys changes pushed to the connected GitHub branch. On Render's free web-service tier, the application can sleep after inactivity, so the first request may take about a minute. The data remains stored in Supabase and appears again after the service wakes.

## Data note

Identifiers stored as Excel text retain leading zeros. Numeric cells with an all-zero number format can also be reconstructed with leading zeros. Zeros already discarded by Excel from a general-format numeric value cannot be recovered.
