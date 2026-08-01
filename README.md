# Nepal Power Plant & Transmission Line License Status Dashboard

A fully self-contained web dashboard for visualizing Nepal's power plant and
transmission line license data — **plus** an independent NEA Operational
Performance module (System Loss, Energy Balance, Financials, etc.) fed from
its own Google Sheet.

## Key Features

- **Zero-config GIS**: Nepal district/province boundaries and protected areas
  are bundled inline — no shapefile uploads or Drive syncs needed
- **Admin panel** at `/admin` for uploading workbooks, syncing Google Sheets,
  and managing settings
- **9 license-status dashboard tabs**: Overview, Power Plants, Transmission
  Lines, GoN Studied, Cancelled, Growth, GIS Map, Compare, Data Table
- **NEA Operational Performance + Forecast Lab**: a second, fully independent
  tab pair (own Google Sheet, own background sync) with statsmodels-driven
  forecasting (Linear / Holt / Moving Average / ARIMA / SARIMA / Hybrid) and
  stacked multi-component scenarios
- **Live ticker** with KPIs and latest connections (both license-status and
  NEA figures)
- **PDF report generation** (10-page report covering both data sources)
- Visitor counter and live clock

## Quick Start

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run locally
python app.py

# 3. Open browser to http://localhost:8050
```

## Environment Variables (optional)

| Variable | Description |
|---|---|
| `DEFAULT_SHEET_URL` | Google Sheet URL/ID for the **License Status** workbook, auto-synced on startup |
| `DEFAULT_GIS_DRIVE_URL` | Google Drive link for high-res GIS zip (optional) |
| `DEFAULT_PA_DRIVE_URL` | Google Drive link for protected areas zip (optional) |
| `NEA_SHEET_URL` | Google Sheet URL/ID for the **NEA Operational Data** workbook (separate from the license workbook above) — required for the "System Operational Performance" and "NEA Forecast Lab" tabs to show any data |
| `NEA_AUTO_REFRESH_HOURS` | Background refresh interval for the NEA sheet, in hours (default: `6`) |
| `ADMIN_PASSWORD` | Password for `/admin` panel (default: `admin123`) |
| `FLASK_SECRET_KEY` | Session encryption key |
| `DATA_DIR` | Path to persistent data storage (also where `nea_config.json` — the admin-saved NEA sheet URL — is written, so it survives a restart) |
| `AUTO_REFRESH_HOURS` | Background refresh interval for the License Status sheet (default: `6`) |

> **Note:** the License Status workbook and the NEA Operational Data workbook
> are two entirely separate Google Sheets, synced independently, on their own
> schedules, by their own code paths. Setting `DEFAULT_SHEET_URL` alone will
> **not** populate the NEA tabs — you need `NEA_SHEET_URL` (or the admin
> panel's "NEA Data Sync" card) for that.

## File Structure

```
.
├── app.py                  # Main Dash application
├── NEA.py                  # NEA Operational Data + Forecast Lab (independent module)
├── data_engine.py          # Data loading, GIS engine, B.S. calendar helpers
├── server_state.py         # Shared state, config persistence, background refresh
├── coordinate_transform.py # WGS-84 / Everest 1830 conversions
├── admin.py                # Admin panel blueprint
├── gis_bundled.py          # Inline Nepal GIS data (districts + protected areas)
├── nea_assets/             # NEA forecast/dashboard HTML templates + vendor JS
├── nepal_flag.png          # Default flag image
├── requirements.txt        # Python dependencies
└── data/                   # Created at runtime (workbooks, uploads, config)
```

## Adding Project Data (License Status Workbook)

### Option 1: Upload via Admin Panel
1. Go to `/admin/login` (default password: `admin123`)
2. Upload an `.xlsx` workbook with columns like:
   - `project_name`, `type`, `status`, `capacity_mw`, `voltage_kv`, `line_length_km`
   - `district`, `province`, `promoter`, `latitude`, `longitude`
   - `license_date`, `cod` (commercial operation date in B.S.)

### Option 2: Google Sheet Sync
1. Share your Google Sheet as "Anyone with the link"
2. Paste the URL in the admin panel or set `DEFAULT_SHEET_URL`

### Option 3: Local File
Place `workbook.xlsx` in the `data/` folder before starting.

## Adding NEA Operational Data (separate workbook)

This feeds the **System Operational Performance** tab and the **NEA Forecast
Lab** tab. Set `NEA_SHEET_URL` (env var), or paste the sheet URL/ID into the
"NEA Data Sync" card on `/admin` (this persists it to `nea_config.json`, so it
survives a restart without needing the env var set on the platform).

The sheet must be shared as **"Anyone with the link can view"**, and must
contain exactly these 11 tabs. Tab names are matched **case- and
whitespace-insensitively**, so `"System Loss "` or `"system loss"` both match
— but the wording itself must match one of the aliases below:

| Internal key | Required tab name (any case/spacing) | Shape |
|---|---|---|
| `system_loss` | **System Loss** | One row per fiscal year: `Fiscal Year \| Transmission Loss (%) \| Distribution Loss (%) \| System Loss (%)` |
| `financial_data` | **Financial Data** | One row per fiscal year: `Fiscal Year \| Overall Revenue \| Profit/(Loss) \| Import (MU) \| Import (Rs.) \| Export (MU) \| Export (Rs.)` — column headers just need to contain the right keywords (e.g. "Import" + "MU", "Import" + "Rs."/"Million") |
| `annual_energy_peak` | **Annual Energy and Peak Load** | One row per category, one column per fiscal year (`Particulars \| FY1 \| FY2 \| ...`). Row labels must contain: "NEA" + "Own", "NEA" + "Sub", "IPP", "India", "Total" + "Availab", "National" + "Peak", "System" + "Peak" |
| `consumers_growth` | **Consumers Growth** | Column-oriented FY, one row per consumer category, plus a "Total Consumers" row and a row containing "growth" |
| `sales_revenue` | **Sales Revenue** | Column-oriented FY, one row per consumer category, plus a "Total Gross Revenue" (or "Total") row and a row containing "growth" |
| `transmission_line` | **Transmission Line Length** | `Fiscal Year (B.S.) \| 66kV \| 132kV \| 220kV \| 400kV \| Total \| Increment` |
| `substation_capacity` | **Substation Capacity** | `Fiscal Year (B.S.) \| Capacity \| Increment` |
| `energy_export` | **Energy Export in GWh** | `Fiscal Year \| 12 month columns` |
| `energy_import` | **Energy Import in GWh from India** | `Fiscal Year \| 12 month columns` |
| `energy_balance` | **Energy Balance in GWh** | Wide grid: row 1 = FY spanning each 12-month block, row 2 = month names, then one row per generation-source category. Row labels must contain: "IPP", "NEA"+"Subsidiary", "ROR", "Import", "NEA"+"Storage", "NEA"+"Solar", "Thermal", "Interruption", "Monthly"+"System"+"Energy"+"Demand", "Export", "Monthly"+"National"+"Energy"+"Demand" |
| `capacity_balance` | **Capacity Balance in MW** | Same wide-grid shape as Energy Balance. Row labels must contain: "IPP", "NEA"+"Subsidiary", "ROR", "Import", "NEA"+"Storage", "Interruption", "Monthly"+"National"+"Peak", "Export", "Monthly"+"System"+"Peak" |

If any of these 11 tabs is missing, `NEA.parse_workbook()` raises a
`ValueError` naming exactly which one(s) are absent — check server logs (or
`/admin`'s sync status) if the NEA tabs come up empty.

### Forecast Lab parameter list

The Forecast Lab's dropdowns (single-series and stacked/composite) are built
entirely from whatever the above 11 sheets currently contain — no fiscal year
or category is hardcoded. Adding a new FY column, or a new consumer/revenue
category row, shows up automatically after the next sync; nothing needs to be
redeployed.

## GIS Data

The bundled GIS data in `gis_bundled.py` contains simplified boundaries for
all 77 Nepal districts and major protected areas. This is sufficient for
choropleth shading and spatial queries.

To use higher-resolution boundaries:
1. Obtain shapefiles from Nepal Survey Department
2. Simplify with mapshaper: `mapshaper -i input.shp -simplify 10% -o output.shp`
3. Upload via admin panel or set `DEFAULT_GIS_DRIVE_URL`

## Deployment

### Render (Recommended)
1. Push to GitHub
2. Create new Web Service on Render
3. Set build command: `pip install -r requirements.txt`
4. Set start command: `gunicorn app:server`
5. Add environment variables as needed (see table above — remember
   `NEA_SHEET_URL` is separate from `DEFAULT_SHEET_URL`)

### Docker

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install -r requirements.txt
COPY . .
CMD ["gunicorn", "app:server", "--bind", "0.0.0.0:8050"]
```

## License

© 2026 Er. Sandeep Neupane. All rights reserved.
