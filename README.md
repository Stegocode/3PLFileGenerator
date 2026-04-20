# 3PLFileGenerator

A Python pipeline that transforms a dealer's daily batch invoice into a carrier-ready delivery flat file, with automatic mileage pricing, crate-status lookup from a Monday.com board, multifamily floor-charge prompts, and a per-stop profit & loss report.

Built to replace a manual, error-prone spreadsheet workflow that took an ops coordinator 60–90 minutes every morning. The full pipeline runs in about a minute and produces two files: the flat file ready for upload to the 3PL carrier portal, and a three-sheet financial report breaking down revenue, cost, and margin by stop, by service code, and by product.

The repo is two scripts meant to run in sequence:

1. **`scraper.py`** — headless Selenium automation that logs into the Homesource dealer management system and exports the four input files (batch invoice, model inventory, serial inventory, orders detail) filtered to the next business day and the target carrier route.
2. **`generator.py`** — reads the exported files, applies the pricing and classification logic, and writes the carrier flat file plus the financial report.

Each script is usable on its own — if you already have the input files from Homesource (manual exports, a different DMS, or a one-off), you can skip the scraper and run the generator directly.

---

## What it does

Given three input files (a batch invoice, a model inventory CSV, and an orders-detail CSV, with an optional serial-number allocation CSV for unit-level costing), the flat file generator:

1. **Classifies every invoice line** — products, installs, parts, haul-aways, redeliveries, accessories, and accounting-only codes that shouldn't ship to the carrier at all.
2. **Geocodes each stop** against a warehouse origin using Open Route Service, and applies a three-tier mileage pricing model (free under 30 mi, flat charge 30–125 mi, flagged for review above 125 mi).
3. **Queries Monday.com** via GraphQL to pull the crate-status column for each scheduled order, then maps it to the correct carrier service code (in-box vs. out-of-box piece charges).
4. **Prompts interactively** for floor-charge data on multifamily orders, where the invoice doesn't capture unit and floor.
5. **Validates ZIP codes** against the carrier's service area before writing the flat file — orders outside the zone are flagged, not silently dropped.
6. **Writes two Excel outputs**: a HUB-format flat file for upload, and a styled three-sheet financial report (By Stop / Service Detail / Product Detail) with color-coded P&L columns.

If a serial-number allocation CSV is provided, product cost is resolved at the unit level, so margin on the financial report reflects the actual cost basis of the specific unit delivered rather than a standard cost average.

---

## Why it exists

This is a real operational tool solving a real bottleneck. It's also a deliberately reusable one: all carrier rates, service codes, and classification logic live in config tables at the top of the script, so it can be adapted to another dealer or another carrier by editing constants rather than rewriting logic.

The code is organized around three themes worth a second look for reviewers:

- **Data transformation under messy real-world constraints.** The invoice data has inconsistent model-number formatting, legacy codes that need to be filtered, labor lines that need to be reclassified as installs, and accessory parts that look like appliances to the dealer's POS but need to be routed differently on the carrier flat file.
- **External API orchestration with graceful degradation.** Geocoding and routing calls go through Open Route Service; crate status comes from Monday.com. Both paths have timeouts, fallback behavior, and readable error surfacing so that a missing API response fails loudly to the operator instead of producing a quietly-wrong flat file.
- **Environment portability.** The same file runs in Google Colab (reading credentials from Colab Secrets and using the Colab file-picker) and on a local Windows box (reading from `.env` and auto-detecting files in a working directory). The Colab shim is six lines at the top of the file.

---

## Install

```bash
git clone https://github.com/Stegocode/3PLFileGenerator.git
cd 3PLFileGenerator
pip install -r requirements.txt
cp .env.example .env
# edit .env with your ORS_API_KEY and MONDAY_API_TOKEN
```

You'll also need the carrier's Excel order template (`.xltx`) in your working directory. The template filename is set in the `CONFIG` block at the top of the script.

---

## Run

Two-stage pipeline. Run the scraper first, then the generator:

```bash
python scraper.py       # pulls 4 CSVs/XLSX from Homesource into ./inbox
python generator.py  # reads ./inbox, writes ./exports
```

The generator auto-detects input files in the working directory by filename pattern (batch invoice `.xlsx`, model inventory `.csv`, orders detail `.csv`, optional serial allocation `.csv`). You'll be prompted for floor and unit on any multifamily orders, then two files will be written to the export directory.

For Colab, paste either script into a notebook cell, set the relevant credentials as Colab Secrets (`HS_USERNAME`/`HS_PASSWORD` for the scraper, `ORS_API_KEY`/`MONDAY_API_TOKEN` for the generator), and run. Both scripts detect the Colab environment automatically and switch to the interactive file-upload / file-picker flow.

---

## Configuration

Both scripts use a `CONFIG` dict at the top of the file for instance-specific values.

**`scraper.py`:**

| Setting | What it controls |
| --- | --- |
| `base_url` | Your Homesource subdomain, e.g. `https://acme1.homesourcesystems.com` |
| `truck_filter` | Truck/route label to filter on the batch invoice page |
| `use_next_business_day` | If True, always targets the next business day (Fri→Mon) |
| `wait_timeout` | Selenium wait timeout in seconds |

**`generator.py`:**

All dealer-specific and carrier-specific values live in the `CONFIG` dict and the rate-constant block near the top of the script:

| Setting | What it controls |
| --- | --- |
| `location_code` | Carrier's account identifier, stamped on every flat file row |
| `warehouse_address` | Origin for all mileage calculations |
| `output_prefix` | Prefix for output filenames, e.g. `HUB_04.18.26.xlsx` |
| `delivery_scheduler_board_id` | Monday.com board to query for crate status |
| `monday_crate_col_id` / `monday_date_col_id` | Column IDs on that board |
| `MIN_CHARGE`, `FLOOR_CHARGE`, `MILEAGE_CHARGE`, `FUEL_SHORT`, `FUEL_LONG` | Rate constants, edit as carrier rates change |
| `SERVICE_CODES` | Full table of X-codes with description, charge, and margin |
| `HUB_ZIPS` | Carrier service area — orders outside this set are flagged |

To adapt this tool to a different dealer or carrier:
1. Update `CONFIG` and the rate constants.
2. Edit `SERVICE_CODES` to match the new carrier's code sheet.
3. Edit `HUB_ZIPS` for the new service area.
4. Review `CATEGORY_MAP`, `PARTS_LOOKUP`, `LABOR_LOOKUP`, and `DELETE_CODES` for dealer-specific POS quirks.

---

## Project structure

```
3PLFileGenerator/
├── scraper.py         # stage 1: pull inputs from Homesource
├── generator.py    # stage 2: transform + write flat file
├── local_config.py               # loads .env and sets up inbox/exports dirs
├── requirements.txt
├── .env.example
├── .gitignore
├── LICENSE
├── README.md
└── docs/
    └── architecture.md           # notes on data flow and design decisions
```

---

## Tech

Python 3.10+, pandas, openpyxl, requests, selenium, python-dotenv. No framework dependencies, no database — the scraper drives a headless Chrome, and the generator reads CSV/XLSX, makes a handful of HTTP calls, and writes XLSX.

---

## License

MIT — see [LICENSE](LICENSE).
