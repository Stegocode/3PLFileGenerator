# Architecture Notes

This document explains how the pipeline is organized, the non-obvious design decisions, and where the interesting edge cases live.

## Pipeline overview

The repo is two scripts that run in sequence. They share a `.env` file and a working directory but otherwise have no code dependency — you can run the generator against manually-exported files and skip the scraper entirely.

```
┌─────────────────────────────┐
│   scraper.py                │   headless Selenium, 4 exports
│   (stage 1: acquire)        │   ─ batch invoice    (xlsx)
│                             │   ─ model inventory  (csv)
│                             │   ─ serial inventory (csv)
│                             │   ─ orders detail    (csv)
└──────────────┬──────────────┘
               │   files land in ./inbox
               ▼
┌─────────────────────────────┐
│   generator.py              │   transform + enrich + price
│   (stage 2: transform)      │   ─ ORS geocoding
│                             │   ─ Monday.com crate lookup
│                             │   ─ interactive floor prompts
└──────────────┬──────────────┘
               │
     ┌─────────┴─────────┐
     ▼                   ▼
 flat file         financial report
 (for carrier)     (for dealer)
```

## Generator data flow

```
┌────────────────────┐    ┌────────────────────┐    ┌────────────────────┐
│  Batch invoice     │    │  Model inventory   │    │  Orders detail     │
│  (.xlsx from POS)  │    │  (.csv export)     │    │  (.csv export)     │
└─────────┬──────────┘    └──────────┬─────────┘    └──────────┬─────────┘
          │                          │                         │
          └──────────────┬───────────┴─────────────────────────┘
                         │
                         ▼
               ┌──────────────────────┐         ┌──────────────────────┐
               │   Line classifier    │◀────────│  SERVICE_CODES /     │
               │  (product / part /   │         │  CATEGORY_MAP /      │
               │   install / delete)  │         │  DELETE_CODES /      │
               └──────────┬───────────┘         │  LABOR_LOOKUP / ...  │
                          │                     └──────────────────────┘
                          ▼
               ┌──────────────────────┐
               │   Stop aggregator    │
               │  (group by order#)   │
               └──────────┬───────────┘
                          │
          ┌───────────────┼───────────────┐
          │               │               │
          ▼               ▼               ▼
┌──────────────┐  ┌──────────────┐  ┌──────────────┐
│   ORS API    │  │  Monday API  │  │  Multifamily │
│  geocoding + │  │  crate-status│  │  floor prompt│
│  driving mi  │  │  per order   │  │  (interactive)│
└──────┬───────┘  └──────┬───────┘  └──────┬───────┘
       │                 │                 │
       └─────────────────┼─────────────────┘
                         ▼
               ┌──────────────────────┐
               │   Pricing engine     │
               │  (mileage tier, fuel,│
               │   floor, crate, X001)│
               └──────────┬───────────┘
                          │
                ┌─────────┴──────────┐
                ▼                    ▼
       ┌─────────────────┐  ┌─────────────────┐
       │  Flat file      │  │  Financial      │
       │  (OrderData)    │  │  report (3-tab) │
       └─────────────────┘  └─────────────────┘
```

## Design decisions

### Why config tables at the top of the file, not a YAML

The tool has one operator, and that operator edits carrier rates and service codes a few times a year. A YAML would have saved maybe 50 lines but required the operator to understand a second file. Keeping `SERVICE_CODES`, `CATEGORY_MAP`, and the rate constants as Python literals means a change to a margin or a new X-code is a one-line edit in the same file that's already open.

### Why the Colab shim lives in the same file

The operator's workflow started in Colab before the script moved local. Rather than maintain two forks or introduce a build step, the `_IS_COLAB` detection and the `_colab_files` fallback class let the same file run both places. The trade-off is a small amount of environmental conditional logic in the I/O paths; the benefit is that the operator can always fall back to Colab if their local Python environment breaks.

### Why ZIP-code validation happens before the flat file is written

Orders outside the carrier's service area are common (dealer takes an order, only discovers at scheduling time that HUB won't deliver there). If those orders silently land on the flat file, the carrier rejects the entire upload. Validating up front means the operator sees the bad orders in the console, can resolve them (move to will-call, hand off to a different carrier), and re-run — instead of debugging a rejected upload later.

### Why the financial report is a separate Excel workbook

The flat file has to match the carrier's exact column spec and can't carry extra data. The financial report is for the dealer's margin review — different audience, different columns, different formatting. Two files, one run.

## Edge cases worth a look

- **Dryer installs without gas/electric specified.** The dealer's POS sometimes writes `DRYER INSTALL` without indicating fuel type. The script resolves this at runtime by checking the associated appliance model against the inventory category map.
- **Labor lines that look like products.** Internal labor codes (`WASHERIN`, `REFERINSTALL`, `CONVERSION-DOORSWING`) appear as line items but need to be reclassified as service codes. `LABOR_LOOKUP` handles the main path; `LABOR_PART_CODES` handles third-party installs that stay classified as PART.
- **Monday crate-status label drift.** Users on the scheduling board have typed variations like `OUT OF BOX + INSTA`, `OUT OF BOX + INSTAL`, `OUT OF BOX + INSTALL(S)`. `CRATE_LABEL_MAP` normalizes all of these to a canonical value.
- **Serial-number allocation fallback.** If a serial-number CSV is provided, product cost is resolved per unit. If not, the script falls back to standard cost from the model inventory. The financial report's `cost_source` column records which path was used per line.

## What's not in the repo

- Real input files (invoices, inventories, order details) — these contain customer PII.
- The carrier's Excel order template (`.xltx`) — this is the carrier's IP, not ours to distribute.
- Real API keys — see `.env.example`.
