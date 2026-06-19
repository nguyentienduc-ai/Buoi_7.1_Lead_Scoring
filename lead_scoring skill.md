---
name: lead-scoring-automation
description: "Downloads real estate customer lead data from Google Sheets, applies AI business rules to score potential, runs a local dashboard for human-in-the-loop verification, and exports results to a styled Excel sheet."
argument-hint: "[optional Google Sheets URL]"
license: MIT
version: 1.0.0
tags: [lead-scoring, real-estate, automation, dashboard, excel]
---
# Real Estate Lead Scoring & Automation Skill

This skill automates the collection, AI scoring, human verification, and exporting of customer leads for real estate.

## Step 1 — Sync Customer Data
Retrieve leads from the Google Sheets document. If a spreadsheet URL is provided, the script will use it; otherwise, it defaults to the pre-configured spreadsheet.
The script downloads the sheet, parses the rows, and registers them for scoring.

## Step 2 — Run the AI Lead Scorer
Evaluate each customer lead description against the target business rules. A base score of **100** is set. The script applies:
- **+50 Points (VIP / Super Potential)** for high budgets (>=20 billion VND or "tài chính mạnh"), premium types ("Biệt thự", "Penthouse", "Shophouse"), prime locations ("Quận 1", "Ven sông", "Phú Mỹ Hưng"), high-value buyers ("Chủ doanh nghiệp", "Nhà đầu tư"), and urgent/clear legal status.
- **-50 Points (Junk / Low Potential)** for unrealistic requests (e.g. District 1 house for 1-2 billion VND), "wrong number", "no demand", spam/advertising, or communication failures.
- **0 Points adjustment (Neutral)** for normal ranges (3-10 billion VND apartments) or general consultations.

Run the script manually using:
```powershell
python scripts/score_leads.py
```
This updates/creates the database file `scored_leads.json`.

## Step 3 — Launch the Human-in-the-Loop Web App
Human verification is required before finalizing scores. Start the local server:
```powershell
python server.py
```
This serves a premium dashboard on `http://localhost:9090`. In the dashboard:
1. Review the KPI metrics (Total leads, VIP leads, Junk leads).
2. Filter or search through leads.
3. Review AI score details and reasons.
4. Manually override scores, modify comments, or change the status if needed.
5. Click **Save** to update the local database.

## Step 4 — Export Scored Data
Once all leads are audited, click the **Export Excel** button on the dashboard or access `http://localhost:9090/api/export` to generate `scored_leads_final.xlsx`.

## Gotchas & Design Guidelines

- **File Encoding**: Standard Windows terminal utilizes CP1252. Ensure all scripts read and write files using `utf-8` encoding explicitly to avoid crash errors with Vietnamese accents.
- **Spreadsheet Quality**: The exported Excel report must be highly professional:
  - Header: Dark Navy Blue (`#1F4E78`) background, white bold text.
  - Fonts: Set uniform font family (e.g., `Segoe UI`).
  - Formats: Phone numbers must be displayed cleanly, scores centered, and descriptions wrapped properly.
  - Gridlines: Explicitly enable gridlines (`ws.views.sheetView[0].showGridLines = True`).
  - Auto-fit: Automatically adjust column widths so no text is truncated.
- **Base Score Rules**: Do not clear existing human overrides when syncing new Google Sheets data. Only update new leads or run scoring for untouched entries.
