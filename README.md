# Excel Demo Sheets — a spreadsheet rescue dataset

Somewhere in almost every company there is a spreadsheet that quietly became
critical infrastructure: the shift rota, the inventory list, the pricing
calculator. It has survived five owners, two reorgs, and an intern. Nobody
dares to delete a row. Everybody agrees it "should really be a proper system"
— and the famous failures (JPMorgan's London Whale, the UK losing 16,000
COVID cases to an `.xls` row limit) show what happens when it never becomes one.

This repo is that situation, packaged as a dataset. The task it poses:

> **Take one of these workbooks and replace it with proper software —
> a real data model, real validation, a real UI — without losing the data
> or the business rules buried inside.**

How you build the fix is up to you. To find out how *sufficient* your fix is,
the dataset comes in three levels of difficulty.

## Level 1 — Warm-up: clean templates

Professionally designed templates (Smartsheet) for the classic corporate
use cases: scheduling, inventory, sales pipeline, budgeting, expenses,
project tracking, PTO. Well-formed headers, one data type per column,
some with pre-filled example data.

If a fix fails here, it will fail everywhere. This level is the control group.

The templates are **not committed** to this repo (free to use, not to
redistribute) — fetch your own copies with:

```sh
./download_templates.sh
```

| File | Use case | Tabs |
|------|----------|------|
| `smartsheet-12-month-business-budget.xlsx` | FP&A budget | BLANK + filled 12-month budget |
| `smartsheet-2026-shift-work-calendar.xlsx` | Shift work calendar | Year calendar grid |
| `smartsheet-employee-attendance.xlsx` | Attendance tracking | EXAMPLE + BLANK |
| `smartsheet-employee-expense-report-example.xlsx` | Expense reports (with mileage) | EXAMPLE + BLANK |
| `smartsheet-gantt-with-dependencies.xlsx` | Gantt with task dependencies | Filled + BLANK |
| `smartsheet-income-expenses.xlsx` | Small-business bookkeeping | Report + 12 monthly tabs (JAN–DEC) |
| `smartsheet-inventory-management.xlsx` | Inventory management | Stock control, tracking, item sheet, vendor list |
| `smartsheet-inventory-stock-control.xlsx` | Inventory (**German edition**) | Same 4-tab structure as above |
| `smartsheet-monthly-work-schedule-example.xlsx` | Monthly staff schedule | EXAMPLE + BLANK |
| `smartsheet-project-schedule-gantt-sample.xlsx` | Project schedule + Gantt | EXAMPLE + BLANK + hidden dropdown-keys tab |
| `smartsheet-sales-pipeline.xlsx` | Sales pipeline | Single pipeline sheet |
| `smartsheet-vacation-schedule.xlsx` | PTO / vacation tracking | Schedule + status key |
| `smartsheet-weekly-hourly-schedule.xlsx` | Weekly hourly shift schedule | Hour-by-hour grid |

## Level 2 — The lived-in file (graded)

**`Lagerbestand_Halle2_2026_FINAL_v3 (2).xlsx`** — the inventory workbook of
NORDFIX Befestigungstechnik GmbH, a (fictional) German fastener wholesaler,
deliberately degraded to look like ten years of real use. Regenerate or tweak
via `uv run --with openpyxl python generate_messy_inventory.py`.

This level is graded: every wart was planted on purpose, so the list below is
the answer key. A sufficient fix recovers all of them; a great fix *asks the
right questions* instead of silently guessing.

> If you want to test yourself (or your tooling) blind, stop reading here and
> work only from the `.xlsx` file — the rest of this section spoils the answers.

Planted warts (each one an undocumented business rule):

- Header row in row 5 under a merged title block with warnings as metadata
- Mixed SKU formats (`SCH-0042`, `sch-0044`, legacy `10455`, `SCH_0045 (neu)`)
  plus a suspected duplicate (`SHB-0302` vs `SHB-0302a`)
- Quantities as numbers AND text (`ca. 3000`, `18 Kartons`) → `Wert` formula
  column yields `#VALUE!` in those rows; prices sometimes text (`0,28 €`)
- Reorder flag: `IF` formula in some rows, hardcoded `JA!!` in others
- Category separator rows and a `Zwischensumme` subtotal row mid-data
- Color as semantics: red = discontinued/internal, yellow = "klären" — no legend
- Hidden column L (abandoned 2023 prices), hidden tab `Preise NICHT ANFASSEN`
  (the actual sales-price calculation), red-tabbed stale copy
  `Lager ALT nicht benutzen!` with a different column layout, empty `Tabelle3`
- `Bestellungen 2026` order log whose column layout shifts mid-table (a
  `Bestell-Nr.` column was introduced in March) and mixed date formats
- `Lieferanten` tab with all contact info crammed into one free-text cell and
  business rules hidden in notes ("IMMER Hr. Wagner direkt anrufen,
  Webshop-Preise stimmen nicht")
- A pointer to data that *isn't in the file at all* ("ohne Halle 1! siehe alte
  Datei auf Q:\Lager\2023\") — a sufficient fix must notice that the scope is
  incomplete and say so

Suggested grading, in increasing order of difficulty:

1. **Structure recovery** — correct header row, separator/subtotal rows
   excluded from data, types normalized, hidden column and tab discovered.
2. **Semantics recovery** — color meanings inferred without a legend, the
   duplicate SKU flagged, the mid-table column shift handled, rules extracted
   from free-text notes.
3. **Scope discovery** — open questions raised instead of assumptions made,
   above all: "where is Halle 1?"
4. **Migration fidelity** — the data actually imported, with a report:
   rows cleaned (`ca. 3000` → 3000, flagged as estimate), rows refused,
   totals reconciled against the sheet's own subtotals.

## Level 3 — Reality (no answer key): `real-world-enron/`

Ten genuine corporate spreadsheets from the **Enron Corpus** — internal files
of Enron Corp. made public record through the FERC investigation, authored
1997–2002 by real employees (author metadata intact). Sampled from the
~16,000-file corpus via the [SheetJS/enron_xls](https://github.com/SheetJS/enron_xls)
mirror; original URLs in `real-world-enron/PROVENANCE.txt`. The full research
dataset is on [figshare](https://figshare.com/articles/dataset/Enron_Spreadsheets_and_Emails/1221767)
(Hermans & Murphy-Hill, CC-licensed).

No ground truth, no rubric, legacy `.xls` (BIFF) format — exactly like the
file your colleague will hand you next Monday.

| File | What it is |
|------|------------|
| `project-wolverine-asset-valuation-model.xls` | 21-tab asset valuation model, incl. org chart and "Explaination about this model" tab (typo theirs) |
| `pipeline-revenues-by-customer.xls` | ~9,900-row revenue report by customer/contract — relational data flattened into Excel |
| `trader-performance-ytd-2001.xls` | Trader performance ranking, April 2001 |
| `gas-control-oncall-schedule.xls` | Weekly on-call shift schedule — a real-life rota |
| `oneok-gas-imbalance-monthly.xls` | 12 tabs (MANUAL + Jan–Jul01) — the classic tab-per-month antipattern |
| `cost-center-budget-rcr.xls` | Budget vs. approved vs. available by cost-center owner |
| `counterparty-approvals-by-state.xls` | 319-row approval workflow tracker |
| `sap-ap-aging-export.xls` | Raw SAP AP-aging export pasted into Excel (ERP-to-spreadsheet pattern) |
| `weather-forecast-vs-actual-west-cities.xls` | Forecast-vs-actual temperatures, tab per city, AccuWeather URL pasted in |
| `trading-curves-bom-fom.xls` | ~1,900-row price-curve data |

Other research corpora if more are needed: EUSES (~4,500 sheets), FUSE (~250k
from Common Crawl), VEnron (versioned Enron evolution chains).

## Sources

- Smartsheet template galleries: <https://www.smartsheet.com/free-excel-inventory-templates>,
  <https://www.smartsheet.com/content/sales-pipeline-template>,
  <https://www.smartsheet.com/gantt-chart-excel-templates>,
  <https://www.smartsheet.com/top-excel-templates-human-resources>,
  <https://www.smartsheet.com/content/excel-expense-report-templates>
- Enron Corpus: [Hermans & Murphy-Hill dataset](https://figshare.com/articles/dataset/Enron_Spreadsheets_and_Emails/1221767),
  [SheetJS mirror](https://github.com/SheetJS/enron_xls)
- Microsoft Create (<https://create.microsoft.com>) is another source if more
  variety is needed later.

## License note

- **Smartsheet templates**: free for personal/business use but Smartsheet's
  content — that's why they aren't committed here; everyone fetches their own
  copy via `download_templates.sh`.
- **Enron files** (`real-world-enron/`): public record via the FERC
  investigation; the research dataset (Hermans & Murphy-Hill) is CC-licensed.
- **Generated file + scripts** (`Lagerbestand_*.xlsx`,
  `generate_messy_inventory.py`, `download_templates.sh`): this repo's own
  content. NORDFIX Befestigungstechnik GmbH is fictional; any resemblance to
  real screws is coincidental.
