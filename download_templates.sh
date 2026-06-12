#!/usr/bin/env bash
# Fetches the free Smartsheet templates referenced in README.md.
# They are not committed to this repo because Smartsheet's terms allow free
# use but not redistribution — so everyone downloads their own copy.
set -euo pipefail
cd "$(dirname "$0")"

declare -a files=(
  "smartsheet-sales-pipeline.xlsx|https://www.smartsheet.com/sites/default/files/sales-pipeline-template.xlsx"
  "smartsheet-2026-shift-work-calendar.xlsx|https://www.smartsheet.com/sites/default/files/2025-09/IC-2026_Shift-Work-Calendar.xlsx"
  "smartsheet-12-month-business-budget.xlsx|https://www.smartsheet.com/sites/default/files/IC-12-Month-Business-Budget-Template-8821-Updated.xlsx"
  "smartsheet-income-expenses.xlsx|https://www.smartsheet.com/sites/default/files/2022-08/IC-Small-Business-Spreadsheet-for-Income-and-Expenses-11313.xlsx"
  "smartsheet-weekly-hourly-schedule.xlsx|https://www.smartsheet.com/sites/default/files/2025-11/IC-Weekly-Hourly-Schedule-Template.xlsx"
  "smartsheet-project-schedule-gantt-sample.xlsx|https://www.smartsheet.com/sites/default/files/2023-03/IC-Sample-Project-Schedule-Template-with-Task-List-and-Gantt-Chart-11697.xlsx"
  "smartsheet-gantt-with-dependencies.xlsx|https://www.smartsheet.com/sites/default/files/2020-09/IC-Excel-Gantt-Chart-Template-with-Dependencies-10889.xlsx"
  "smartsheet-inventory-management.xlsx|https://www.smartsheet.com/sites/default/files/IC-Inventory-Management-Template-Updated-8857.xlsx"
  "smartsheet-inventory-stock-control.xlsx|https://www.smartsheet.com/sites/default/files/3-Inventory-Stock-Control-Template-DE1.xlsx"
  "smartsheet-employee-expense-report-example.xlsx|https://www.smartsheet.com/sites/default/files/2022-12/IC-Employee-Expense-Report-with-Mileage-11656.xlsx"
  "smartsheet-vacation-schedule.xlsx|https://www.smartsheet.com/media/files/ic-vacation-schedule-template.xlsx"
  "smartsheet-employee-attendance.xlsx|https://www.smartsheet.com/sites/default/files/2025-07/IC-Employee-Attendance-Spreadsheet-Template-8872.xlsx"
  "smartsheet-monthly-work-schedule-example.xlsx|https://www.smartsheet.com/media/files/ic-monthly-work-schedule-template-example.xlsx"
)

for entry in "${files[@]}"; do
  name="${entry%%|*}"
  url="${entry#*|}"
  if [[ -f "$name" ]]; then
    echo "skip   $name (exists)"
    continue
  fi
  echo "fetch  $name"
  curl -fsSL -o "$name" "$url"
  # sanity check: xlsx files are zip archives (start with PK)
  if [[ "$(head -c 2 "$name")" != "PK" ]]; then
    echo "WARNING: $name does not look like a valid xlsx (URL moved?)" >&2
    mv "$name" "$name.invalid"
  fi
done
echo "done."
