# EOF-Script-blind-test-dashboard
EOF Script Blind Test Dashboard (v7) Internal review dashboard for survey results

## Rebuild Dashboard Data

When the survey Excel and qualitative analysis Excel are updated, run:

```bash
python3 scripts/rebuild_dashboard_data.py
```

The script will automatically find the latest matching `.xlsx` files in the repo and rebuild:

- `blind-test-dashboard-data-0401.js`
- `blind-test-open-response-data.js`
- `workbook_rows_0401.json`
- `workbook-open-response-rows.js`

You can also point it to specific files:

```bash
python3 scripts/rebuild_dashboard_data.py \
  --survey "/absolute/path/to/updated-survey.xlsx" \
  --analysis "/absolute/path/to/updated-analysis.xlsx"
```
