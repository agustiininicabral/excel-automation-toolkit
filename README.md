# Excel Automation Toolkit

Automates the creation of formatted Excel business reports from structured sales CSV files. The toolkit validates input, separates rejected rows, creates summary sheets and embeds a monthly revenue chart.

## Features

- CSV extraction with required-column and row-level validation.
- Duplicate order detection, ISO date parsing and positive monetary amount checks.
- Formatted `Sales Data`, `Summary` and `Data Quality` worksheets.
- Monthly revenue aggregation and an embedded Excel bar chart.
- Freeze panes, filters, readable column widths, number/date formats and styled headers.
- Deterministic sample data, CLI and automated workbook inspection tests.

## Run

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
# macOS/Linux: source .venv/bin/activate
pip install -e ".[dev]"
python -m excel_toolkit --input data/raw/sales.csv --output output/sales_report.xlsx
pytest
```

Open the generated workbook in Excel or LibreOffice. Values in the summary are computed by Python before writing, so the report does not depend on Excel formula recalculation.

## Input contract

Required CSV columns: `order_id,date,customer,category,amount`. Order IDs must be unique, dates must be ISO `YYYY-MM-DD`, and amounts must be positive finite numbers. Rows failing validation are preserved in the Data Quality sheet with their source line and rejection reason.

## Output sheets

- **Sales Data** — accepted transactions with typed dates and numeric amounts.
- **Summary** — KPI totals, monthly revenue and a monthly revenue chart.
- **Data Quality** — rejected input rows and explanations.

## Project structure

- `src/excel_toolkit/report.py`: validation, aggregation and workbook generation.
- `data/raw/sales.csv`: synthetic demonstration input.
- `tests/`: unit and integration tests.
- `.github/workflows/ci.yml`: Python 3.11–3.13 CI.

## Limitations

This project generates new reports from CSV input; it does not modify arbitrary existing workbooks or preserve macros. Synthetic data is used. No Excel installation is required to generate the workbook.

## License

MIT. See [LICENSE](LICENSE).
