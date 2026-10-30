import csv
import sys
from datetime import date
from decimal import Decimal
from pathlib import Path

import pytest
from openpyxl import load_workbook

from excel_toolkit.report import build_report, create_workbook, read_and_validate, summarize, main


ROOT = Path(__file__).resolve().parents[1]


def test_sample_report_creates_three_sheets_and_chart(tmp_path):
    output = tmp_path / "report.xlsx"
    result = build_report(ROOT / "data" / "raw" / "sales.csv", output)
    assert result["accepted_rows"] == 8
    assert result["rejected_rows"] == 4
    workbook = load_workbook(output)
    assert workbook.sheetnames == ["Sales Data", "Summary", "Data Quality"]
    assert len(workbook["Summary"]._charts) == 1
    assert workbook["Sales Data"].max_row == 9
    assert workbook["Data Quality"].max_row == 5
    assert workbook["Sales Data"].freeze_panes == "A2"
    assert workbook["Sales Data"]["E2"].number_format == "#,##0.00"


def test_validation_rejects_missing_columns(tmp_path):
    source = tmp_path / "bad.csv"
    source.write_text("order_id,date\n1,2026-01-01\n", encoding="utf-8")
    with pytest.raises(ValueError, match="Missing required columns"):
        read_and_validate(source)


def test_validation_rejects_duplicate_and_invalid_rows(tmp_path):
    source = tmp_path / "sales.csv"
    source.write_text(
        "order_id,date,customer,category,amount\n"
        "A,2026-01-01,Ann,Books,10.00\n"
        "A,2026-01-02,Ann,Books,11.00\n"
        "B,2026-01-03,Bob,Home,-2.00\n",
        encoding="utf-8",
    )
    valid, rejected = read_and_validate(source)
    assert len(valid) == 1
    assert len(rejected) == 2
    assert rejected[0]["reason"] == "duplicate order_id"


def test_summary_handles_empty_records():
    result = summarize([])
    assert result["transaction_count"] == 0
    assert result["total_revenue"] == Decimal("0.00")
    assert result["monthly_revenue"] == []


def test_summary_aggregates_month_and_category():
    rows = [
        {"date": date(2026, 1, 1), "amount": Decimal("10.00"), "category": "Books"},
        {"date": date(2026, 1, 2), "amount": Decimal("15.00"), "category": "Books"},
        {"date": date(2026, 2, 1), "amount": Decimal("5.00"), "category": "Home"},
    ]
    result = summarize(rows)
    assert result["total_revenue"] == Decimal("30.00")
    assert result["monthly_revenue"] == [("2026-01", Decimal("25.00")), ("2026-02", Decimal("5.00"))]
    assert result["category_revenue"][0] == ("Books", Decimal("25.00"))


def test_empty_records_still_create_valid_workbook(tmp_path):
    output = tmp_path / "empty.xlsx"
    result = create_workbook(output, [], [])
    workbook = load_workbook(output)
    assert result["accepted_rows"] == 0
    assert len(workbook["Summary"]._charts) == 0
    assert workbook["Sales Data"].max_row == 1


def test_cli_entrypoint(tmp_path, monkeypatch):
    output = tmp_path / "cli.xlsx"
    monkeypatch.setattr(sys, "argv", ["excel-report", "--input", str(ROOT / "data" / "raw" / "sales.csv"), "--output", str(output)])
    assert main() == 0
    assert output.exists()
