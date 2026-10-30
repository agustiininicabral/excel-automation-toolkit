from __future__ import annotations

import argparse
import csv
from collections import defaultdict
from datetime import date
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


REQUIRED_COLUMNS = {"order_id", "date", "customer", "category", "amount"}
ROOT = Path(__file__).resolve().parents[2]


def read_and_validate(csv_path: Path) -> tuple[list[dict[str, Any]], list[dict[str, str]]]:
    """Read a sales CSV and split valid typed records from rejected source rows."""
    with csv_path.open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        if reader.fieldnames is None or not REQUIRED_COLUMNS.issubset(set(reader.fieldnames)):
            missing = sorted(REQUIRED_COLUMNS - set(reader.fieldnames or []))
            raise ValueError(f"Missing required columns: {', '.join(missing)}")
        valid: list[dict[str, Any]] = []
        rejected: list[dict[str, str]] = []
        seen_ids: set[str] = set()
        for line_number, row in enumerate(reader, start=2):
            try:
                order_id = str(row.get("order_id", "")).strip()
                customer = str(row.get("customer", "")).strip()
                category = str(row.get("category", "")).strip()
                if not order_id or not customer or not category:
                    raise ValueError("order_id, customer and category are required")
                if order_id in seen_ids:
                    raise ValueError("duplicate order_id")
                order_date = date.fromisoformat(str(row.get("date", "")).strip())
                amount = Decimal(str(row.get("amount", "")).strip()).quantize(Decimal("0.01"))
                if not amount.is_finite() or amount <= 0:
                    raise ValueError("amount must be a positive finite number")
                seen_ids.add(order_id)
                valid.append({
                    "order_id": order_id,
                    "date": order_date,
                    "customer": customer,
                    "category": category,
                    "amount": amount,
                })
            except (ValueError, InvalidOperation) as exc:
                rejected.append({
                    "source_row": str(line_number),
                    "order_id": str(row.get("order_id", "")),
                    "reason": str(exc),
                })
    return valid, rejected


def summarize(records: list[dict[str, Any]]) -> dict[str, Any]:
    monthly: dict[str, Decimal] = defaultdict(Decimal)
    category_totals: dict[str, Decimal] = defaultdict(Decimal)
    for row in records:
        monthly[row["date"].strftime("%Y-%m")] += row["amount"]
        category_totals[row["category"]] += row["amount"]
    return {
        "transaction_count": len(records),
        "total_revenue": sum((row["amount"] for row in records), Decimal("0.00")),
        "average_order_value": (
            sum((row["amount"] for row in records), Decimal("0.00")) / len(records)
            if records else Decimal("0.00")
        ),
        "monthly_revenue": sorted(monthly.items()),
        "category_revenue": sorted(category_totals.items(), key=lambda item: item[1], reverse=True),
    }


def _style_sheet(sheet, header_row: int = 1) -> None:
    header_fill = PatternFill("solid", fgColor="17365D")
    for cell in sheet[header_row]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")
    sheet.freeze_panes = f"A{header_row + 1}"
    if sheet.max_row >= header_row:
        sheet.auto_filter.ref = sheet.dimensions
    for column_cells in sheet.columns:
        max_length = max((len(str(cell.value or "")) for cell in column_cells), default=0)
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = min(max(max_length + 2, 12), 34)


def create_workbook(
    output_path: Path,
    records: list[dict[str, Any]],
    rejected: list[dict[str, str]],
) -> dict[str, Any]:
    summary = summarize(records)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    data_sheet = workbook.active
    data_sheet.title = "Sales Data"
    data_sheet.append(["Order ID", "Date", "Customer", "Category", "Amount"])
    for row in records:
        data_sheet.append([row["order_id"], row["date"], row["customer"], row["category"], float(row["amount"])])
    _style_sheet(data_sheet)
    for cell in data_sheet["B"][1:]:
        cell.number_format = "yyyy-mm-dd"
    for cell in data_sheet["E"][1:]:
        cell.number_format = '#,##0.00'

    summary_sheet = workbook.create_sheet("Summary")
    summary_sheet.append(["Metric", "Value"])
    summary_sheet.append(["Transactions", summary["transaction_count"]])
    summary_sheet.append(["Total Revenue", float(summary["total_revenue"])])
    summary_sheet.append(["Average Order Value", float(summary["average_order_value"])])
    summary_sheet["B3"].number_format = '#,##0.00'
    summary_sheet["B4"].number_format = '#,##0.00'
    summary_sheet.append([])
    summary_sheet.append(["Month", "Revenue"])
    for month, revenue in summary["monthly_revenue"]:
        summary_sheet.append([month, float(revenue)])
    _style_sheet(summary_sheet, 1)
    summary_sheet["A6"].font = Font(bold=True)
    summary_sheet["B6"].font = Font(bold=True)
    if summary["monthly_revenue"]:
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Monthly Revenue"
        chart.y_axis.title = "Revenue"
        chart.x_axis.title = "Month"
        data = Reference(summary_sheet, min_col=2, min_row=6, max_row=6 + len(summary["monthly_revenue"]))
        categories = Reference(summary_sheet, min_col=1, min_row=7, max_row=6 + len(summary["monthly_revenue"]))
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        chart.height = 8
        chart.width = 16
        summary_sheet.add_chart(chart, "D2")

    quality_sheet = workbook.create_sheet("Data Quality")
    quality_sheet.append(["Source Row", "Order ID", "Rejection Reason"])
    for row in rejected:
        quality_sheet.append([row["source_row"], row["order_id"], row["reason"]])
    _style_sheet(quality_sheet)

    workbook.save(output_path)
    return {
        "output": str(output_path),
        "accepted_rows": len(records),
        "rejected_rows": len(rejected),
        "total_revenue": str(summary["total_revenue"].quantize(Decimal("0.01"))),
        "monthly_periods": len(summary["monthly_revenue"]),
    }


def build_report(input_path: Path, output_path: Path) -> dict[str, Any]:
    records, rejected = read_and_validate(input_path)
    return create_workbook(output_path, records, rejected)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a validated Excel sales report.")
    parser.add_argument("--input", type=Path, default=ROOT / "data" / "raw" / "sales.csv")
    parser.add_argument("--output", type=Path, default=ROOT / "output" / "sales_report.xlsx")
    args = parser.parse_args()
    report = build_report(args.input, args.output)
    for key, value in report.items():
        print(f"{key}: {value}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
