# 처리 결과를 인쇄 가능한 집계표와 검토·상세·검산 시트로 출력합니다.
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .core import ProcessingResult
from .settings import CompanySettings


NAVY = "1F4E78"
RED = "F4CCCC"
WHITE = "FFFFFF"
THIN_GRAY = Side(style="thin", color="B7B7B7")
LEDGER_INK = "404040"
LEDGER_RULE = Side(style="thin", color="B7B7B7")
LEDGER_TOTAL_RULE = Side(style="medium", color="666666")
BLOCK_START_COLUMNS = (1, 5, 8, 11)
SUMMARY_END_COLUMN = 13


def _excel_value(value: object) -> object:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _period_label(months: list[str]) -> str:
    if not months:
        return "집계 기간 없음"
    years = {month[:4] for month in months}
    month_numbers = {int(month[5:]) for month in months}
    if len(years) == 1:
        year = months[0][:4]
        periods = (
            ({1, 2, 3}, "1기 예정"),
            ({4, 5, 6}, "1기 확정"),
            ({7, 8, 9}, "2기 예정"),
            ({10, 11, 12}, "2기 확정"),
        )
        for period_months, label in periods:
            if month_numbers and month_numbers.issubset(period_months):
                return f"{year}년 {label}"
    return f"{months[0]} ~ {months[-1]}"


def _set_summary_heading(sheet, company_name: str, months: list[str]) -> None:
    headings = (
        (1, 4, _period_label(months), "left", 10),
        (5, 9, company_name, "center", 12),
        (10, 13, "부가세 집계표", "right", 10),
    )
    for start_column, end_column, text, alignment, size in headings:
        sheet.merge_cells(
            start_row=1,
            start_column=start_column,
            end_row=1,
            end_column=end_column,
        )
        cell = sheet.cell(1, start_column, text)
        cell.font = Font(name="맑은 고딕", size=size, bold=True, color=LEDGER_INK)
        cell.alignment = Alignment(horizontal=alignment, vertical="center")
    for column in range(1, SUMMARY_END_COLUMN + 1):
        sheet.cell(1, column).border = Border(bottom=LEDGER_TOTAL_RULE)
    sheet.row_dimensions[1].height = 25
    sheet.row_dimensions[2].height = 8


def _set_title(sheet, text: str, end_column: int) -> None:
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_column)
    cell = sheet.cell(1, 1, text)
    cell.font = Font(name="맑은 고딕", size=16, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 30


def _lookup(frame: pd.DataFrame, key_column: str, key: str, month: str, value: str) -> object:
    rows = frame[(frame[key_column].eq(key)) & (frame["month"].eq(month))]
    if rows.empty:
        return 0
    return _excel_value(rows.iloc[0][value])


def _month_label(month: str, months: list[str]) -> str:
    years = {value[:4] for value in months}
    if len(years) > 1:
        return f"{month[:4]}.{int(month[5:])}월"
    return f"{int(month[5:])}월"


def _item_fields(mode: str) -> tuple[str | None, str | None, str | None]:
    if mode == "exempt":
        return ("count", "supply_amount", None)
    if mode == "total":
        return (None, "total_amount", None)
    return ("count", "supply_amount", "tax_amount")


def _write_ledger_grid(
    sheet,
    start_row: int,
    marker: str,
    items: list[tuple[str, str, str]],
    frame: pd.DataFrame,
    key_column: str,
    months: list[str],
) -> int:
    if not items:
        return start_row

    rows_per_band = len(months) + 3
    for band_start in range(0, len(items), len(BLOCK_START_COLUMNS)):
        band = band_start // len(BLOCK_START_COLUMNS)
        title_row = start_row + band * rows_per_band
        first_data_row = title_row + 1
        total_row = first_data_row + len(months)
        band_items = items[band_start : band_start + len(BLOCK_START_COLUMNS)]

        for row in range(title_row, total_row + 1):
            for column in range(1, SUMMARY_END_COLUMN + 1):
                cell = sheet.cell(row, column)
                cell.font = Font(name="맑은 고딕", size=9, color=LEDGER_INK)
                cell.border = Border(bottom=LEDGER_RULE)
                cell.alignment = Alignment(vertical="center")
            sheet.row_dimensions[row].height = 19

        for month_index, month in enumerate(months):
            month_cell = sheet.cell(first_data_row + month_index, 1, _month_label(month, months))
            month_cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.cell(total_row, 1, "계")
        sheet.cell(total_row, 1).alignment = Alignment(horizontal="center", vertical="center")

        for slot, (display, key, mode) in enumerate(band_items):
            start_column = BLOCK_START_COLUMNS[slot]
            end_column = start_column + (3 if slot == 0 else 2)
            value_start_column = start_column + 1 if slot == 0 else start_column
            sheet.merge_cells(
                start_row=title_row,
                start_column=start_column,
                end_row=title_row,
                end_column=end_column,
            )
            title = f"{marker}  {display}" if band_start == 0 and slot == 0 else display
            title_cell = sheet.cell(title_row, start_column, title)
            title_cell.font = Font(name="맑은 고딕", size=10, bold=True, color=LEDGER_INK)
            title_cell.alignment = Alignment(horizontal="center", vertical="center")
            title_cell.border = Border(bottom=LEDGER_TOTAL_RULE)

            fields = _item_fields(mode)
            for month_index, month in enumerate(months):
                row = first_data_row + month_index
                for offset, field in enumerate(fields):
                    column = value_start_column + offset
                    if field is not None:
                        sheet.cell(row, column, _lookup(frame, key_column, key, month, field))
                    cell = sheet.cell(row, column)
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    cell.number_format = '"("0")"' if field == "count" else "#,##0;[Red]-#,##0"

            for offset, field in enumerate(fields):
                column = value_start_column + offset
                cell = sheet.cell(total_row, column)
                if field is not None:
                    letter = get_column_letter(column)
                    cell.value = (
                        f"=SUM({letter}{first_data_row}:{letter}{total_row - 1})" if months else 0
                    )
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = '"("0")"' if field == "count" else "#,##0;[Red]-#,##0"

        for column in range(1, SUMMARY_END_COLUMN + 1):
            cell = sheet.cell(total_row, column)
            cell.font = Font(name="맑은 고딕", size=9, bold=True, color=LEDGER_INK)
            cell.border = Border(top=LEDGER_TOTAL_RULE, bottom=LEDGER_TOTAL_RULE)

    band_count = (len(items) + len(BLOCK_START_COLUMNS) - 1) // len(BLOCK_START_COLUMNS)
    return start_row + (band_count - 1) * rows_per_band + len(months) + 1


def _ordered_categories(settings: CompanySettings) -> list[str]:
    return [settings.account_146_label, "원재료(도급)", "제조경비", "도급경비", "기타", "고정"]


def _build_summary(workbook: Workbook, result: ProcessingResult, settings: CompanySettings) -> None:
    sheet = workbook.active
    sheet.title = "집계표"
    months = sorted(result.transactions["month"].dropna().astype(str).unique())
    _set_summary_heading(sheet, settings.name, months)

    row = 3
    category_items = [
        (category, category, "tax") for category in _ordered_categories(settings)
    ]
    row = _write_ledger_grid(
        sheet,
        row,
        "①",
        category_items,
        result.purchase_by_category,
        "account_category",
        months,
    )

    row += 2
    purchase_items = [
        ("일반", "일반매입", "tax"),
        ("고정", "고정", "tax"),
        ("계", "세금계산서 매입계", "tax"),
        ("카과", "카과", "tax"),
        ("현과", "현과", "tax"),
        ("카드계", "카드매입", "tax"),
        ("과세계", "과세 매입 총계", "tax"),
        ("불공", "불공", "tax"),
        ("과매계", "과매계", "tax"),
        ("면세", "면세 매입", "exempt"),
    ]
    row = _write_ledger_grid(
        sheet,
        row,
        "②",
        purchase_items,
        result.purchase_summary,
        "item",
        months,
    )

    row += 2
    accounts = sorted(result.sales_by_account["account_name"].dropna().astype(str).unique())
    account_items = [(account, account, "tax") for account in accounts]
    row = _write_ledger_grid(
        sheet,
        row,
        "③",
        account_items,
        result.sales_by_account,
        "account_name",
        months,
    )

    row += 2
    row = _write_ledger_grid(
        sheet,
        row,
        "④",
        [
            ("과세계", "과세매출 총계", "tax"),
            ("면세", "면세 매출", "exempt"),
            ("카드", "카드매출", "total"),
            ("현영", "현영매출", "total"),
        ],
        result.sales_summary,
        "item",
        months,
    )

    for row_cells in sheet.iter_rows(
        min_row=1,
        max_row=row,
        min_col=1,
        max_col=SUMMARY_END_COLUMN,
    ):
        for cell in row_cells:
            cell.fill = PatternFill("solid", fgColor=WHITE)

    sheet.sheet_view.showGridLines = False
    widths = (7, 7, 14, 12, 7, 14, 12, 7, 14, 12, 7, 14, 12)
    for column, width in enumerate(widths, start=1):
        sheet.column_dimensions[get_column_letter(column)].width = width
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperWidth = "170mm"
    sheet.page_setup.paperHeight = "240mm"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins.left = 0.15
    sheet.page_margins.right = 0.15
    sheet.page_margins.top = 0.2
    sheet.page_margins.bottom = 0.2
    sheet.page_margins.header = 0
    sheet.page_margins.footer = 0
    sheet.print_options.horizontalCentered = True
    sheet.print_area = f"A1:{get_column_letter(SUMMARY_END_COLUMN)}{row}"


def _write_table_sheet(
    workbook: Workbook,
    name: str,
    title: str,
    headers: list[tuple[str, str]],
    frame: pd.DataFrame,
    *,
    alert_column: str | None = None,
) -> None:
    sheet = workbook.create_sheet(name)
    _set_title(sheet, title, len(headers))
    for column, (_, label) in enumerate(headers, start=1):
        cell = sheet.cell(3, column, label)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="맑은 고딕", bold=True, color=WHITE)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row_number, (_, record) in enumerate(frame.iterrows(), start=4):
        for column, (field, _) in enumerate(headers, start=1):
            value = _excel_value(record.get(field, ""))
            cell = sheet.cell(row_number, column, value)
            cell.border = Border(bottom=THIN_GRAY)
            cell.alignment = Alignment(vertical="top", wrap_text=field in {"item", "candidate_reason", "nondeductible_reason", "review_memo", "detail"})
            if field in {"supply_amount", "tax_amount", "total_amount", "expected", "actual", "difference"}:
                cell.number_format = "#,##0;[Red]-#,##0"
            if field == "date" and isinstance(value, (date, datetime)):
                cell.number_format = "yyyy-mm-dd"
        if alert_column and str(record.get(alert_column, "")) in {"실패", "판단 보류", "미분류"}:
            for cell in sheet[row_number]:
                cell.fill = PatternFill("solid", fgColor=RED)

    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{max(3, 3 + len(frame))}"
    sheet.sheet_view.showGridLines = False
    for index, (field, label) in enumerate(headers, start=1):
        width = max(12, min(42, len(label) * 2 + 4))
        if field in {"vendor", "item", "candidate_reason", "nondeductible_reason", "review_memo", "detail"}:
            width = 28 if field in {"vendor", "item"} else 36
        sheet.column_dimensions[get_column_letter(index)].width = width


def export_workbook(
    result: ProcessingResult,
    settings: CompanySettings,
    output_path: str | Path,
) -> Path:
    """처리 결과를 정본에 정의된 다섯 개 시트로 저장합니다."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    _build_summary(workbook, result, settings)

    review_headers = [
        ("date", "날짜"),
        ("vendor", "거래처"),
        ("item", "품명"),
        ("account_code", "계정코드"),
        ("account_name", "계정과목"),
        ("supply_amount", "공급가액"),
        ("tax_amount", "세액"),
        ("original_type", "원본 유형"),
        ("candidate_reason", "후보 근거"),
        ("review_status", "최종 판정"),
        ("final_type", "최종 유형"),
        ("nondeductible_reason", "불공 사유"),
        ("review_memo", "메모"),
    ]
    _write_table_sheet(
        workbook,
        "불공검토",
        "불공 검토",
        review_headers,
        result.review,
        alert_column="review_status",
    )

    detail_headers = [
        ("row_id", "원본 행 ID"),
        ("date", "전표일자"),
        ("division", "매입·매출"),
        ("month", "월"),
        ("vendor", "거래처"),
        ("item", "품명"),
        ("supply_amount", "공급가액"),
        ("tax_amount", "세액"),
        ("total_amount", "합계금액"),
        ("account_code", "계정코드"),
        ("account_name", "계정과목"),
        ("account_category", "계정 분류"),
        ("original_type", "원본 유형"),
        ("final_type", "최종 유형"),
        ("nondeductible", "불공 여부"),
        ("nondeductible_reason", "불공 사유"),
        ("candidate_reason", "후보 근거"),
        ("review_status", "최종 판정"),
        ("review_memo", "메모"),
    ]
    _write_table_sheet(workbook, "거래상세", "원본 상세 거래 및 파생 값", detail_headers, result.transactions)

    unclassified = result.transactions[result.transactions["account_category"].eq("미분류")]
    _write_table_sheet(
        workbook,
        "미분류",
        "계정 미분류 거래",
        detail_headers,
        unclassified,
        alert_column="account_category",
    )

    validation_headers = [
        ("check", "검산 항목"),
        ("status", "상태"),
        ("expected", "기대값"),
        ("actual", "결과값"),
        ("difference", "차이"),
        ("detail", "확인 사항"),
    ]
    _write_table_sheet(
        workbook,
        "검산",
        f"검산 결과 · {'완료' if result.validation_passed else '실패'}",
        validation_headers,
        result.validation,
        alert_column="status",
    )
    workbook.save(path)
    return path
