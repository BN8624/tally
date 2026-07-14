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
BLUE = "D9EAF7"
LIGHT_BLUE = "EAF3F8"
ORANGE = "FCE4D6"
RED = "F4CCCC"
GREEN = "D9EAD3"
GRAY = "E7E6E6"
WHITE = "FFFFFF"
THIN_GRAY = Side(style="thin", color="B7B7B7")
MEDIUM_NAVY = Side(style="medium", color=NAVY)


def _excel_value(value: object) -> object:
    if value is None or pd.isna(value):
        return ""
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _set_title(sheet, text: str, end_column: int) -> None:
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=end_column)
    cell = sheet.cell(1, 1, text)
    cell.font = Font(name="맑은 고딕", size=16, bold=True, color=WHITE)
    cell.fill = PatternFill("solid", fgColor=NAVY)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 30


def _style_section(sheet, row: int, text: str, end_column: int, color: str = BLUE) -> None:
    sheet.merge_cells(start_row=row, start_column=1, end_row=row, end_column=end_column)
    cell = sheet.cell(row, 1, text)
    cell.font = Font(name="맑은 고딕", size=12, bold=True, color=NAVY)
    cell.fill = PatternFill("solid", fgColor=color)
    cell.alignment = Alignment(horizontal="left", vertical="center")
    cell.border = Border(bottom=MEDIUM_NAVY)
    sheet.row_dimensions[row].height = 23


def _lookup(frame: pd.DataFrame, key_column: str, key: str, month: str, value: str) -> object:
    rows = frame[(frame[key_column].eq(key)) & (frame["month"].eq(month))]
    if rows.empty:
        return 0
    return _excel_value(rows.iloc[0][value])


def _write_month_headers(sheet, row: int, months: list[str]) -> tuple[int, int]:
    sheet.cell(row, 1, "항목")
    column = 2
    for month in months:
        sheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 3)
        sheet.cell(row, column, month)
        for offset, label in enumerate(("매수", "공급가액", "세액", "합계금액")):
            sheet.cell(row + 1, column + offset, label)
        column += 4
    total_start = column
    sheet.merge_cells(start_row=row, start_column=column, end_row=row, end_column=column + 3)
    sheet.cell(row, column, "합계")
    for offset, label in enumerate(("매수", "공급가액", "세액", "합계금액")):
        sheet.cell(row + 1, column + offset, label)

    end_column = column + 3
    for header_row in (row, row + 1):
        for cell in sheet[header_row][:end_column]:
            cell.fill = PatternFill("solid", fgColor=NAVY)
            cell.font = Font(name="맑은 고딕", bold=True, color=WHITE)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)
    sheet.merge_cells(start_row=row, start_column=1, end_row=row + 1, end_column=1)
    sheet.cell(row, 1).alignment = Alignment(horizontal="center", vertical="center")
    return total_start, end_column


def _write_summary_row(
    sheet,
    row: int,
    label: str,
    frame: pd.DataFrame,
    key_column: str,
    months: list[str],
    total_start: int,
    *,
    emphasis: bool = False,
    alert: bool = False,
) -> None:
    sheet.cell(row, 1, label)
    sheet.cell(row, 1).font = Font(name="맑은 고딕", bold=emphasis)
    for month_index, month in enumerate(months):
        start = 2 + month_index * 4
        for offset, value_name in enumerate(("count", "supply_amount", "tax_amount", "total_amount")):
            sheet.cell(row, start + offset, _lookup(frame, key_column, label, month, value_name))
    for offset in range(4):
        source_columns = [get_column_letter(2 + month_index * 4 + offset) for month_index in range(len(months))]
        formula = "+".join(f"{column}{row}" for column in source_columns) or "0"
        sheet.cell(row, total_start + offset, f"={formula}")

    end_column = total_start + 3
    fill_color = RED if alert else (LIGHT_BLUE if emphasis else WHITE)
    for cell in sheet[row][:end_column]:
        cell.fill = PatternFill("solid", fgColor=fill_color)
        cell.border = Border(bottom=THIN_GRAY)
        cell.alignment = Alignment(horizontal="right" if cell.column > 1 else "left", vertical="center")
        if cell.column > 1:
            cell.number_format = "#,##0;[Red]-#,##0"


def _ordered_categories(settings: CompanySettings, frame: pd.DataFrame) -> list[str]:
    preferred = [settings.account_146_label, "원재료(도급)", "제조경비", "도급경비", "기타", "고정"]
    present = set(frame["account_category"].dropna().astype(str))
    return [category for category in preferred if category in present or category == settings.account_146_label]


def _build_summary(workbook: Workbook, result: ProcessingResult, settings: CompanySettings) -> None:
    sheet = workbook.active
    sheet.title = "집계표"
    months = sorted(result.transactions["month"].dropna().astype(str).unique())
    end_column = 1 + (len(months) + 1) * 4
    _set_title(sheet, f"{settings.name} 부가세 월별 집계표", end_column)
    period = " ~ ".join((months[0], months[-1])) if months else "거래 없음"
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=end_column)
    sheet.cell(2, 1, f"신고 기간 {period} · 검산 {'완료' if result.validation_passed else '실패'}")
    sheet.cell(2, 1).alignment = Alignment(horizontal="center")
    sheet.cell(2, 1).font = Font(name="맑은 고딕", bold=True, color="38761D" if result.validation_passed else "9C0006")

    row = 4
    _style_section(sheet, row, "매입 계정 분류별 월별 집계", end_column)
    row += 1
    total_start, _ = _write_month_headers(sheet, row, months)
    row += 2
    for category in _ordered_categories(settings, result.purchase_by_category):
        _write_summary_row(
            sheet, row, category, result.purchase_by_category, "account_category", months, total_start
        )
        row += 1

    row += 1
    _style_section(sheet, row, "매입 요약", end_column)
    row += 1
    _write_month_headers(sheet, row, months)
    row += 2
    purchase_items = [
        "일반매입",
        "고정",
        "세금계산서 매입계",
        "카과",
        "현과",
        "카드매입",
        "과세 매입 총계",
        "불공",
        "과매계",
        "면세 매입",
    ]
    for item in purchase_items:
        _write_summary_row(
            sheet,
            row,
            item,
            result.purchase_summary,
            "item",
            months,
            total_start,
            emphasis=item in {"세금계산서 매입계", "카드매입", "과세 매입 총계", "과매계"},
            alert=item == "불공",
        )
        row += 1

    row += 1
    _style_section(sheet, row, "매출 계정별 월별 집계", end_column, color=ORANGE)
    row += 1
    _write_month_headers(sheet, row, months)
    row += 2
    accounts = sorted(result.sales_by_account["account_name"].dropna().astype(str).unique())
    for account in accounts:
        _write_summary_row(sheet, row, account, result.sales_by_account, "account_name", months, total_start)
        row += 1

    row += 1
    _style_section(sheet, row, "매출 요약 및 카드·현영 보조 집계", end_column, color=ORANGE)
    row += 1
    _write_month_headers(sheet, row, months)
    row += 2
    for item in ("과세매출 총계", "면세 매출", "카드매출", "현영매출"):
        _write_summary_row(
            sheet,
            row,
            item,
            result.sales_summary,
            "item",
            months,
            total_start,
            emphasis=item == "과세매출 총계",
        )
        row += 1

    sheet.freeze_panes = "B7"
    sheet.sheet_view.showGridLines = False
    sheet.column_dimensions["A"].width = 24
    for column in range(2, end_column + 1):
        sheet.column_dimensions[get_column_letter(column)].width = 13
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = f"A1:{get_column_letter(end_column)}{row}"
    sheet.auto_filter.ref = f"A5:{get_column_letter(end_column)}{row}"


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
