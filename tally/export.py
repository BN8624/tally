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
DOUBLE_NAVY = Side(style="double", color=NAVY)
BLOCK_START_COLUMNS = (1, 7, 13)
SUMMARY_END_COLUMN = 17


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


def _month_label(month: str, months: list[str]) -> str:
    years = {value[:4] for value in months}
    if len(years) > 1:
        return f"{month[:4]}.{int(month[5:])}월"
    return f"{int(month[5:])}월"


def _write_month_block(
    sheet,
    start_row: int,
    start_column: int,
    title: str,
    frame: pd.DataFrame,
    key_column: str,
    key: str,
    months: list[str],
    *,
    title_fill: str,
    emphasis: bool = False,
    alert: bool = False,
) -> int:
    end_column = start_column + 4
    fill = RED if alert else (GREEN if emphasis else title_fill)
    sheet.merge_cells(
        start_row=start_row,
        start_column=start_column,
        end_row=start_row,
        end_column=end_column,
    )
    title_cell = sheet.cell(start_row, start_column, title)
    title_cell.fill = PatternFill("solid", fgColor=fill)
    title_cell.font = Font(name="맑은 고딕", size=11, bold=True, color=NAVY)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = Border(bottom=MEDIUM_NAVY)

    header_row = start_row + 1
    for offset, label in enumerate(("월", "매수", "공급가액", "세액", "합계금액")):
        cell = sheet.cell(header_row, start_column + offset, label)
        cell.fill = PatternFill("solid", fgColor=GRAY)
        cell.font = Font(name="맑은 고딕", size=9, bold=True, color=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=THIN_GRAY)

    first_data_row = header_row + 1
    for month_index, month in enumerate(months):
        row = first_data_row + month_index
        sheet.cell(row, start_column, _month_label(month, months))
        for offset, value_name in enumerate(
            ("count", "supply_amount", "tax_amount", "total_amount"),
            start=1,
        ):
            sheet.cell(
                row,
                start_column + offset,
                _lookup(frame, key_column, key, month, value_name),
            )
        for cell in sheet[row][start_column - 1 : end_column]:
            cell.border = Border(bottom=THIN_GRAY)
            cell.alignment = Alignment(
                horizontal="right" if cell.column > start_column else "center",
                vertical="center",
            )
        sheet.cell(row, start_column + 1).number_format = '"("0")"'
        for column in range(start_column + 2, end_column + 1):
            sheet.cell(row, column).number_format = "#,##0;[Red]-#,##0"

    total_row = first_data_row + len(months)
    sheet.cell(total_row, start_column, "계")
    for column in range(start_column + 1, end_column + 1):
        letter = get_column_letter(column)
        if months:
            sheet.cell(total_row, column, f"=SUM({letter}{first_data_row}:{letter}{total_row - 1})")
        else:
            sheet.cell(total_row, column, 0)
    for cell in sheet[total_row][start_column - 1 : end_column]:
        cell.fill = PatternFill("solid", fgColor=LIGHT_BLUE if not alert else RED)
        cell.font = Font(name="맑은 고딕", bold=True)
        cell.border = Border(top=MEDIUM_NAVY, bottom=DOUBLE_NAVY)
        cell.alignment = Alignment(
            horizontal="right" if cell.column > start_column else "center",
            vertical="center",
        )
    sheet.cell(total_row, start_column + 1).number_format = '"("0")"'
    for column in range(start_column + 2, end_column + 1):
        sheet.cell(total_row, column).number_format = "#,##0;[Red]-#,##0"
    return total_row


def _write_block_grid(
    sheet,
    start_row: int,
    items: list[str],
    frame: pd.DataFrame,
    key_column: str,
    months: list[str],
    *,
    title_fill: str,
    emphasis: set[str] | None = None,
    alerts: set[str] | None = None,
) -> int:
    if not items:
        return start_row
    emphasis = emphasis or set()
    alerts = alerts or set()
    block_height = len(months) + 3
    for index, item in enumerate(items):
        band = index // len(BLOCK_START_COLUMNS)
        slot = index % len(BLOCK_START_COLUMNS)
        block_row = start_row + band * (block_height + 1)
        _write_month_block(
            sheet,
            block_row,
            BLOCK_START_COLUMNS[slot],
            item,
            frame,
            key_column,
            item,
            months,
            title_fill=title_fill,
            emphasis=item in emphasis,
            alert=item in alerts,
        )
    band_count = (len(items) + len(BLOCK_START_COLUMNS) - 1) // len(BLOCK_START_COLUMNS)
    return start_row + band_count * (block_height + 1) - 1


def _ordered_categories(settings: CompanySettings, frame: pd.DataFrame) -> list[str]:
    preferred = [settings.account_146_label, "원재료(도급)", "제조경비", "도급경비", "기타", "고정"]
    present = set(frame["account_category"].dropna().astype(str))
    return [category for category in preferred if category in present or category == settings.account_146_label]


def _build_summary(workbook: Workbook, result: ProcessingResult, settings: CompanySettings) -> None:
    sheet = workbook.active
    sheet.title = "집계표"
    months = sorted(result.transactions["month"].dropna().astype(str).unique())
    _set_title(sheet, f"{settings.name} 부가세 월별 집계표", SUMMARY_END_COLUMN)
    period = " ~ ".join((months[0], months[-1])) if months else "거래 없음"
    sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=SUMMARY_END_COLUMN)
    sheet.cell(2, 1, f"신고 기간 {period} · 검산 {'완료' if result.validation_passed else '실패'}")
    sheet.cell(2, 1).alignment = Alignment(horizontal="center")
    sheet.cell(2, 1).font = Font(name="맑은 고딕", bold=True, color="38761D" if result.validation_passed else "9C0006")

    row = 4
    _style_section(sheet, row, "① 매입 계정 분류별 월별 집계", SUMMARY_END_COLUMN)
    row += 1
    row = _write_block_grid(
        sheet,
        row,
        _ordered_categories(settings, result.purchase_by_category),
        result.purchase_by_category,
        "account_category",
        months,
        title_fill=BLUE,
    )

    row += 1
    _style_section(sheet, row, "② 매입 요약 · 불공 차감", SUMMARY_END_COLUMN)
    row += 1
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
    row = _write_block_grid(
        sheet,
        row,
        purchase_items,
        result.purchase_summary,
        "item",
        months,
        title_fill=BLUE,
        emphasis={"세금계산서 매입계", "카드매입", "과세 매입 총계", "과매계"},
        alerts={"불공"},
    )

    row += 1
    _style_section(sheet, row, "③ 매출 계정별 월별 집계", SUMMARY_END_COLUMN, color=ORANGE)
    row += 1
    accounts = sorted(result.sales_by_account["account_name"].dropna().astype(str).unique())
    row = _write_block_grid(
        sheet,
        row,
        accounts,
        result.sales_by_account,
        "account_name",
        months,
        title_fill=ORANGE,
    )

    row += 1
    _style_section(sheet, row, "④ 매출 요약 · 카드·현영 보조 집계", SUMMARY_END_COLUMN, color=ORANGE)
    row += 1
    row = _write_block_grid(
        sheet,
        row,
        ["과세매출 총계", "면세 매출", "카드매출", "현영매출"],
        result.sales_summary,
        "item",
        months,
        title_fill=ORANGE,
        emphasis={"과세매출 총계"},
    )

    sheet.freeze_panes = "A5"
    sheet.sheet_view.showGridLines = False
    for start_column in BLOCK_START_COLUMNS:
        widths = (10, 9, 16, 14, 17)
        for offset, width in enumerate(widths):
            sheet.column_dimensions[get_column_letter(start_column + offset)].width = width
    for gap_column in (6, 12):
        sheet.column_dimensions[get_column_letter(gap_column)].width = 3
    sheet.page_setup.orientation = "portrait"
    sheet.page_setup.paperWidth = "170mm"
    sheet.page_setup.paperHeight = "240mm"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 1
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.page_margins.left = 0.2
    sheet.page_margins.right = 0.2
    sheet.page_margins.top = 0.25
    sheet.page_margins.bottom = 0.25
    sheet.page_margins.header = 0.1
    sheet.page_margins.footer = 0.1
    sheet.print_options.horizontalCentered = True
    sheet.print_title_rows = "1:2"
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
