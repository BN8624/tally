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
DETAIL_START_COLUMNS = (1, 5)
SUMMARY_END_COLUMN = 13
COUNT_FORMAT = '"("0")"'
MONEY_FORMAT = "#,##0_);[Red](#,##0)"


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


def _has_values(
    frame: pd.DataFrame,
    key_column: str,
    key: str,
    fields: tuple[str, ...] = ("count", "supply_amount", "tax_amount", "total_amount"),
) -> bool:
    rows = frame[frame[key_column].eq(key)]
    return any(
        field in rows and pd.to_numeric(rows[field], errors="coerce").fillna(0).ne(0).any()
        for field in fields
    )


def _total_value(frame: pd.DataFrame, key_column: str, key: str, field: str) -> object:
    rows = frame[frame[key_column].eq(key)]
    if rows.empty or field not in rows:
        return 0
    return _excel_value(pd.to_numeric(rows[field], errors="coerce").fillna(0).sum())


def _style_ledger_cell(cell, *, bold: bool = False, total: bool = False) -> None:
    cell.fill = PatternFill("solid", fgColor=WHITE)
    cell.font = Font(name="맑은 고딕", size=9, bold=bold, color=LEDGER_INK)
    cell.alignment = Alignment(vertical="center")
    cell.border = Border(
        top=LEDGER_TOTAL_RULE if total else None,
        bottom=LEDGER_TOTAL_RULE if total else LEDGER_RULE,
    )


def _write_month_table(
    sheet,
    title_row: int,
    start_column: int,
    title: str,
    frame: pd.DataFrame,
    key_column: str,
    key: str,
    months: list[str],
    *,
    marker: str = "",
    exempt: bool = False,
    title_start_column: int | None = None,
) -> int:
    title_start = title_start_column or start_column
    end_column = start_column + 3
    sheet.merge_cells(
        start_row=title_row,
        start_column=title_start,
        end_row=title_row,
        end_column=end_column,
    )
    title_cell = sheet.cell(title_row, title_start, f"{marker}  {title}" if marker else title)
    _style_ledger_cell(title_cell, bold=True)
    title_cell.font = Font(name="맑은 고딕", size=10, bold=True, color=LEDGER_INK)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.border = Border(bottom=LEDGER_TOTAL_RULE)

    first_data_row = title_row + 1
    fields: tuple[str | None, ...] = (
        "count",
        "supply_amount",
        None if exempt else "tax_amount",
    )
    for month_index, month in enumerate(months):
        row = first_data_row + month_index
        month_cell = sheet.cell(row, start_column, _month_label(month, months))
        _style_ledger_cell(month_cell)
        month_cell.alignment = Alignment(horizontal="center", vertical="center")
        for offset, field in enumerate(fields, start=1):
            cell = sheet.cell(row, start_column + offset)
            if field is not None:
                cell.value = _lookup(frame, key_column, key, month, field)
            _style_ledger_cell(cell)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = COUNT_FORMAT if field == "count" else MONEY_FORMAT

    total_row = first_data_row + len(months)
    total_label = sheet.cell(total_row, start_column, "계")
    _style_ledger_cell(total_label, bold=True, total=True)
    total_label.alignment = Alignment(horizontal="center", vertical="center")
    for offset, field in enumerate(fields, start=1):
        column = start_column + offset
        cell = sheet.cell(total_row, column)
        if field is not None:
            letter = get_column_letter(column)
            cell.value = f"=SUM({letter}{first_data_row}:{letter}{total_row - 1})" if months else 0
        _style_ledger_cell(cell, bold=True, total=True)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = COUNT_FORMAT if field == "count" else MONEY_FORMAT
    return total_row


def _write_compact_month_table(
    sheet,
    title_row: int,
    start_column: int,
    title: str,
    frame: pd.DataFrame,
    key_column: str,
    key: str,
    months: list[str],
) -> int:
    sheet.merge_cells(
        start_row=title_row,
        start_column=start_column,
        end_row=title_row,
        end_column=start_column + 2,
    )
    title_cell = sheet.cell(title_row, start_column, title)
    _style_ledger_cell(title_cell, bold=True)
    title_cell.font = Font(name="맑은 고딕", size=10, bold=True, color=LEDGER_INK)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")

    first_data_row = title_row + 1
    fields = ("count", "supply_amount", "tax_amount")
    for month_index, month in enumerate(months):
        row = first_data_row + month_index
        for offset, field in enumerate(fields):
            cell = sheet.cell(row, start_column + offset, _lookup(frame, key_column, key, month, field))
            _style_ledger_cell(cell)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = COUNT_FORMAT if field == "count" else MONEY_FORMAT

    total_row = first_data_row + len(months)
    for offset, field in enumerate(fields):
        column = start_column + offset
        letter = get_column_letter(column)
        cell = sheet.cell(
            total_row,
            column,
            f"=SUM({letter}{first_data_row}:{letter}{total_row - 1})" if months else 0,
        )
        _style_ledger_cell(cell, bold=True, total=True)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = COUNT_FORMAT if field == "count" else MONEY_FORMAT
    return total_row


def _write_total_amount_table(
    sheet,
    title_row: int,
    items: list[tuple[str, str]],
    frame: pd.DataFrame,
    months: list[str],
) -> tuple[int, list[int]]:
    for offset, (title, _key) in enumerate(items, start=2):
        cell = sheet.cell(title_row, offset, title)
        _style_ledger_cell(cell, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=LEDGER_TOTAL_RULE)

    first_data_row = title_row + 1
    for month_index, month in enumerate(months):
        row = first_data_row + month_index
        month_cell = sheet.cell(row, 1, _month_label(month, months))
        _style_ledger_cell(month_cell)
        month_cell.alignment = Alignment(horizontal="center", vertical="center")
        for column, (_title, key) in enumerate(items, start=2):
            cell = sheet.cell(row, column, _lookup(frame, "item", key, month, "total_amount"))
            _style_ledger_cell(cell)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = MONEY_FORMAT

    total_row = first_data_row + len(months)
    label = sheet.cell(total_row, 1, "계")
    _style_ledger_cell(label, bold=True, total=True)
    label.alignment = Alignment(horizontal="center", vertical="center")
    total_columns: list[int] = []
    for column in range(2, 2 + len(items)):
        letter = get_column_letter(column)
        cell = sheet.cell(
            total_row,
            column,
            f"=SUM({letter}{first_data_row}:{letter}{total_row - 1})" if months else 0,
        )
        _style_ledger_cell(cell, bold=True, total=True)
        cell.alignment = Alignment(horizontal="right", vertical="center")
        cell.number_format = MONEY_FORMAT
        total_columns.append(column)
    return total_row, total_columns


def _ordered_categories(settings: CompanySettings) -> list[str]:
    return [settings.account_146_label, "원재료(도급)", "제조경비", "도급경비", "기타", "고정"]


def _build_summary(workbook: Workbook, result: ProcessingResult, settings: CompanySettings) -> None:
    sheet = workbook.active
    sheet.title = "집계표"
    months = sorted(result.transactions["month"].dropna().astype(str).unique())
    _set_summary_heading(sheet, settings.name, months)

    purchase_items: list[tuple[str, pd.DataFrame, str, str, bool]] = []
    for category in _ordered_categories(settings):
        if _has_values(result.purchase_by_category, "account_category", category):
            purchase_items.append(
                (category, result.purchase_by_category, "account_category", category, False)
            )

    purchase_title_row = 3
    rows_per_band = len(months) + 3
    last_purchase_detail_row = purchase_title_row - 1
    for item_index in range(0, len(purchase_items), len(DETAIL_START_COLUMNS)):
        band = item_index // len(DETAIL_START_COLUMNS)
        title_row = purchase_title_row + band * rows_per_band
        for slot, (title, frame, key_column, key, exempt) in enumerate(
            purchase_items[item_index : item_index + len(DETAIL_START_COLUMNS)]
        ):
            start_column = DETAIL_START_COLUMNS[slot]
            total_row = _write_month_table(
                sheet,
                title_row,
                start_column,
                title,
                frame,
                key_column,
                key,
                months,
                marker="①" if item_index == 0 and slot == 0 else "",
                exempt=exempt,
            )
            last_purchase_detail_row = max(last_purchase_detail_row, total_row)

    purchase_total_row: int | None = None
    purchase_summary_column = 9
    if _has_values(result.purchase_summary, "item", "세금계산서 매입계"):
        purchase_total_row = _write_compact_month_table(
            sheet,
            purchase_title_row,
            purchase_summary_column,
            "①  세금계산서 매입계" if not purchase_items else "세금계산서 매입계",
            result.purchase_summary,
            "item",
            "세금계산서 매입계",
            months,
        )

    purchase_end_row = max(last_purchase_detail_row, purchase_total_row or 2)
    invoice_total_row = purchase_total_row
    if purchase_total_row is not None and _has_values(result.purchase_summary, "item", "고정"):
        adjustment_row = purchase_total_row + 1
        general_row = adjustment_row + 1
        fixed_row = general_row + 1
        invoice_total_row = fixed_row + 1
        split_rows = (
            (
                adjustment_row,
                "단수차이 조정",
                f"=J{purchase_total_row}",
                f"=ROUNDDOWN(J{purchase_total_row}*0.1,0)",
            ),
            (
                general_row,
                "일반",
                _total_value(result.purchase_summary, "item", "일반매입", "supply_amount"),
                f"=K{adjustment_row}-K{fixed_row}",
            ),
            (
                fixed_row,
                "고정",
                _total_value(result.purchase_summary, "item", "고정", "supply_amount"),
                _total_value(result.purchase_summary, "item", "고정", "tax_amount"),
            ),
            (
                invoice_total_row,
                "계",
                f"=J{general_row}+J{fixed_row}",
                f"=K{general_row}+K{fixed_row}",
            ),
        )
        for row, title, supply_value, tax_value in split_rows:
            for column, value in zip(
                range(purchase_summary_column, purchase_summary_column + 3),
                (title, supply_value, tax_value),
            ):
                cell = sheet.cell(row, column, value)
                _style_ledger_cell(cell, bold=True, total=title == "계")
                cell.alignment = Alignment(horizontal="right", vertical="center")
                if column > purchase_summary_column:
                    cell.number_format = MONEY_FORMAT
        purchase_end_row = max(purchase_end_row, invoice_total_row)

    other_items = [
        (title, key)
        for title, key in (("카드외", "카드외"), ("의제매입세액", "의제매입세액"))
        if _has_values(result.purchase_summary, "item", key)
    ]
    other_total_row: int | None = None
    if other_items:
        other_title_row = purchase_end_row + 2
        other_detail_end_row = other_title_row - 1
        for slot, (title, key) in enumerate(other_items):
            other_detail_end_row = max(
                other_detail_end_row,
                _write_month_table(
                    sheet,
                    other_title_row,
                    DETAIL_START_COLUMNS[slot],
                    title,
                    result.purchase_summary,
                    "item",
                    key,
                    months,
                    marker="②" if slot == 0 else "",
                ),
            )
        other_total_row = _write_compact_month_table(
            sheet,
            other_title_row,
            purchase_summary_column,
            "그 밖의 공제매입세액",
            result.purchase_summary,
            "item",
            "그 밖의 공제매입세액",
            months,
        )
        purchase_end_row = max(other_detail_end_row, other_total_row)

    if _has_values(result.purchase_summary, "item", "매입세액 합계"):
        summary_start_row = purchase_end_row + 1
        overall_row = summary_start_row
        total_references = [
            row
            for row in (invoice_total_row, other_total_row)
            if row is not None
        ]
        overall_values = ["계"]
        for column in (10, 11):
            letter = get_column_letter(column)
            overall_values.append(
                "=" + "+".join(f"{letter}{row}" for row in total_references)
            )
        for column, value in zip(
            range(purchase_summary_column, purchase_summary_column + 3),
            overall_values,
        ):
            cell = sheet.cell(overall_row, column, value)
            _style_ledger_cell(cell, bold=True, total=True)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if column > purchase_summary_column:
                cell.number_format = MONEY_FORMAT

        deduction_rows: list[int] = []
        row = overall_row + 1
        for title in ("불공", "공통"):
            if _has_values(result.purchase_summary, "item", title):
                deduction_rows.append(row)
                values = (
                    title,
                    _total_value(result.purchase_summary, "item", title, "supply_amount"),
                    _total_value(result.purchase_summary, "item", title, "tax_amount"),
                )
                for column, value in zip(
                    range(purchase_summary_column, purchase_summary_column + 3),
                    values,
                ):
                    cell = sheet.cell(row, column, value)
                    _style_ledger_cell(cell, bold=True)
                    cell.alignment = Alignment(horizontal="right", vertical="center")
                    if column > purchase_summary_column:
                        cell.number_format = MONEY_FORMAT
                row += 1

        deduction_total_row = row
        deduction_values = ["차감계"]
        for column in (10, 11):
            letter = get_column_letter(column)
            deduction_values.append(
                f"={letter}{overall_row}"
                + "".join(f"-{letter}{deduction_row}" for deduction_row in deduction_rows)
            )
        for column, value in zip(
            range(purchase_summary_column, purchase_summary_column + 3),
            deduction_values,
        ):
            cell = sheet.cell(row, column, value)
            _style_ledger_cell(cell, bold=True, total=True)
            cell.alignment = Alignment(horizontal="right", vertical="center")
            if column > purchase_summary_column:
                cell.number_format = MONEY_FORMAT
        purchase_end_row = deduction_total_row

        deduction_items = [
            title
            for title in ("불공", "공통")
            if _has_values(result.purchase_summary, "item", title)
        ]
        if deduction_items:
            deduction_title_row = purchase_end_row + 1
            for slot, title in enumerate(deduction_items):
                deduction_total_row = _write_month_table(
                    sheet,
                    deduction_title_row,
                    DETAIL_START_COLUMNS[slot],
                    title,
                    result.purchase_summary,
                    "item",
                    title,
                    months,
                    title_start_column=DETAIL_START_COLUMNS[slot] + 1,
                )
                purchase_end_row = max(purchase_end_row, deduction_total_row)

    sales_title_row = purchase_end_row + 3
    sales_accounts = [
        account
        for account in sorted(
            result.sales_by_account["account_name"].dropna().astype(str).unique()
        )
        if _has_values(result.sales_by_account, "account_name", account)
    ]
    sales_items: list[tuple[str, pd.DataFrame, str, str, bool]] = [
        (account, result.sales_by_account, "account_name", account, False)
        for account in sales_accounts
    ]
    if _has_values(result.sales_summary, "item", "면세 매출"):
        sales_items.append(("면세", result.sales_summary, "item", "면세 매출", True))

    last_sales_detail_row = sales_title_row - 1
    sales_table_totals: list[tuple[str, int, int]] = []
    for item_index in range(0, len(sales_items), len(DETAIL_START_COLUMNS)):
        band = item_index // len(DETAIL_START_COLUMNS)
        title_row = sales_title_row + band * rows_per_band
        for slot, (title, frame, key_column, key, exempt) in enumerate(
            sales_items[item_index : item_index + len(DETAIL_START_COLUMNS)]
        ):
            start_column = DETAIL_START_COLUMNS[slot]
            total_row = _write_month_table(
                sheet,
                title_row,
                start_column,
                title,
                frame,
                key_column,
                key,
                months,
                marker="③" if item_index == 0 and slot == 0 else "",
                exempt=exempt,
            )
            sales_table_totals.append((key, total_row, start_column))
            last_sales_detail_row = max(last_sales_detail_row, total_row)

    taxable_sales_total_ref: str | None = None
    if len(sales_accounts) > 1 and _has_values(
        result.sales_summary,
        "item",
        "과세매출 총계",
    ):
        compact_total_row = _write_compact_month_table(
            sheet,
            sales_title_row,
            9,
            "계",
            result.sales_summary,
            "item",
            "과세매출 총계",
            months,
        )
        taxable_sales_total_ref = f"J{compact_total_row}"
        last_sales_detail_row = max(last_sales_detail_row, compact_total_row)
    elif sales_accounts:
        first_account_total = next(
            item for item in sales_table_totals if item[0] == sales_accounts[0]
        )
        taxable_sales_total_ref = (
            f"{get_column_letter(first_account_total[2] + 2)}{first_account_total[1]}"
        )
    elif _has_values(result.sales_summary, "item", "과세매출 총계"):
        total_row = _write_month_table(
            sheet,
            sales_title_row,
            1,
            "계",
            result.sales_summary,
            "item",
            "과세매출 총계",
            months,
            marker="③",
        )
        taxable_sales_total_ref = f"C{total_row}"
        last_sales_detail_row = total_row

    row = max(purchase_end_row, last_sales_detail_row)
    if taxable_sales_total_ref is not None:
        sales_declaration_row = last_sales_detail_row + 1
        supply_cell = sheet.cell(
            sales_declaration_row,
            3,
            f"={taxable_sales_total_ref}",
        )
        tax_cell = sheet.cell(
            sales_declaration_row,
            4,
            f"=ROUNDDOWN(C{sales_declaration_row}*0.1,0)",
        )
        for cell in (supply_cell, tax_cell):
            _style_ledger_cell(cell, bold=True)
            cell.border = Border()
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.number_format = MONEY_FORMAT
        row = sales_declaration_row

        card_items = [
            (display, key)
            for display, key in (("카드", "카드매출"), ("현영", "현영매출"))
            if _has_values(
                result.sales_summary,
                "item",
                key,
                fields=("total_amount",),
            )
        ]
        if card_items:
            card_outside_row = sales_declaration_row + 1
            card_table_title_row = card_outside_row + 5
            card_total_row = card_table_title_row + len(months) + 1
            gross_refs = [
                f"{get_column_letter(column)}{card_total_row}"
                for column in range(2, 2 + len(card_items))
            ]
            gross_formula = "+".join(gross_refs)
            label = sheet.cell(card_outside_row, 2, "카드외")
            supply = sheet.cell(
                card_outside_row,
                3,
                f"=ROUND(({gross_formula})/1.1,0)",
            )
            tax = sheet.cell(
                card_outside_row,
                4,
                f"={gross_formula}-C{card_outside_row}",
            )
            for cell in (label, supply, tax):
                _style_ledger_cell(cell, bold=True)
                cell.border = Border()
                cell.alignment = Alignment(horizontal="right", vertical="center")
                cell.number_format = MONEY_FORMAT

            row, _ = _write_total_amount_table(
                sheet,
                card_table_title_row,
                card_items,
                result.sales_summary,
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
    widths = (4, 10, 13, 9, 4, 7, 11, 9, 10, 10, 9, 7, 7)
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
