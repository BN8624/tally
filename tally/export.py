# 처리 결과를 업종 집계양식 틀의 집계표와 전자·종이·불공후보·거래상세·미분류·검산 시트로 저장합니다.
from __future__ import annotations

from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Callable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .core import (
    BLOCK_TITLES,
    CARD,
    EXEMPT,
    NONDEDUCTIBLE,
    PAYMENT,
    PURCHASE,
    SALES,
    Column,
    ProcessingResult,
)
from .profiles import EXEMPT_PAYMENT, INPUT_CARD_EXEMPT, INPUT_COMMON, INPUT_DEEMED

FIELDS = {"tax": ("count", "supply", "tax"), "exempt": ("count", "supply"), "amount": ("total",)}
MONEY_FORMAT = '#,##0;-#,##0;"-"'
NAVY = "1F4E78"
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEADER_FILL = PatternFill("solid", fgColor="EDEDED")
INPUT_FILL = PatternFill("solid", fgColor="C6EFCE")
SECTION_FILL = PatternFill("solid", fgColor="DDEBF7")
FONT = Font(name="맑은 고딕", size=10)
BOLD = Font(name="맑은 고딕", size=10, bold=True)
TABLE_MONEY_FORMAT = "#,##0;[Red]-#,##0"
TABLE_DATE_FORMAT = "yyyy-mm-dd"
TABLE_WRAP_FIELDS = {"item", "candidate_reason", "detail"}
TABLE_MONEY_FIELDS = {"supply_amount", "tax_amount", "total_amount", "expected", "actual", "difference"}
TABLE_WIDE_FIELDS = {"vendor", "item", "candidate_reason", "detail"}
TABLE_ALERT_VALUES = {"실패", "미분류"}
TABLE_ALERT_FILL = PatternFill("solid", fgColor="F4CCCC")


def _excel_value(value: object) -> object:
    if value is None or (not isinstance(value, (str, bool)) and pd.isna(value)):
        return ""
    if isinstance(value, Decimal):
        return int(value) if value == value.to_integral_value() else float(value)
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    return value


def _keep_text(cell) -> None:
    if isinstance(cell.value, str) and cell.value.startswith("="):
        cell.data_type = "s"


def _style(cell, *, bold: bool = False, fill: PatternFill | None = None, number: bool = False) -> None:
    cell.font = BOLD if bold else FONT
    cell.border = BOX
    if fill is not None:
        cell.fill = fill
    if number:
        cell.number_format = MONEY_FORMAT
        cell.alignment = Alignment(horizontal="right", vertical="center")
    else:
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _month_label(month: str) -> str:
    return f"{int(month[5:])}월"


class _Sheet:
    """셀을 쓸 때마다 위치를 기록해 서로 덮어쓰지 않았는지 확인할 수 있게 합니다."""

    def __init__(self, sheet) -> None:
        self.sheet = sheet
        self.money_columns: set[int] = set()
        self.count_columns: set[int] = set()
        self.last_row = 1
        self.last_column = 1

    def write(self, row: int, column: int, value: object, *, count: bool = False, text: bool = False, **style) -> object:
        cell = self.sheet.cell(row, column, value)
        if text:
            _keep_text(cell)
        _style(cell, **style)
        if style.get("number"):
            (self.count_columns if count else self.money_columns).add(column)
        self.last_row = max(self.last_row, row)
        self.last_column = max(self.last_column, column)
        return cell

    def set_widths(self, label_column: int | None = None) -> None:
        self.sheet.column_dimensions["A"].width = 7
        for column in range(2, self.last_column + 1):
            letter = get_column_letter(column)
            if column == label_column:
                width = 24
            elif column in self.money_columns:
                width = 13
            elif column in self.count_columns:
                width = 6
            else:
                width = 3
            self.sheet.column_dimensions[letter].width = width


def _render_block(
    out: _Sheet,
    top: int,
    left: int,
    title: str | None,
    columns: list[Column],
    value_of: Callable[[str, str, str], object],
    months: list[str],
) -> tuple[dict[tuple[str, str], str], int]:
    """블록 하나를 그리고 (열, 값 종류)별 계 셀 주소와 오른쪽 끝 열을 돌려줍니다."""
    positions: dict[str, int] = {}
    column = left
    for spec in columns:
        positions[spec.label] = column
        column += len(FIELDS[spec.kind])
    right = column - 1
    if title is not None:
        out.write(top, 1, title, bold=True, fill=HEADER_FILL)
    for spec in columns:
        start = positions[spec.label]
        width = len(FIELDS[spec.kind])
        out.write(top, start, spec.label, bold=True, fill=HEADER_FILL, text=True)
        for offset in range(1, width):
            out.write(top, start + offset, None, fill=HEADER_FILL)
        if width > 1:
            out.sheet.merge_cells(start_row=top, start_column=start, end_row=top, end_column=start + width - 1)

    first, last = top + 1, top + len(months)
    for index, month in enumerate(months):
        row = first + index
        if title is not None:
            out.write(row, 1, _month_label(month))
        for spec in columns:
            for offset, field_name in enumerate(FIELDS[spec.kind]):
                target = positions[spec.label] + offset
                if spec.sum_of:
                    parts = [
                        f"{get_column_letter(positions[label] + FIELDS[spec.kind].index(field_name))}{row}"
                        for label in spec.sum_of
                        if label in positions
                    ]
                    value = "=" + "+".join(parts) if parts else 0
                else:
                    value = value_of(spec.label, month, field_name)
                out.write(row, target, value, number=True, count=field_name == "count")

    total_row = last + 1
    if title is not None:
        out.write(total_row, 1, "계", bold=True, fill=HEADER_FILL)
    refs: dict[tuple[str, str], str] = {}
    for spec in columns:
        for offset, field_name in enumerate(FIELDS[spec.kind]):
            target = positions[spec.label] + offset
            letter = get_column_letter(target)
            formula = f"=SUM({letter}{first}:{letter}{last})" if months else 0
            out.write(total_row, target, formula, bold=True, fill=HEADER_FILL, number=True, count=field_name == "count")
            refs[(spec.label, field_name)] = f"{letter}{total_row}"
    return refs, right


def _cell_lookup(cells: pd.DataFrame, block: str) -> Callable[[str, str, str], object]:
    rows = cells[cells["block"] == block]
    table = {(record.column, record.month): record for record in rows.itertuples(index=False)}

    def value_of(label: str, month: str, field_name: str) -> object:
        record = table.get((label, month))
        return 0 if record is None else _excel_value(getattr(record, field_name))

    return value_of


def _summary_rows(result: ProcessingResult) -> list[tuple[str, str, str]]:
    inputs = result.profile.inputs
    rows: list[tuple[str, str, str]] = [("section", "매입", "")]
    rows += [("general", "일반", "formula"), ("fixed", "고정", "formula"), ("invoice", "세금계산서 계", "formula")]
    if INPUT_DEEMED in inputs:
        rows.append(("deemed", "의제", "input2"))
    rows += [("card", "카드", "formula"), ("nondeductible", "불공", "formula"), ("extra_nondeductible", "추가 불공", "input2")]
    if INPUT_COMMON in inputs:
        rows.append(("common", INPUT_COMMON, "input2"))
    rows += [("purchase", "매입계", "formula"), ("section", "매출", "")]
    rows += [("sales_invoice", "세금계산서", "formula"), ("zero_rated", "영세", "formula")]
    if INPUT_CARD_EXEMPT in inputs:
        rows.append(("card_exempt", f"{INPUT_CARD_EXEMPT}(합계금액)", "input1"))
    rows += [
        ("card_out", "카드외", "formula"),
        ("other", "기타(건별)", "formula"),
        ("taxable", "과세계", "formula"),
        ("exempt_invoice", "면세 계산서", "formula"),
        ("exempt_payment", "면세 결제", "formula"),
        ("exempt", "면세계", "formula"),
        ("total_sales", "총매출", "formula"),
        ("section", "납부", ""),
        ("payable", "납부(환급)세액", "formula"),
        ("card_deduction", "카드공제" if result.settings.card_deduction_eligible else "카드공제(비대상)", "formula"),
    ]
    if result.profile.taxi_relief:
        rows.append(("taxi_relief", "택시 경감세액", "prefill_tax"))
    rows += [
        ("relief", "경감·공제세액", "input_tax"),
        ("prior_refund", "예정 미환급세액", "input_tax"),
        ("prior_notice", "예정 고지세액", "input_tax"),
        ("buyer_paid", "매입자 납부특례 기납부세액", "input_tax"),
        ("penalty", "가산세", "input_tax"),
        ("final", "차감 납부할 세액", "formula"),
    ]
    if result.profile.taxi_relief:
        rows += [
            ("section", "택시 경감 배분", ""),
            ("taxi_drivers", "운수종사자 지급(90%)", "formula"),
            ("taxi_reduction", "감차 보상(5%)", "formula"),
            ("taxi_welfare", "운수종사자 복지(4%)", "formula"),
        ]
    return rows


def _render_summary(
    out: _Sheet,
    top: int,
    left: int,
    result: ProcessingResult,
    refs: dict[str, dict[tuple[str, str], str]],
) -> None:
    label_column, supply_column, tax_column = left, left + 1, left + 2
    supply_letter, tax_letter = get_column_letter(supply_column), get_column_letter(tax_column)
    out.write(top, label_column, "요약", bold=True, fill=HEADER_FILL)
    out.write(top, supply_column, "공급가액", bold=True, fill=HEADER_FILL)
    out.write(top, tax_column, "세액", bold=True, fill=HEADER_FILL)
    specs = _summary_rows(result)
    row_of: dict[str, int] = {}
    for offset, (key, _, _) in enumerate(specs, start=1):
        if key != "section":
            row_of[key] = top + offset

    def ref(block: str, label: str, field_name: str) -> str:
        return refs.get(block, {}).get((label, field_name), "0")

    def s(key: str) -> str:
        return f"{supply_letter}{row_of[key]}" if key in row_of else "0"

    def t(key: str) -> str:
        return f"{tax_letter}{row_of[key]}" if key in row_of else "0"

    payment_taxable = ref(PAYMENT, "과세 계", "total")
    card_base = f"({payment_taxable}-{s('card_exempt')})"
    exempt_payment_refs = [ref(PAYMENT, label, "total") for label in EXEMPT_PAYMENT if (label, "total") in refs.get(PAYMENT, {})]
    formulas: dict[str, tuple[object, object]] = {
        "general": (f"={s('invoice')}-{s('fixed')}", f"={t('invoice')}-{t('fixed')}"),
        "fixed": (f"={ref(PURCHASE, '고정', 'supply')}", f"={ref(PURCHASE, '고정', 'tax')}"),
        "invoice": (f"={ref(PURCHASE, '계', 'supply')}", f"=ROUNDDOWN({s('invoice')}*0.1,0)"),
        "card": (f"={ref(CARD, '공제 계', 'supply')}", f"={ref(CARD, '공제 계', 'tax')}"),
        "nondeductible": (f"={ref(NONDEDUCTIBLE, '불공', 'supply')}", f"={ref(NONDEDUCTIBLE, '불공', 'tax')}"),
        "purchase": (
            f"={s('invoice')}+{s('deemed')}+{s('card')}-{s('nondeductible')}-{s('extra_nondeductible')}-{s('common')}",
            f"={t('invoice')}+{t('deemed')}+{t('card')}-{t('nondeductible')}-{t('extra_nondeductible')}-{t('common')}",
        ),
        "sales_invoice": (f"={ref(SALES, '계', 'supply')}", f"=ROUNDDOWN({s('sales_invoice')}*0.1,0)"),
        "zero_rated": (f"={ref(SALES, '영세', 'supply')}", 0),
        "card_out": (f"=ROUNDUP({card_base}/1.1,0)", f"={card_base}-{s('card_out')}"),
        "other": (f"=ROUNDUP({ref(PAYMENT, '현금', 'total')}/1.1,0)", f"={ref(PAYMENT, '현금', 'total')}-{s('other')}"),
        "taxable": (
            f"={s('sales_invoice')}+{s('zero_rated')}+{s('card_out')}+{s('other')}",
            f"={t('sales_invoice')}+{t('zero_rated')}+{t('card_out')}+{t('other')}",
        ),
        "exempt_invoice": (f"={ref(SALES, '면세', 'supply')}", None),
        "exempt_payment": ("=" + "+".join([*exempt_payment_refs, s("card_exempt")]), None),
        "exempt": (f"={s('exempt_invoice')}+{s('exempt_payment')}", None),
        "total_sales": (f"={s('taxable')}+{s('exempt')}", None),
        "payable": (None, f"={t('taxable')}-{t('purchase')}"),
        "card_deduction": (
            None,
            f"=MAX(0,MIN(ROUNDDOWN({card_base}*13/1000,0),{t('payable')}))"
            if result.settings.card_deduction_eligible
            else 0,
        ),
        "taxi_relief": (None, f"=MAX(0,ROUNDDOWN({t('payable')}*0.99,0))"),
        "final": (
            None,
            f"={t('payable')}"
            + "".join(
                f"-{t(key)}"
                for key in ("card_deduction", "taxi_relief", "relief", "prior_refund", "prior_notice", "buyer_paid")
                if key in row_of
            )
            + f"+{t('penalty')}",
        ),
        "taxi_drivers": (None, f"=MAX(0,ROUNDUP({t('payable')}*0.9,0))"),
        "taxi_reduction": (None, f"=MAX(0,ROUNDDOWN({t('payable')}*0.05,0))"),
        "taxi_welfare": (None, f"=MAX(0,ROUNDDOWN({t('payable')}*0.04,0))"),
    }
    for offset, (key, label, kind) in enumerate(specs, start=1):
        row = top + offset
        if key == "section":
            for column in (label_column, supply_column, tax_column):
                out.write(row, column, label if column == label_column else None, bold=True, fill=SECTION_FILL)
            continue
        out.write(row, label_column, label, bold=key in {"purchase", "taxable", "final", "payable"})
        if kind in {"formula", "prefill_tax"}:
            supply_value, tax_value = formulas[key]
            out.write(row, supply_column, supply_value, number=True)
            out.write(row, tax_column, tax_value, number=True, fill=INPUT_FILL if kind == "prefill_tax" else None)
        else:
            supply_fill = INPUT_FILL if kind in {"input1", "input2"} else None
            tax_fill = INPUT_FILL if kind in {"input2", "input_tax"} else None
            out.write(row, supply_column, None, number=True, fill=supply_fill)
            out.write(row, tax_column, None, number=True, fill=tax_fill)


def _build_summary_sheet(workbook: Workbook, result: ProcessingResult) -> None:
    sheet = workbook.active
    sheet.title = "집계표"
    out = _Sheet(sheet)
    months = result.period.months if result.period else []
    out.write(1, 1, "상호", bold=True, fill=HEADER_FILL)
    out.write(1, 2, result.settings.name, bold=True, text=True)
    sheet.merge_cells(start_row=1, start_column=2, end_row=1, end_column=4)
    out.write(1, 5, result.period.label if result.period else "과세기간 없음", bold=True)
    sheet.merge_cells(start_row=1, start_column=5, end_row=1, end_column=7)
    out.write(1, 8, f"업종 {result.profile.name}", bold=True)
    sheet.merge_cells(start_row=1, start_column=8, end_row=1, end_column=10)
    out.write(1, 11, "부가세 집계표", bold=True)
    sheet.merge_cells(start_row=1, start_column=11, end_row=1, end_column=13)

    refs: dict[str, dict[tuple[str, str], str]] = {}
    row = 3
    right_edge = 13
    for block in (PURCHASE, EXEMPT, CARD, SALES, PAYMENT):
        if block not in result.layout:
            continue
        refs[block], right = _render_block(
            out, row, 2, BLOCK_TITLES[block], result.layout[block], _cell_lookup(result.cells, block), months
        )
        if block == EXEMPT:
            refs[NONDEDUCTIBLE], right = _render_block(
                out, row, right + 2, None, result.layout[NONDEDUCTIBLE], _cell_lookup(result.cells, NONDEDUCTIBLE), months
            )
        right_edge = max(right_edge, right)
        row += len(months) + 3

    summary_left = right_edge + 2
    _render_summary(out, 3, summary_left, result, refs)

    out.set_widths(label_column=summary_left)
    sheet.sheet_view.showGridLines = False
    sheet.freeze_panes = "B3"
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    for side in ("left", "right", "top", "bottom"):
        setattr(sheet.page_margins, side, 0.3)
    sheet.print_area = f"A1:{get_column_letter(out.last_column)}{out.last_row}"


def _electronic_cells(result: ProcessingResult) -> tuple[pd.DataFrame, dict[str, list[Column]]]:
    data = result.transactions
    invoices = data[
        data["in_period"]
        & (
            data["block"].eq(PURCHASE)
            | (data["block"].eq(EXEMPT) & data["column"].ne("영세"))
            | data["block"].eq(SALES)
        )
    ].copy()
    invoices["column"] = invoices["column"] + " " + invoices["electronic"]
    grouped = (
        invoices.groupby(["block", "column", "month"])
        .agg(count=("row_id", "size"), supply=("supply_amount", "sum"), tax=("tax_amount", "sum"), total=("total_amount", "sum"))
        .reset_index()
        if not invoices.empty
        else pd.DataFrame(columns=["block", "column", "month", "count", "supply", "tax", "total"])
    )
    layout: dict[str, list[Column]] = {}
    for block in (PURCHASE, EXEMPT, SALES):
        columns: list[Column] = []
        for spec in result.layout.get(block, []):
            if spec.sum_of or spec.label == "영세" and block == EXEMPT:
                continue
            columns += [Column(f"{spec.label} 전자", spec.kind), Column(f"{spec.label} 종이", spec.kind)]
        layout[block] = columns
    return grouped, layout


def _build_electronic_sheet(workbook: Workbook, result: ProcessingResult) -> None:
    sheet = workbook.create_sheet("전자·종이")
    out = _Sheet(sheet)
    out.write(1, 1, "전자·종이 구분", bold=True, fill=HEADER_FILL)
    if not result.transactions["electronic"].astype(str).ne("").any():
        out.write(3, 1, "원장에 전자세금 열이 없어 전자·종이를 구분하지 않았습니다.")
        sheet.column_dimensions["A"].width = 60
        return
    cells, layout = _electronic_cells(result)
    months = result.period.months if result.period else []
    titles = {PURCHASE: "매입 세금계산서", EXEMPT: "면세 계산서 매입", SALES: "매출 세금계산서·계산서"}
    row = 3
    for block in (PURCHASE, EXEMPT, SALES):
        _render_block(out, row, 2, titles[block], layout[block], _cell_lookup(cells, block), months)
        row += len(months) + 3
    out.set_widths()
    sheet.sheet_view.showGridLines = False
    sheet.page_setup.orientation = "landscape"
    sheet.page_setup.fitToWidth = 1
    sheet.page_setup.fitToHeight = 0
    sheet.sheet_properties.pageSetUpPr.fitToPage = True
    sheet.print_area = f"A1:{get_column_letter(out.last_column)}{out.last_row}"


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
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = sheet.cell(1, 1, title)
    title_cell.font = Font(name="맑은 고딕", size=16, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill("solid", fgColor=NAVY)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    sheet.row_dimensions[1].height = 30
    for column, (_, label) in enumerate(headers, start=1):
        cell = sheet.cell(3, column, label)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.font = Font(name="맑은 고딕", bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap = Alignment(vertical="top", wrap_text=True)
    plain = Alignment(vertical="top", wrap_text=False)
    border = Border(bottom=THIN)
    for row_number, (_, record) in enumerate(frame.iterrows(), start=4):
        for column, (field_name, _) in enumerate(headers, start=1):
            value = _excel_value(record.get(field_name, ""))
            cell = sheet.cell(row_number, column, value)
            _keep_text(cell)
            cell.border = border
            cell.alignment = wrap if field_name in TABLE_WRAP_FIELDS else plain
            if field_name in TABLE_MONEY_FIELDS:
                cell.number_format = TABLE_MONEY_FORMAT
            elif field_name == "date" and isinstance(value, (date, datetime)):
                cell.number_format = TABLE_DATE_FORMAT
        if alert_column and str(record.get(alert_column, "")) in TABLE_ALERT_VALUES:
            for cell in sheet[row_number]:
                cell.fill = TABLE_ALERT_FILL
    sheet.freeze_panes = "A4"
    sheet.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{max(3, 3 + len(frame))}"
    sheet.sheet_view.showGridLines = False
    for index, (field_name, label) in enumerate(headers, start=1):
        width = max(12, min(42, len(label) * 2 + 4))
        if field_name in TABLE_WIDE_FIELDS:
            width = 28 if field_name in {"vendor", "item"} else 36
        sheet.column_dimensions[get_column_letter(index)].width = width


DETAIL_HEADERS = [
    ("row_id", "원본 행 ID"),
    ("sheet", "원본 시트"),
    ("source_row", "원본 행"),
    ("date", "전표일자"),
    ("division", "매입·매출"),
    ("month", "월"),
    ("period_state", "과세기간"),
    ("vendor", "거래처"),
    ("item", "품명"),
    ("card_company", "카드사"),
    ("card_number", "카드번호"),
    ("supply_amount", "공급가액"),
    ("tax_amount", "세액"),
    ("total_amount", "합계금액"),
    ("account_code", "계정코드"),
    ("account_name", "계정과목"),
    ("original_type", "원본 유형"),
    ("electronic", "전자"),
    ("block_title", "집계 블록"),
    ("column", "집계 열"),
    ("candidate_reason", "불공 후보 근거"),
]


def export_workbook(result: ProcessingResult, output_path: str | Path) -> Path:
    """처리 결과를 정본 CANON-10의 여섯 시트로 저장합니다."""
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    workbook = Workbook()
    _build_summary_sheet(workbook, result)
    _build_electronic_sheet(workbook, result)

    details = result.transactions.assign(
        period_state=result.transactions["in_period"].map({True: "포함", False: "제외"}),
        block_title=result.transactions["block"].map(lambda block: BLOCK_TITLES.get(block, "")),
    )
    candidate_headers = [
        ("row_id", "원본 행 ID"),
        ("date", "전표일자"),
        ("vendor", "거래처"),
        ("item", "품명"),
        ("account_code", "계정코드"),
        ("account_name", "계정과목"),
        ("original_type", "원본 유형"),
        ("supply_amount", "공급가액"),
        ("tax_amount", "세액"),
        ("candidate_reason", "후보 근거"),
    ]
    _write_table_sheet(workbook, "불공후보", "불공 후보 · 판정 없이 참고용", candidate_headers, result.candidates)
    _write_table_sheet(workbook, "거래상세", "원본 상세 거래 및 집계 위치", DETAIL_HEADERS, details)
    _write_table_sheet(
        workbook,
        "미분류",
        "계정 미분류 거래",
        DETAIL_HEADERS,
        details[details["column"].eq("미분류")],
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
