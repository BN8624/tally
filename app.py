# 업체·업종 설정, 매입매출장 처리, 과세기간 선택, 결과 저장을 제공하는 로컬 데스크톱 앱입니다.
from __future__ import annotations

from collections.abc import Callable
from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

from tally import (
    INDUSTRIES,
    CompanySettings,
    InputWorkbookError,
    SettingsStore,
    export_workbook,
    parse_workbook,
    process_transactions,
)
from tally.core import BLOCK_TITLES, cell_total

PROCESSING_ERRORS = (InputWorkbookError, OSError, ValueError, KeyError)
APP_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
SETTINGS_PATH = APP_DIR / "data" / "companies.json"


def _error_text(exc: Exception) -> str:
    if isinstance(exc, KeyError) and exc.args:
        return str(exc.args[0])
    return str(exc)


def _result_matches(
    context: tuple[str, str, dict[str, object]] | None,
    file_path: str,
    company: str,
    store: SettingsStore,
) -> bool:
    if context is None:
        return False
    try:
        current_settings = store.get(company).to_dict()
    except KeyError:
        return False
    return context == (file_path, company, current_settings)


class CompanyDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc) -> None:
        super().__init__(parent)
        self.title("새 업체 설정")
        self.resizable(False, False)
        self.result: CompanySettings | None = None
        self.transient(parent)
        self.grab_set()

        self.name_var = tk.StringVar()
        self.industry_var = tk.StringVar(value="기본")
        self.fixed_var = tk.StringVar()
        self.overrides_var = tk.StringVar()
        self.vehicle_var = tk.StringVar(value="차량유지비, 주유, 수리, 타이어, 세차, 자동차")
        self.vendor_var = tk.StringVar()
        self.personal_var = tk.StringVar(value="개인사용, 가사용, 사적, 가사")
        self.card_deduction_var = tk.BooleanVar(value=True)

        frame = ttk.Frame(self, padding=16)
        frame.grid(sticky="nsew")
        ttk.Label(frame, text="업체명").grid(row=0, column=0, sticky="w", pady=4, padx=(0, 12))
        ttk.Entry(frame, textvariable=self.name_var, width=48).grid(row=0, column=1, sticky="ew", pady=4)
        ttk.Label(frame, text="업종").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(frame, textvariable=self.industry_var, values=INDUSTRIES, state="readonly", width=35).grid(
            row=1, column=1, sticky="ew", pady=4
        )
        fields = [
            ("고정자산 계정코드 · 쉼표 구분", self.fixed_var),
            ("계정 예외 · 예: 899=도급", self.overrides_var),
            ("불공 후보 차량 키워드", self.vehicle_var),
            ("불공 후보 거래처 키워드", self.vendor_var),
            ("불공 후보 사업 무관 키워드", self.personal_var),
        ]
        row = 2
        for label, variable in fields:
            ttk.Label(frame, text=label).grid(row=row, column=0, sticky="w", pady=4, padx=(0, 12))
            ttk.Entry(frame, textvariable=variable, width=48).grid(row=row, column=1, sticky="ew", pady=4)
            row += 1
        ttk.Checkbutton(
            frame,
            text="카드매출 세액공제 대상 · 과세 결제 합계의 13/1000, 납부세액 한도",
            variable=self.card_deduction_var,
        ).grid(row=row, column=1, sticky="w", pady=4)
        row += 1

        button_frame = ttk.Frame(frame)
        button_frame.grid(row=row, column=0, columnspan=2, sticky="e", pady=(12, 0))
        ttk.Button(button_frame, text="취소", command=self.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(button_frame, text="저장", command=self._save).pack(side="right")
        self.bind("<Return>", lambda _: self._save())
        self.wait_visibility()
        self.focus_set()

    @staticmethod
    def _split(value: str) -> list[str]:
        return [item.strip() for item in value.split(",") if item.strip()]

    def _save(self) -> None:
        try:
            overrides: dict[str, str] = {}
            for item in self._split(self.overrides_var.get()):
                if "=" not in item:
                    raise ValueError(f"계정 예외는 코드=열 형식이어야 합니다. 입력={item}")
                code, category = item.split("=", 1)
                overrides[code.strip()] = category.strip()
            self.result = CompanySettings(
                name=self.name_var.get(),
                industry=self.industry_var.get(),
                fixed_asset_codes=set(self._split(self.fixed_var.get())),
                account_overrides=overrides,
                vehicle_keywords=self._split(self.vehicle_var.get()),
                vendor_keywords=self._split(self.vendor_var.get()),
                personal_keywords=self._split(self.personal_var.get()),
                card_deduction_eligible=self.card_deduction_var.get(),
            )
        except ValueError as exc:
            messagebox.showerror("업체 설정 오류", str(exc), parent=self)
            return
        self.destroy()


class TallyApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Tally · 부가세 월별 집계")
        self.geometry("1280x860")
        self.minsize(1000, 640)

        self.store = SettingsStore(SETTINGS_PATH)
        self.source_data: pd.DataFrame | None = None
        self.result = None
        self.result_settings: CompanySettings | None = None
        self.result_context: tuple[str, str, dict[str, object]] | None = None
        self.file_var = tk.StringVar()
        self.company_var = tk.StringVar()
        self.period_var = tk.StringVar()
        self.status_var = tk.StringVar(value="업체와 엑셀 파일을 선택하세요.")

        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("맑은 고딕", 16, "bold"))
        style.configure("Status.TLabel", font=("맑은 고딕", 10, "bold"))

        self._build_ui()
        self._refresh_companies()

    def _build_ui(self) -> None:
        header = ttk.Frame(self, padding=(16, 12))
        header.pack(fill="x")
        ttk.Label(header, text="Tally 부가세 월별 집계", style="Title.TLabel").pack(side="left")
        ttk.Label(header, textvariable=self.status_var, style="Status.TLabel").pack(side="right")

        form = ttk.LabelFrame(self, text="입력", padding=16)
        form.pack(fill="x", padx=16)
        ttk.Label(form, text="업체").grid(row=0, column=0, sticky="w", padx=(0, 12), pady=6)
        self.company_combo = ttk.Combobox(form, textvariable=self.company_var, state="readonly", width=45)
        self.company_combo.grid(row=0, column=1, sticky="ew", pady=6)
        self.company_combo.bind("<<ComboboxSelected>>", lambda _: self._refresh_stale_state())
        ttk.Button(form, text="새 업체 설정 추가", command=self._add_company).grid(row=0, column=2, padx=(10, 0))
        ttk.Label(form, text="매입매출장").grid(row=1, column=0, sticky="w", padx=(0, 12), pady=6)
        ttk.Entry(form, textvariable=self.file_var, state="readonly").grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(form, text="엑셀 선택", command=self._select_file).grid(row=1, column=2, padx=(10, 0))
        ttk.Button(form, text="처리 시작", command=self._process_file).grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 6))
        ttk.Label(form, text="과세기간").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=6)
        self.period_combo = ttk.Combobox(form, textvariable=self.period_var, state="readonly", width=45)
        self.period_combo.grid(row=3, column=1, sticky="w", pady=6)
        self.period_combo.bind("<<ComboboxSelected>>", lambda _: self._change_period())
        form.columnconfigure(1, weight=1)

        controls = ttk.Frame(self, padding=(16, 10, 16, 6))
        controls.pack(fill="x")
        self.result_banner = ttk.Label(controls, text="아직 처리한 결과가 없습니다.", style="Status.TLabel")
        self.result_banner.pack(side="left")
        ttk.Button(controls, text="엑셀 저장", command=self._save_output).pack(side="right")

        text_frame = ttk.Frame(self, padding=(16, 0, 16, 16))
        text_frame.pack(fill="both", expand=True)
        self.result_text = tk.Text(text_frame, wrap="none", font=("Consolas", 10), state="disabled")
        y_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.result_text.yview)
        x_scroll = ttk.Scrollbar(text_frame, orient="horizontal", command=self.result_text.xview)
        self.result_text.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.result_text.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")
        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        guide = (
            "원본 파일은 이 PC에서만 처리되며 서버나 외부 API로 전송되지 않습니다. "
            "불공·의제·공통 안분·경감·예정·가산세 등은 저장한 엑셀의 녹색 칸에 입력하세요."
        )
        ttk.Label(self, text=guide, padding=(16, 0, 16, 12)).pack(anchor="w")

    def _refresh_companies(self) -> None:
        companies = list(self.store.load_all())
        self.company_combo["values"] = companies
        if companies and self.company_var.get() not in companies:
            self.company_var.set(companies[0])

    def _add_company(self) -> None:
        dialog = CompanyDialog(self)
        self.wait_window(dialog)
        if dialog.result is None:
            return
        self.store.save(dialog.result)
        self._refresh_companies()
        self.company_var.set(dialog.result.name)
        self._refresh_stale_state()

    def _select_file(self) -> None:
        path = filedialog.askopenfilename(
            parent=self,
            title="매입매출장 선택",
            filetypes=(("Excel 통합 문서", "*.xlsx *.xlsm"), ("모든 파일", "*.*")),
        )
        if path:
            self.file_var.set(path)
            self._refresh_stale_state()

    def _guarded(self, title: str, action: Callable[[], None]) -> bool:
        try:
            action()
        except PROCESSING_ERRORS as exc:
            messagebox.showerror(title, _error_text(exc), parent=self)
            return False
        return True

    def _process_file(self) -> None:
        if not self.company_var.get():
            messagebox.showwarning("업체 선택", "업체 설정을 먼저 추가하고 선택하세요.", parent=self)
            return
        if not self.file_var.get():
            messagebox.showwarning("파일 선택", "매입매출장 엑셀을 선택하세요.", parent=self)
            return

        def load() -> None:
            self._clear_result()
            file_path = self.file_var.get()
            company = self.company_var.get()
            settings = self.store.get(company)
            source_data = parse_workbook(file_path)
            self.result = process_transactions(source_data, settings)
            self.source_data = source_data
            self.result_settings = settings
            self.result_context = (file_path, company, settings.to_dict())
            self._refresh_periods()
            self._refresh_views()
            self._warn_review()

        if not self._guarded("처리 실패", load):
            self.status_var.set("처리 실패. 이전 결과는 지웠습니다. 오류를 확인한 뒤 다시 처리하세요.")

    def _warn_review(self) -> None:
        review = self.result.review
        if review.empty:
            return
        codes = ", ".join(sorted(set(review["account_code"])))
        messagebox.showerror(
            "구분요망",
            f"{self.result.profile.name} 업종에 맞지 않는 계정 {codes} 매입이 {len(review):,}건 있습니다. "
            "계정과목 구분이 잘못되었을 수 있습니다.\n"
            "저장한 엑셀의 구분요망 시트에서 확인하세요. 의도한 계정이면 업체 설정의 계정 예외로 지정하세요.",
            parent=self,
        )

    def _clear_result(self) -> None:
        self.result = None
        self.source_data = None
        self.result_settings = None
        self.result_context = None
        self.period_combo["values"] = ()
        self.period_var.set("")
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.configure(state="disabled")
        self.result_banner.configure(text="아직 처리한 결과가 없습니다.", foreground="")

    def _refresh_periods(self) -> None:
        self.period_combo["values"] = [period.label for period in self.result.periods]
        self.period_var.set(self.result.period.label if self.result.period else "")

    def _change_period(self) -> None:
        if not self._require_current_result("과세기간을 바꾸"):
            return
        period = next((item for item in self.result.periods if item.label == self.period_var.get()), None)
        if period is None:
            return

        def recalculate() -> None:
            self.result = process_transactions(self.source_data, self.result_settings, period)
            self._refresh_views()
            self._warn_review()

        self._guarded("과세기간 변경 실패", recalculate)

    def _result_is_current(self) -> bool:
        return _result_matches(self.result_context, self.file_var.get(), self.company_var.get(), self.store)

    def _refresh_stale_state(self) -> None:
        if self.result is not None and not self._result_is_current():
            self.result_banner.configure(text="재처리 필요", foreground="#9C0006")
            self.status_var.set("파일·업체·업체 설정이 바뀌었습니다. 처리 시작을 다시 실행하세요.")

    def _require_current_result(self, action: str) -> bool:
        if self.result is None:
            messagebox.showwarning("처리 결과", "먼저 엑셀 파일을 처리하세요.", parent=self)
            return False
        if not self._result_is_current():
            messagebox.showwarning(
                "재처리 필요",
                f"처리한 뒤 파일·업체·업체 설정이 바뀌었습니다. 처리 시작을 다시 실행한 뒤 {action}세요.",
                parent=self,
            )
            return False
        return True

    def _refresh_views(self) -> None:
        result = self.result
        state = "검산 완료" if result.validation_passed else "검산 실패"
        self.result_banner.configure(text=state, foreground="#38761D" if result.validation_passed else "#9C0006")
        self.status_var.set(
            f"{result.period.label if result.period else '과세기간 없음'} · 업종 {result.profile.name} · "
            f"기간 안 {int(result.transactions['in_period'].sum()):,}건 · 불공 후보 {len(result.candidates):,}건"
        )
        lines = [f"[{state}] {result.settings.name} · {result.period.label if result.period else ''}", ""]
        lines.append("블록별 기간 합계 (매수 · 공급가액 · 세액, 결제 채널은 합계금액)")
        for block, columns in result.layout.items():
            lines.append(f"  {BLOCK_TITLES[block]}")
            for column in columns:
                labels = column.sum_of or (column.label,)
                if column.kind == "amount":
                    amount = cell_total(result.cells, block, labels, "total")
                    lines.append(f"    {column.label:<12} {amount:>16,}")
                else:
                    count = cell_total(result.cells, block, labels, "count")
                    supply = cell_total(result.cells, block, labels, "supply")
                    tax = cell_total(result.cells, block, labels, "tax")
                    lines.append(f"    {column.label:<12} {count:>6,} {supply:>16,} {tax:>14,}")
        lines += ["", "요약 (입력 칸은 0으로 계산, 최종 금액은 저장한 엑셀에서 확인)"]
        for row in result.summary:
            supply = "" if row.supply is None else f"{row.supply:,}"
            tax = "" if row.tax is None else f"{row.tax:,}"
            lines.append(f"  {row.label:<14} {supply:>16} {tax:>14}")
        lines += ["", "검산"]
        for _, check in result.validation.iterrows():
            lines.append(f"  [{check['status']}] {check['check']}: {check['actual']} {check['detail']}")
        self.result_text.configure(state="normal")
        self.result_text.delete("1.0", "end")
        self.result_text.insert("1.0", "\n".join(lines))
        self.result_text.configure(state="disabled")

    def _save_output(self) -> None:
        if not self._require_current_result("저장하"):
            return
        if not self.result.validation_passed and not messagebox.askyesno(
            "검산 실패",
            "미집계 유형·미분류 등 검산 실패 항목이 있습니다. 실패 상태를 포함한 검토용 엑셀을 저장할까요?",
            parent=self,
        ):
            return
        period = self.result.period.label.replace(" ", "_") if self.result.period else "기간없음"
        path = filedialog.asksaveasfilename(
            parent=self,
            title="결과 엑셀 저장",
            defaultextension=".xlsx",
            initialfile=f"{self.company_var.get()}_{period}_부가세_집계.xlsx",
            filetypes=(("Excel 통합 문서", "*.xlsx"),),
        )
        if not path:
            return
        try:
            export_workbook(self.result, path)
        except PermissionError:
            messagebox.showerror(
                "저장 실패",
                "결과 파일에 쓸 수 없습니다. 같은 이름의 엑셀이 열려 있으면 닫고 다시 저장하세요.\n" f"{Path(path)}",
                parent=self,
            )
            return
        except PROCESSING_ERRORS as exc:
            messagebox.showerror("저장 실패", _error_text(exc), parent=self)
            return
        messagebox.showinfo("저장 완료", f"결과 엑셀을 저장했습니다.\n{Path(path)}", parent=self)


def main() -> None:
    TallyApp().mainloop()


if __name__ == "__main__":
    main()
