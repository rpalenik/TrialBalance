#!/usr/bin/env python3
"""Convert zostava.xlsx into Trial Balance template (new output file)."""

from __future__ import annotations

import argparse
import datetime as dt
import shutil
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, Iterable, Tuple

import pandas as pd
import yaml
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import column_index_from_string


@dataclass(frozen=True)
class Config:
    source: Dict[str, Any]
    template: Dict[str, Any]
    output: Dict[str, Any]
    report: Dict[str, Any]
    rules: Dict[str, Any]


@dataclass(frozen=True)
class ReportPeriod:
    year: int
    month: int


@dataclass(frozen=True)
class MissingAccount:
    account: str
    debit: float
    credit: float
    diff: float


@dataclass(frozen=True)
class RunSummary:
    selected_year: int
    selected_month: int
    rows_before_filter: int
    rows_after_filter: int
    debit_field: str
    credit_field: str
    source_debit: float
    source_credit: float
    source_diff: float
    template_debit: float
    template_credit: float
    template_diff: float
    tolerance: float
    missing_accounts: list[MissingAccount]


@dataclass(frozen=True)
class RunResult:
    output_path: Path
    report_path: Path | None
    period: ReportPeriod
    summary: RunSummary
    summary_path: Path


@dataclass(frozen=True)
class SourceReadResult:
    dataframe: pd.DataFrame
    period: ReportPeriod
    rows_before_filter: int
    rows_after_filter: int


def load_config(path: Path) -> Config:
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    return Config(
        source=data["source"],
        template=data["template"],
        output=data["output"],
        report=data.get("report", {"enabled": False}),
        rules=data["rules"],
    )


def col_letter_to_index(letter: str) -> int:
    return column_index_from_string(letter) - 1


def normalize_account(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return str(value).strip()
    if isinstance(value, int):
        return str(value)
    text = str(value).strip()
    return text if text else None


def normalize_account_part(value: Any, width: int | None) -> str | None:
    account = normalize_account(value)
    if account is None:
        return None
    if width and account.isdigit():
        return account.zfill(width)
    return account

def map_account(account: str | None) -> str | None:
    if not account:
        return None
    # If full account ends with 001, map to corresponding 000 account.
    if account.endswith("001"):
        return f"{account[:-3]}000"
    return account


def parse_number(value: Any, fmt: Dict[str, Any]) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)

    text = str(value).strip()
    if not text:
        return 0.0

    negative = False
    if fmt.get("negative_parens") and text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    if fmt.get("thousands_sep_space"):
        text = text.replace(" ", "").replace("\u00a0", "")

    if fmt.get("decimal_comma"):
        text = text.replace(",", ".")

    try:
        num = float(text)
    except ValueError:
        return 0.0

    return -num if negative else num


def _find_header_col(headers: pd.Series, name: str, input_path: Path) -> int:
    target = name.upper()
    for idx, value in headers.items():
        if str(value).strip().upper() == target:
            return int(idx)
    raise ValueError(
        f"Missing required '{name}' column in source file '{input_path}'."
    )


def select_latest_period_rows(
    df: pd.DataFrame, header_row: int, data_start_row: int, input_path: Path
) -> tuple[pd.DataFrame, ReportPeriod, int, int]:
    headers = df.iloc[header_row - 1]
    month_col = _find_header_col(headers, "MESIC", input_path)
    year_col = _find_header_col(headers, "ROK", input_path)

    data = df.iloc[data_start_row - 1 :].copy()
    rows_before = int(len(data))

    month_values = pd.to_numeric(data.iloc[:, month_col], errors="coerce")
    year_values = pd.to_numeric(data.iloc[:, year_col], errors="coerce")

    valid_mask = month_values.notna() & year_values.notna()
    data = data.loc[valid_mask].copy()
    month_values = month_values.loc[valid_mask]
    year_values = year_values.loc[valid_mask]

    if data.empty:
        raise ValueError(
            f"No rows with valid 'MESIC' and 'ROK' values in '{input_path}'."
        )

    max_month = int(month_values.max())
    if max_month < 1 or max_month > 12:
        raise ValueError(
            f"Invalid latest month '{max_month}' in 'MESIC' in '{input_path}'."
        )

    latest_mask = month_values.astype(int) == max_month
    filtered = data.loc[latest_mask].copy()
    rows_after = int(len(filtered))
    if rows_after == 0:
        raise ValueError(
            f"No rows found for latest month '{max_month}' in '{input_path}'."
        )

    years = sorted({int(v) for v in year_values.loc[latest_mask].astype(int)})
    if len(years) != 1:
        raise ValueError(
            f"Expected exactly one year for latest month '{max_month}', "
            f"found {years} in '{input_path}'."
        )

    year = years[0]
    if year < 1900 or year > 3000:
        raise ValueError(
            f"Invalid processing year '{year}' in 'ROK' in '{input_path}'."
        )

    return filtered, ReportPeriod(year=year, month=max_month), rows_before, rows_after


def extract_report_period(cfg: Config, input_path: Path) -> ReportPeriod:
    source_cfg = cfg.source
    sheet = source_cfg.get("sheet", 0)
    header_row = int(source_cfg.get("header_row", 1))
    data_start_row = int(source_cfg.get("data_start_row", header_row + 1))

    df = pd.read_excel(
        input_path,
        sheet_name=sheet,
        header=None,
        engine="calamine",
    )
    _, period, _, _ = select_latest_period_rows(
        df, header_row=header_row, data_start_row=data_start_row, input_path=input_path
    )
    return period


def read_source(cfg: Config, input_path: Path) -> SourceReadResult:
    source_cfg = cfg.source
    sheet = source_cfg.get("sheet", 0)
    header_row = int(source_cfg.get("header_row", 1))
    data_start_row = int(source_cfg.get("data_start_row", header_row + 1))

    df = pd.read_excel(
        input_path,
        sheet_name=sheet,
        header=None,
        engine="calamine",
    )

    col_map = {
        key: col_letter_to_index(letter)
        for key, letter in source_cfg["columns"].items()
    }

    data, period, rows_before, rows_after = select_latest_period_rows(
        df, header_row=header_row, data_start_row=data_start_row, input_path=input_path
    )

    missing = [
        f"{key}({letter})"
        for key, letter in source_cfg["columns"].items()
        if col_letter_to_index(letter) not in data.columns
    ]
    if missing:
        raise ValueError(
            f"Configured source columns not found in '{input_path}': {', '.join(missing)}."
        )

    account_parts = source_cfg.get("account_parts")
    if account_parts:
        syn_idx = col_letter_to_index(account_parts["synthetic_col"])
        ana_idx = col_letter_to_index(account_parts["analytic_col"])
        syn_width = account_parts.get("synthetic_width")
        ana_width = account_parts.get("analytic_width")

        data = data.rename(columns={syn_idx: "_synthetic", ana_idx: "_analytic"})
        data["_synthetic"] = data["_synthetic"].apply(
            lambda v: normalize_account_part(v, syn_width)
        )
        data["_analytic"] = data["_analytic"].apply(
            lambda v: normalize_account_part(v, ana_width)
        )
        data["account"] = data.apply(
            lambda r: (
                f"{r['_synthetic']}{r['_analytic']}"
                if r["_synthetic"] and r["_analytic"]
                else r["_synthetic"]
            ),
            axis=1,
        )
        data["account"] = data["account"].apply(map_account)
    else:
        account_idx = col_map["account"]
        data = data.rename(columns={account_idx: "account"})
        data["account"] = data["account"].apply(normalize_account).apply(map_account)

    rename_map = {col_map[key]: key for key in col_map.keys()}
    data = data.rename(columns=rename_map)
    ordered_cols = ["account"] + list(col_map.keys())
    data = data[ordered_cols]
    data = data[data["account"].notna()]

    fmt = source_cfg.get("number_format", {})
    for col in list(col_map.keys()):
        data[col] = data[col].apply(lambda v: parse_number(v, fmt))

    if cfg.rules.get("duplicate_accounts") == "sum":
        data = data.groupby("account", as_index=False).sum(numeric_only=True)

    return SourceReadResult(
        dataframe=data,
        period=period,
        rows_before_filter=rows_before,
        rows_after_filter=rows_after,
    )


def read_template_accounts(cfg: Config, template_path: Path) -> Tuple[Any, Dict[str, int]]:
    template_cfg = cfg.template
    wb = load_workbook(template_path)
    ws = wb.worksheets[int(template_cfg.get("sheet", 0))]

    data_start_row = int(template_cfg.get("data_start_row", 2))
    account_col = template_cfg.get("account_col", "A")
    account_idx = column_index_from_string(account_col)

    accounts: Dict[str, int] = {}
    for row in range(data_start_row, ws.max_row + 1):
        value = ws.cell(row=row, column=account_idx).value
        account = normalize_account(value)
        if account is None:
            continue
        accounts[account] = row

    return wb, accounts


def write_output(
    cfg: Config,
    wb: Any,
    accounts: Dict[str, int],
    source_df: pd.DataFrame,
    output_path: Path,
    period: ReportPeriod,
    rows_before_filter: int,
    rows_after_filter: int,
) -> RunSummary:
    template_cfg = cfg.template
    write_cols = template_cfg["write_columns"]

    col_targets = {
        key: column_index_from_string(letter)
        for key, letter in write_cols.items()
    }

    ws = wb.worksheets[int(template_cfg.get("sheet", 0))]
    source_map = {row["account"]: row for _, row in source_df.iterrows()}

    totals = {"debit": 0.0, "credit": 0.0}
    closing_issues = []
    tol = float(cfg.rules.get("closing_tolerance", 0.01))
    debit_field = cfg.rules.get("debit_field", "debit")
    credit_field = cfg.rules.get("credit_field", "credit")
    debit_credit_tol = float(cfg.rules.get("debit_credit_tolerance", 0.01))
    closing_fields = cfg.rules.get("closing_check_fields", {})

    for account, row_idx in accounts.items():
        row = source_map.get(account)
        for field, col_idx in col_targets.items():
            value = row[field] if row is not None else 0.0
            ws.cell(row=row_idx, column=col_idx, value=value)

        if debit_field in col_targets:
            totals["debit"] += row[debit_field] if row is not None else 0.0
        if credit_field in col_targets:
            totals["credit"] += row[credit_field] if row is not None else 0.0

        if cfg.rules.get("closing_check", True) and closing_fields:
            open_f = closing_fields.get("opening_field")
            debit_f = closing_fields.get("debit_field")
            credit_f = closing_fields.get("credit_field")
            close_f = closing_fields.get("closing_field")
            if all([open_f, debit_f, credit_f, close_f]):
                opening = row[open_f] if row is not None else 0.0
                debit = row[debit_f] if row is not None else 0.0
                credit = row[credit_f] if row is not None else 0.0
                closing = row[close_f] if row is not None else 0.0
                expected = opening + debit - credit
                if abs(expected - closing) > tol:
                    closing_issues.append((account, expected, closing))

    if cfg.rules.get("debit_credit_check", True):
        diff = abs(totals["debit"] - totals["credit"])
        if diff > debit_credit_tol:
            print(
                f"WARNING: Debit/Credit mismatch: debit={totals['debit']:.2f} "
                f"credit={totals['credit']:.2f} diff={diff:.2f}"
            )

    if closing_issues:
        print(f"WARNING: Closing balance check failed for {len(closing_issues)} accounts")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    source_debit = float(source_df[debit_field].sum()) if debit_field in source_df else 0.0
    source_credit = (
        float(source_df[credit_field].sum()) if credit_field in source_df else 0.0
    )

    missing_accounts = []
    missing_df = source_df[~source_df["account"].isin(accounts.keys())]
    if debit_field in missing_df and credit_field in missing_df:
        for _, row in missing_df.iterrows():
            debit = float(row.get(debit_field, 0.0) or 0.0)
            credit = float(row.get(credit_field, 0.0) or 0.0)
            missing_accounts.append(
                MissingAccount(
                    account=str(row["account"]),
                    debit=debit,
                    credit=credit,
                    diff=debit - credit,
                )
            )
        missing_accounts.sort(key=lambda item: abs(item.diff), reverse=True)

    return RunSummary(
        selected_year=period.year,
        selected_month=period.month,
        rows_before_filter=rows_before_filter,
        rows_after_filter=rows_after_filter,
        debit_field=debit_field,
        credit_field=credit_field,
        source_debit=source_debit,
        source_credit=source_credit,
        source_diff=source_debit - source_credit,
        template_debit=float(totals["debit"]),
        template_credit=float(totals["credit"]),
        template_diff=float(totals["debit"] - totals["credit"]),
        tolerance=debit_credit_tol,
        missing_accounts=missing_accounts,
    )


def write_run_report(output_path: Path, summary: RunSummary) -> Path:
    report_path = output_path.with_name(f"{output_path.stem}_run_report.txt")
    lines = [
        "Trial Balance Run Report",
        f"Generated from: {output_path.name}",
        "",
        f"Selected latest month: {summary.selected_year}/{summary.selected_month:02d}",
        f"Rows before filter: {summary.rows_before_filter}",
        f"Rows after filter: {summary.rows_after_filter}",
        "",
        f"Debit field: {summary.debit_field}",
        f"Credit field: {summary.credit_field}",
        "",
        "Source totals (all accounts from input):",
        f"  debit  = {summary.source_debit:.2f}",
        f"  credit = {summary.source_credit:.2f}",
        f"  diff   = {summary.source_diff:.2f}",
        "",
        "Template-covered totals (accounts present in template):",
        f"  debit  = {summary.template_debit:.2f}",
        f"  credit = {summary.template_credit:.2f}",
        f"  diff   = {summary.template_diff:.2f}",
        f"  tolerance = {summary.tolerance:.2f}",
        "",
    ]

    if abs(summary.template_diff) <= summary.tolerance:
        lines.append("Status: OK (template MD/DAL check within tolerance)")
    else:
        lines.append("Status: MISMATCH (template MD/DAL check failed)")
        lines.append(
            "Most common cause: account appears in input but is missing in template."
        )

    lines.append("")
    lines.append(f"Missing accounts in template: {len(summary.missing_accounts)}")
    if summary.missing_accounts:
        lines.append("Account | Debit | Credit | Diff")
        for item in summary.missing_accounts:
            lines.append(
                f"{item.account} | {item.debit:.2f} | {item.credit:.2f} | {item.diff:.2f}"
            )

    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    return report_path


def build_output_path(cfg: Config, override: str | None) -> Path:
    if override:
        return Path(override)
    out_cfg = cfg.output
    date_format = out_cfg.get("date_format", "%Y-%m-%d")
    date_str = dt.date.today().strftime(date_format)
    filename = out_cfg.get("filename", "Trial Balance_{date}.xlsx").format(date=date_str)
    return Path(out_cfg.get("dir", "output")) / filename


def build_report_path(cfg: Config) -> Path:
    rep_cfg = cfg.report
    date_format = rep_cfg.get("date_format", "%Y-%m-%d")
    date_str = dt.date.today().strftime(date_format)
    filename = rep_cfg.get("filename", "Trial Balance_report_{date}.xlsx").format(
        date=date_str
    )
    return Path(rep_cfg.get("output_dir", "output")) / filename


def build_report(
    cfg: Config, source_output_path: Path, period: ReportPeriod
) -> Path | None:
    rep_cfg = cfg.report
    if not rep_cfg.get("enabled", False):
        return None

    df = pd.read_excel(source_output_path, sheet_name=0)

    wb = Workbook()
    ws = wb.active
    ws.title = rep_cfg.get("sheet_name", "Report")

    # Layout constants
    max_col = 8
    header_font = Font(name="Arial", size=9, bold=False)
    title_font = Font(name="Arial", size=10, bold=True)
    table_header_font = Font(name="Arial", size=9, bold=True)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    right = Alignment(horizontal="right", vertical="center", wrap_text=True)
    thin = Side(style="thin", color="000000")
    medium = Side(style="medium", color="000000")

    # Column widths
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 45
    for col in "CDEFGH":
        ws.column_dimensions[col].width = 14

    # Report header (rows 1-5)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    ws.cell(row=1, column=1, value=rep_cfg.get("company_line", "")).font = header_font
    ws.cell(row=1, column=1).alignment = left

    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=max_col)
    ws.cell(row=2, column=1, value=rep_cfg.get("company_address", "")).font = header_font
    ws.cell(row=2, column=1).alignment = left

    ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=max_col)
    ws.cell(row=3, column=1, value=rep_cfg.get("report_title", "")).font = title_font
    ws.cell(row=3, column=1).alignment = left

    ws.cell(row=4, column=1, value=rep_cfg.get("report_number", "")).font = header_font
    ws.cell(row=4, column=1).alignment = left
    ws.cell(row=4, column=6, value=rep_cfg.get("language", "")).font = header_font
    ws.cell(row=4, column=6).alignment = left
    ws.cell(row=4, column=7, value=dt.datetime.now().strftime(rep_cfg.get("datetime_format", "%d.%m.%Y/%H:%M"))).font = header_font
    ws.cell(row=4, column=7).alignment = left

    # Table header rows (6-7)
    ws.merge_cells(start_row=6, start_column=1, end_row=7, end_column=1)
    ws.cell(row=6, column=1, value="Konto\nnummer").font = table_header_font
    ws.cell(row=6, column=1).alignment = center
    # Ensure account number column stays as text (keep leading zeros).
    ws.column_dimensions["A"].number_format = "@"

    ws.merge_cells(start_row=6, start_column=2, end_row=7, end_column=2)
    ws.cell(row=6, column=2, value="Bezeichnung").font = table_header_font
    ws.cell(row=6, column=2).alignment = center

    ws.merge_cells(start_row=6, start_column=3, end_row=7, end_column=3)
    ws.cell(row=6, column=3, value="Erstsaldo").font = table_header_font
    ws.cell(row=6, column=3).alignment = center

    ws.merge_cells(start_row=6, start_column=4, end_row=6, end_column=5)
    period_label = (
        f"Periode {period.year}/{period.month:02d} - {period.year}/{period.month:02d}"
    )
    ws.cell(row=6, column=4, value=period_label).font = table_header_font
    ws.cell(row=6, column=4).alignment = center
    ws.cell(row=7, column=4, value="Soll").font = table_header_font
    ws.cell(row=7, column=4).alignment = center
    ws.cell(row=7, column=5, value="Haben").font = table_header_font
    ws.cell(row=7, column=5).alignment = center

    ws.merge_cells(start_row=6, start_column=6, end_row=6, end_column=7)
    cumulative_label = f"Kumuliert {period.year}/01 - {period.year}/{period.month:02d}"
    ws.cell(row=6, column=6, value=cumulative_label).font = table_header_font
    ws.cell(row=6, column=6).alignment = center
    ws.cell(row=7, column=6, value="Soll").font = table_header_font
    ws.cell(row=7, column=6).alignment = center
    ws.cell(row=7, column=7, value="Haben").font = table_header_font
    ws.cell(row=7, column=7).alignment = center

    ws.merge_cells(start_row=6, start_column=8, end_row=7, end_column=8)
    ws.cell(row=6, column=8, value="Saldo").font = table_header_font
    ws.cell(row=6, column=8).alignment = center

    # Data rows start
    start_row = 8
    # Use dot in format code so Excel maps to locale correctly (comma decimal locales
    # will still display commas), and keep space as thousands separator.
    num_fmt = "# ##0.00;-# ##0.00;0.00"

    def num(value: Any) -> float:
        try:
            return round(float(value or 0.0), 2)
        except (TypeError, ValueError):
            return 0.0

    def account_class(value: Any) -> str | None:
        if value is None or value == "":
            return None
        text = str(value).strip()
        if not text:
            return None
        # keep digits only for class detection
        digits = "".join(ch for ch in text if ch.isdigit())
        if not digits:
            return None
        return digits.zfill(6)[0]

    def write_data_row(r: int, row: pd.Series) -> None:
        acct = row.get("sal_konto", "")
        acct_text = str(acct).strip()
        acct_digits = "".join(ch for ch in acct_text if ch.isdigit())
        acct_value = acct_digits.zfill(6) if acct_digits else acct_text
        cell = ws.cell(row=r, column=1, value=acct_value)
        cell.alignment = left
        cell.number_format = "@"
        ws.cell(row=r, column=2, value=row.get("kos_bezeich_2", "")).alignment = left
        ws.cell(row=r, column=3, value=num(row.get("eb_saldo", 0.0))).number_format = num_fmt
        ws.cell(row=r, column=4, value=num(row.get("peri_soll", 0.0))).number_format = num_fmt
        ws.cell(row=r, column=5, value=num(row.get("peri_haben", 0.0))).number_format = num_fmt
        ws.cell(row=r, column=6, value=num(row.get("kum_soll", 0.0))).number_format = num_fmt
        ws.cell(row=r, column=7, value=num(row.get("kum_haben", 0.0))).number_format = num_fmt
        ws.cell(row=r, column=8, value=num(row.get("saldo", 0.0))).number_format = num_fmt
        for c in range(3, 9):
            ws.cell(row=r, column=c).alignment = right

    def write_subtotal_row(r: int, klass: str, totals: dict[str, float]) -> None:
        ws.cell(row=r, column=2, value=f"Summe Kontenklasse {klass}").font = table_header_font
        ws.cell(row=r, column=2).alignment = left
        ws.cell(row=r, column=3, value=num(totals["eb_saldo"])).number_format = num_fmt
        ws.cell(row=r, column=4, value=num(totals["peri_soll"])).number_format = num_fmt
        ws.cell(row=r, column=5, value=num(totals["peri_haben"])).number_format = num_fmt
        ws.cell(row=r, column=6, value=num(totals["kum_soll"])).number_format = num_fmt
        ws.cell(row=r, column=7, value=num(totals["kum_haben"])).number_format = num_fmt
        ws.cell(row=r, column=8, value=num(totals["saldo"])).number_format = num_fmt
        for c in range(3, 9):
            ws.cell(row=r, column=c).alignment = right
        # Thick border above and below subtotal row
        for c in range(1, max_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = Border(top=medium, bottom=medium, left=cell.border.left, right=cell.border.right)

    current_class = None
    subtotal = {
        "saldo": 0.0,
        "eb_saldo": 0.0,
        "peri_soll": 0.0,
        "peri_haben": 0.0,
        "kum_soll": 0.0,
        "kum_haben": 0.0,
    }
    r = start_row
    for _, row in df.iterrows():
        # Skip rows where all numeric values are zero.
        values = [
            row.get("eb_saldo", 0.0),
            row.get("peri_soll", 0.0),
            row.get("peri_haben", 0.0),
            row.get("kum_soll", 0.0),
            row.get("kum_haben", 0.0),
            row.get("saldo", 0.0),
        ]
        if all(float(v or 0.0) == 0.0 for v in values):
            continue
        klass = account_class(row.get("sal_konto"))
        if current_class is None:
            current_class = klass
        if klass is not None and current_class is not None and klass != current_class:
            write_subtotal_row(r, current_class, subtotal)
            r += 1
            subtotal = {k: 0.0 for k in subtotal}
            current_class = klass

        write_data_row(r, row)
        r += 1

        for key in subtotal.keys():
            value = row.get(key, 0.0)
            try:
                subtotal[key] += float(value)
            except (TypeError, ValueError):
                continue

    if current_class is not None:
        write_subtotal_row(r, current_class, subtotal)
        r += 1

    # Borders: vertical lines for all table columns and outer frame
    table_top = 6
    table_bottom = r - 1
    for row_idx in range(table_top, table_bottom + 1):
        for col_idx in range(1, max_col + 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            left_side = thin
            right_side = thin
            top_side = cell.border.top
            bottom_side = cell.border.bottom

            if col_idx == 1:
                left_side = medium
            if col_idx == max_col:
                right_side = medium
            if row_idx == table_top:
                top_side = medium
            if row_idx == table_bottom:
                bottom_side = medium

            cell.border = Border(
                left=left_side,
                right=right_side,
                top=top_side,
                bottom=bottom_side,
            )

    # Header bottom thick line (row 7)
    for c in range(1, max_col + 1):
        cell = ws.cell(row=7, column=c)
        cell.border = Border(
            left=cell.border.left,
            right=cell.border.right,
            top=cell.border.top,
            bottom=medium,
        )

    # Print setup
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = ws.ORIENTATION_PORTRAIT
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.print_area = f"A1:H{r - 1}"
    ws.print_title_rows = rep_cfg.get("header_rows", "1:7")
    ws.page_margins.left = 0.4
    ws.page_margins.right = 0.4
    ws.page_margins.top = 0.5
    ws.page_margins.bottom = 0.5
    ws.page_margins.header = 0.3
    ws.page_margins.footer = 0.3
    ws.oddFooter.right.text = "Seite: &P/&N"

    out_path = build_report_path(cfg)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)
    return out_path


def run_conversion(
    cfg: Config,
    input_path: Path,
    template_path: Path,
    output_override: str | None = None,
) -> RunResult:
    output_path = build_output_path(cfg, output_override)
    source_result = read_source(cfg, input_path)
    source_df = source_result.dataframe
    report_period = source_result.period
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    wb, accounts = read_template_accounts(cfg, output_path)
    summary = write_output(
        cfg,
        wb,
        accounts,
        source_df,
        output_path,
        period=report_period,
        rows_before_filter=source_result.rows_before_filter,
        rows_after_filter=source_result.rows_after_filter,
    )
    report_path = build_report(cfg, output_path, report_period)
    summary_path = write_run_report(output_path, summary)
    return RunResult(
        output_path=output_path,
        report_path=report_path,
        period=report_period,
        summary=summary,
        summary_path=summary_path,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Convert zostava.xlsx to Trial Balance template")
    parser.add_argument("--config", default="config.yaml", help="Path to config.yaml")
    parser.add_argument("--input", help="Override input Excel path")
    parser.add_argument("--template", help="Override template Excel path")
    parser.add_argument("--output", help="Override output Excel path")
    args = parser.parse_args()

    cfg = load_config(Path(args.config))

    input_path = Path(args.input or cfg.source["path"])
    template_path = Path(args.template or cfg.template["path"])
    result = run_conversion(cfg, input_path, template_path, args.output)

    print(f"Saved: {result.output_path}")
    if result.report_path:
        print(f"Report saved: {result.report_path}")
    print(f"Run report: {result.summary_path}")


if __name__ == "__main__":
    main()
