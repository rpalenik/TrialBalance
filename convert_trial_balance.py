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
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string


@dataclass(frozen=True)
class Config:
    source: Dict[str, Any]
    template: Dict[str, Any]
    output: Dict[str, Any]
    rules: Dict[str, Any]


def load_config(path: Path) -> Config:
    data = yaml.safe_load(path.read_text(encoding="utf-8"))
    return Config(
        source=data["source"],
        template=data["template"],
        output=data["output"],
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


def read_source(cfg: Config, input_path: Path) -> pd.DataFrame:
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

    data = df.iloc[data_start_row - 1 :].copy()

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

    return data


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
) -> None:
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
        if diff > float(cfg.rules.get("debit_credit_tolerance", 0.01)):
            print(
                f"WARNING: Debit/Credit mismatch: debit={totals['debit']:.2f} "
                f"credit={totals['credit']:.2f} diff={diff:.2f}"
            )

    if closing_issues:
        print(f"WARNING: Closing balance check failed for {len(closing_issues)} accounts")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def build_output_path(cfg: Config, override: str | None) -> Path:
    if override:
        return Path(override)
    out_cfg = cfg.output
    date_format = out_cfg.get("date_format", "%Y-%m-%d")
    date_str = dt.date.today().strftime(date_format)
    filename = out_cfg.get("filename", "Trial Balance_{date}.xlsx").format(date=date_str)
    return Path(out_cfg.get("dir", "output")) / filename


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
    output_path = build_output_path(cfg, args.output)

    source_df = read_source(cfg, input_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(template_path, output_path)
    wb, accounts = read_template_accounts(cfg, output_path)
    write_output(cfg, wb, accounts, source_df, output_path)

    print(f"Saved: {output_path}")


if __name__ == "__main__":
    main()
