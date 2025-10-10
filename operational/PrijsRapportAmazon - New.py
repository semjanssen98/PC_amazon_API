# -*- coding: utf-8 -*-
"""Amazon Payment CSV → consolidated XLSX (with reconciliation)
================================================================

*   **Modular & readable** — everything lives in small testable functions.
*   **No global state** — run‐time parameters are passed in a `Config` dataclass.
*   **Language‑agnostic translations** are cached once, then reused.
*   **Single pandas pipeline** per marketplace (no manual per‑cell loops).
*   **Reconciliation table** is printed automatically.
*   Still driven by the four variables you edit at the top: `MONTH`, `YEAR`,
    `CLIENT`, `MARKETS`.

Usage
-----

```powershell
python payment_report_merge.py  # edit the constants below first
```
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List
import re
import csv

import pandas as pd
import openpyxl

# ---------------------------------------------------------------------------
# 1. User‑editable parameters
# ---------------------------------------------------------------------------

MONTH:   str = "September"  # full month name, e.g. "June"
YEAR:    str = "2025"
CLIENT:  str = "PAT" # client prefix, e.g. "PAT"
MARKETS: List[str] = ["DE", "FR", "NL"]   # marketplace 2‑letter codes

BASE_DIR = Path("/home/semja/github/personal/PC_amazon_API/operational")

# ---------------------------------------------------------------------------
# 2. Derived constants (rarely change)
# ---------------------------------------------------------------------------

_MONTH_NUM = {
    m: i for i, m in enumerate(
        ["January", "February", "March", "April", "May", "June",
         "July", "August", "September", "October", "November", "December"],
        start=1,
    )
}

FINAL_COLS: List[str] = [
    "country", "date/time", "settlement id", "type", "order id", "sku",
    "description", "quantity", "marketplace", "fulfilment", "order city",
    "order state", "order postal", "product sales", "product sales tax",
    "postage credits", "shipping credits tax", "gift wrap credits",
    "gift wrap credits tax", "promotional rebates",
    "promotional rebates tax", "marketplace withheld tax", "selling fees",
    "fba fees", "other transactions fees", "other", "total",
]

MONEY_COLS = [c for c in FINAL_COLS if c not in {
    "country", "date/time", "settlement id", "type", "order id", "sku",
    "description", "quantity", "marketplace", "fulfilment",
    "order city", "order state", "order postal",
}]

# ---------------------------------------------------------------------------
# 3. Dataclass configuration object
# ---------------------------------------------------------------------------

@dataclass(frozen=True)
class Config:
    month: str
    year: str
    client: str
    markets: List[str]
    base_dir: Path

    @property
    def month_num(self) -> int:  # 6 for "June"
        return _MONTH_NUM[self.month]

    @property
    def input_dir(self) -> Path:
        return self.base_dir / "Input"

    @property
    def output_dir(self) -> Path:
        return self.base_dir / "Output" / self.client

    @property
    def translation_wb(self) -> Path:
        return self.base_dir / "Payments report link vertalingen.xlsx"


CFG = Config(MONTH, YEAR, CLIENT, MARKETS, BASE_DIR)
CFG.output_dir.mkdir(parents=True, exist_ok=True)

# ---------------------------------------------------------------------------
# 4. Translation dictionaries (cached once)
# ---------------------------------------------------------------------------

def build_translation_dicts(wb_path: Path) -> tuple[Dict[str, str], Dict[str, str]]:
    """Return (header_map, payment_type_map) using **all** populated rows.

    The workbook layout is fixed only in the sense that row 2 is English.
    Any subsequent row with at least one non‑empty header cell is treated as a
    local‑language row and included automatically—so you can add new
    marketplaces without touching the code.
    """
    wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    ws_h, ws_p = wb[wb.sheetnames[0]], wb[wb.sheetnames[1]]

    # -------- Column headers ------------------------------------------------
    eng_hdr = [(ws_h.cell(2, c).value or "").strip() for c in range(2, 29)]
    col_map: Dict[str, str] = {}

    row = 3
    while True:
        header_cells = [ws_h.cell(row, c).value for c in range(2, 29)]
        if all(v in (None, "") for v in header_cells):
            break  # reached the first fully blank row → stop scanning

        for c, eng in zip(range(2, 29), eng_hdr):
            local = (ws_h.cell(row, c).value or "").strip()
            if local:
                col_map[local.lower()] = eng
        row += 1

    # -------- Payment types -------------------------------------------------
    eng_types: List[str] = []
    col_idx = 1
    while (val := ws_p.cell(2, col_idx).value):
        eng_types.append(str(val).strip())
        col_idx += 1

    pay_map: Dict[str, str] = {}
    row = 3
    while True:
        if all((ws_p.cell(row, i).value in (None, "") for i in range(1, len(eng_types) + 1))):
            break
        for idx, eng in enumerate(eng_types, start=1):
            local = (ws_p.cell(row, idx).value or "").strip()
            if local:
                pay_map[local.lower()] = eng
        row += 1

    return col_map, pay_map


COL_MAP, PAY_MAP = build_translation_dicts(CFG.translation_wb)

# ---------------------------------------------------------------------------
# 5. Utility functions
# ---------------------------------------------------------------------------

_EU_NBSP = "\u202f"


def parse_eu_number(text: str) -> float:
    """Convert '1 234,56' → 1234.56. Returns 0.0 for blank strings."""
    if not text:
        return 0.0
    clean = (text.replace(_EU_NBSP, "")
                  .replace(" ", "")
                  .replace(".", "")
                  .replace(",", "."))
    return float(clean)


def format_eu(amount: float) -> str:
    sign = "- " if amount < 0 else ""
    amt = abs(amount)
    parts = f"{amt:,.2f}".replace(",", " ").replace(".", ",")
    return f"{sign}€ {parts}"


def reformat_date(series: pd.Series, month_num: int, german: bool) -> pd.Series:
    """Return dd-mm-yyyy regardless of locale."""
    if german:
        pattern = r"(\d{1,2})\.(\d{1,2})\.(\d{4})"
        func = lambda m: f"{m.group(1)}-{month_num}-{m.group(3)}"
    else:
        pattern = r"(\d{1,2})\s+\S+\s+(\d{4})"
        func = lambda m: f"{m.group(1)}-{month_num}-{m.group(2)}"
    return series.str.replace(pattern, func, regex=True)


# ---------------------------------------------------------------------------
# 6. Per‑marketplace processing
# ---------------------------------------------------------------------------

def load_marketplace_csv(cc: str) -> pd.DataFrame:
    csv_file = CFG.input_dir / f"{CFG.year}_{CFG.month_num}_Date_Range_Reports_{CFG.client}_{cc}.csv"
    print(f"\n📥  Processing {csv_file}")
    df = pd.read_csv(csv_file, skiprows=7, dtype=str, quoting=csv.QUOTE_MINIMAL).fillna("")

    # Rename columns using translation map
    df = df.rename(columns={c: COL_MAP[c.lower()] for c in df.columns if c.lower() in COL_MAP})

    # Ensure all expected columns exist (empty if absent)
    for col in FINAL_COLS:
        if col not in df.columns:
            df[col] = ""

    # Translate payment‑type values
    df["type"] = df["type"].map(lambda x: PAY_MAP.get(x.lower(), x))

    # Country tag & date normalisation
    df["country"] = cc
    df["date/time"] = reformat_date(df["date/time"], CFG.month_num, german=(cc == "DE"))

    # Remove transfer rows
    df = df[df["type"] != "Transfer"].copy()

    # Money cols ➜ EU formatting for output (but keep numeric agg copy)
    for col in MONEY_COLS:
        df[f"_{col}_float"] = df[col].apply(parse_eu_number)  # helper col for agg
        df[col] = (df[col]
                    .str.replace(_EU_NBSP, "")
                    .str.replace(".", ",", regex=False))

    return df[FINAL_COLS + [f"_{c}_float" for c in MONEY_COLS]]


# ---------------------------------------------------------------------------
# 7. Main
# ---------------------------------------------------------------------------

def main() -> None:
    frames = [load_marketplace_csv(cc) for cc in CFG.markets]

    if not frames:
        print("No marketplaces processed, exiting.")
        return

    combined = pd.concat(frames, ignore_index=True)

    out_path = CFG.output_dir / f"{CFG.year}{CFG.month_num:02d}_{CFG.client}.xlsx"
    combined[FINAL_COLS].to_excel(out_path, index=False)

    # ---- reconciliation --------------------------------------------------
    metrics = {
        "product sales":   "_product sales_float",
        "selling fees":    "_selling fees_float",
        "fba fees":        "_fba fees_float",
        "total":           "_total_float",
    }

    print("\n✅  Consolidated payment report written to:\n   ", out_path)
    print("\nReconciliation (all marketplaces combined):\n")
    print(f"{'Metric':<20}{'Source CSV':>18}{'Output XLSX':>18}{'Match?':>8}")

    for m, helper_col in metrics.items():
        src_sum = combined[helper_col].sum()
        out_sum = combined[m].apply(parse_eu_number).sum()
        ok = abs(src_sum - out_sum) < 1e-6
        print(f"{m:<20}{format_eu(src_sum):>18}{format_eu(out_sum):>18}{'✅' if ok else '❌':>8}")


if __name__ == "__main__":
    main()