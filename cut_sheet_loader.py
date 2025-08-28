"""
Cut Sheet Loader (v5) — per‑tab BA:BD import
============================================
Reads these tabs individually:
  - Rods_Straight_1
  - Rods_Straight_2
  - Rods_Straight_3
  - Rods_Curved

Expected headers in BA:BD on each tab (any case/spacing):
  Diameter_in | Material | Qty | Length

Outputs a tidy DataFrame with:
  ['diameter_in','Material','Qty','Length','tab']

Install deps:
    python -m pip install pandas openpyxl
"""
from __future__ import annotations

import sys
import types
from fractions import Fraction
from pathlib import Path
from typing import List

# --- sandbox guard for environments that expect 'micropip' -------------------
if "micropip" not in sys.modules:
    stub = types.ModuleType("micropip")
    stub.install = lambda *_, **__: None  # type: ignore[arg-type]
    sys.modules["micropip"] = stub

try:
    import pandas as pd  # type: ignore
except ModuleNotFoundError:
    print(
        "\nERROR: The 'pandas' library is required. Install it with:\n"
        "    python -m pip install pandas openpyxl\n",
        file=sys.stderr,
    )
    sys.exit(1)

# ----------------------- constants used by the optimizer ----------------------

KERF_INCHES: float = 1.0 / 8.0  # 0.125 in

STICK_LENGTHS: dict[tuple[str, float], int] = {
    ("C", 0.375): 20 * 12 - 0.75,
    ("C", 0.500): 22 * 12 - 0.75,
    ("C", 0.750): 24 * 12 - 0.75,
    ("C", 1.000): 24 * 12 - 0.75,
    ("C", 1.250): 24 * 12 - 0.75,
    ("C", 1.500): 24 * 12 - 0.75,
    ("304 PC", 0.375): 20 * 12 - 0.75,
    ("304 PC", 0.500): 20 * 12 - 0.75,
    ("304 PC", 0.750): 20 * 12 - 0.75,
    ("304 PC", 1.000): 20 * 12 - 0.75,
    ("304 PC", 1.250): 20 * 12 - 0.75,
    ("304 PC", 1.500): 20 * 12 - 0.75,
    ("304 OC", 0.375): 20 * 12 - 0.75,
    ("304 OC", 0.500): 20 * 12 - 0.75,
    ("304 OC", 0.750): 20 * 12 - 0.75,
    ("304 OC", 1.000): 20 * 12 - 0.75,
    ("304 OC", 1.250): 20 * 12 - 0.75,
    ("304 OC", 1.500): 20 * 12 - 0.75,
}

# ----------------------- tabs & parsing helpers -------------------------------

TABS: List[str] = [
    "Rods_Straight_1",
    "Rods_Straight_2",
    "Rods_Straight_3",
    "Rods_Curved",
]

def _norm_material(raw) -> str | None:
    """Uppercase, trim; blank/0/NONE/N/A -> None."""
    if raw is None or (isinstance(raw, float) and pd.isna(raw)):
        return None
    s = str(raw).strip().upper()
    if s in {"", "0", "NONE", "N/A"}:
        return None
    return s

def _to_float_lenient(x) -> float | None:
    """
    Parse numbers in common shop formats:
      71.5, 0.75, '3/4', '1 1/4', '1-1/4', '3/4"' (inch mark optional).
    Returns None if unparsable.
    """
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    s = (
        str(x)
        .strip()
        .rstrip('"')           # remove inch mark if present
        .replace("–", "-")
        .replace("—", "-")
    )
    # fast path: plain numeric
    try:
        return float(s)
    except ValueError:
        pass
    # normalize '1-1/4' -> '1 1/4'
    s = s.replace("-", " ")
    try:
        total = sum(Fraction(p) for p in s.split())
        return float(total)
    except Exception:
        return None

def _canonicalize_headers(cols: List[str]) -> List[str]:
    out = []
    for c in cols:
        k = c.strip().replace(" ", "_").replace("-", "_")
        k_up = k.upper()
        if k_up in {"DIAMETER_IN", "DIAMETER"}:
            out.append("diameter_in")
        elif k_up == "MATERIAL":
            out.append("Material")
        elif k_up in {"QTY", "QUANTITY"}:
            out.append("Qty")
        elif k_up == "LENGTH":
            out.append("Length")
        else:
            out.append(c.strip())
    return out

def _load_one_tab(path: Path, tab: str) -> pd.DataFrame:
    # Read only BA:BD; assume headers in the first row of that block.
    df = pd.read_excel(
        path,
        sheet_name=tab,
        usecols="BA:BD",
        engine="openpyxl",
        dtype={"Diameter_in": object, "Material": object, "Qty": object, "Length": object},
    )
    df.columns = _canonicalize_headers([str(c) for c in df.columns])

    # keep only expected columns (ignore any extras)
    cols = [c for c in ["diameter_in", "Material", "Qty", "Length"] if c in df.columns]
    df = df[cols]

    # clean types
    df["Material"] = df["Material"].apply(_norm_material)
    df["Qty"] = pd.to_numeric(df["Qty"], errors="coerce")
    df["Length"] = df["Length"].apply(_to_float_lenient)
    df["diameter_in"] = df["diameter_in"].apply(_to_float_lenient)

    # filter invalids
    df = df[df["Material"].notna()]
    df = df[df["Qty"].notna() & (df["Qty"] > 0)]
    df = df[df["Length"].notna() & (df["Length"] > 0)]
    df = df[df["diameter_in"].notna() & (df["diameter_in"] > 0)]

    df["tab"] = tab
    return df

# ----------------------- public API ------------------------------------------

def load_cut_demand(xls_path: str | Path) -> pd.DataFrame:
    """
    Load BA:BD from each tab and return one tidy DataFrame:
      columns = ['diameter_in','Material','Qty','Length','tab']
    """
    path = Path(xls_path)
    if not path.exists():
        raise FileNotFoundError(path)

    frames: List[pd.DataFrame] = []
    for tab in TABS:
        try:
            frames.append(_load_one_tab(path, tab))
        except Exception as exc:
            print(f"Warning: failed to load '{tab}': {exc}", file=sys.stderr)

    if not frames:
        raise RuntimeError("No tabs could be read. Check tab names and BA:BD headers.")
    return pd.concat(frames, ignore_index=True)

# ----------------------- CLI / smoke test ------------------------------------

if __name__ == "__main__":
    import argparse
    p = argparse.ArgumentParser(description="Load cut demand from BA:BD on four tabs.")
    p.add_argument("excel", help="Path to workbook (*.xlsm / *.xlsx)")
    args = p.parse_args()

    tidy = load_cut_demand(args.excel)
    print(tidy.head())
    print(f"Rows parsed: {len(tidy)}")
    print("Per‑tab counts:\n", tidy.groupby("tab")["Qty"].count())
