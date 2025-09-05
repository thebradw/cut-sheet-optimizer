"""
Cut Optimizer (v2.0)
====================
Generates a minimal-scrap cut sheet from a Fabrication Summaries workbook.

What's new in v2.0
------------------
- `optimise_cuts` now accepts an optional `kerf` argument to override the default.
- The core optimization logic is now parameterized with `kerf`.
- Output name is always **Cut_sheet.xlsx** in the workbook's folder.
- If Cut_sheet.xlsx is locked (open in Excel), write a timestamped fallback:
  Cut_sheet_YYYYMMDD_HHMMSS.xlsx
- Prints the exact output path and returns it from optimise_cuts().
"""
from __future__ import annotations

import argparse
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Tuple
from fractions import Fraction

import pandas as pd

# Local modules
from cut_sheet_loader import load_cut_demand, STICK_LENGTHS

# Default saw blade width, used if not provided.
DEFAULT_KERF_INCHES: float = 1.0 / 8.0


# ---------------------- helpers ----------------------------------------------

def _fmt_frac(val: float) -> str:
    """Format decimal inches → nearest 1/16-in string, e.g. 71.6875 → '71 11/16'."""
    frac = Fraction(val).limit_denominator(16)
    whole = int(frac)
    remainder = frac - whole
    if remainder == 0:
        return f"{whole}"
    if whole == 0:
        return f"{remainder}"
    return f"{whole} {remainder}"


def _first_fit_desc(cuts: List[float], stick_len: int, kerf: float) -> List[Tuple[List[float], float]]:
    """Simple first-fit-descending bin pack. Returns [(pattern, drop_in), ...]."""
    bins: List[List[float]] = []
    for length in sorted(cuts, reverse=True):
        placed = False
        for pat in bins:
            used = sum(pat) + kerf * max(len(pat) - 1, 0)
            if used + kerf + length <= stick_len:
                pat.append(length)
                placed = True
                break
        if not placed:
            bins.append([length])
    # Calculate drop for each pattern
    final_bins = []
    for pat in bins:
        num_cuts = len(pat)
        total_material_used = sum(pat)
        total_kerf_loss = kerf * (num_cuts - 1) if num_cuts > 1 else 0
        drop = stick_len - (total_material_used + total_kerf_loss)
        final_bins.append((pat, drop))
    return final_bins


def find_latest_workbook(folder: Path) -> Path:
    workbooks = [p for p in folder.glob("*.xls*") if p.is_file()]
    if not workbooks:
        raise FileNotFoundError(f"No Excel workbook found in {folder}")
    return max(workbooks, key=lambda p: p.stat().st_mtime)


def _build_rows(patterns, material, diameter, stick_len, tab):
    rows = []
    stick_id = 0
    for pat, drop in patterns:
        stick_id += 1
        rows.append({
            "Stick_ID": f"{material}-{diameter}-{stick_len}-{stick_id:04d}",
            "Material": material,
            "Diameter_in": diameter,
            "Stick_length_in": stick_len,
            "Pattern": ", ".join(_fmt_frac(p) for p in pat),
            "Pieces": len(pat),
            "Drop_in": drop,
            "tab": tab
        })
    return pd.DataFrame(rows)


# ---------------------- core group optimise ----------------------------------

def optimise_group(group: pd.DataFrame, material: str, diameter: float, tab: str, kerf: float) -> pd.DataFrame:
    """Runs the First-Fit-Descending heuristic for a single group of parts."""
    try:
        stick_len = STICK_LENGTHS[(material, diameter)]
    except KeyError:
        raise ValueError(f"No stock length defined for Material='{material}', Diameter={diameter}")

    demand = group.groupby("Length")["Qty"].sum().dropna().to_dict()

    # Create a flat list of all cuts required
    cuts_full = [l for l, q in demand.items() for _ in range(int(q))]

    # Run the packing algorithm
    patterns = _first_fit_desc(cuts_full, stick_len, kerf)

    # Format the results into a DataFrame
    return _build_rows(patterns, material, diameter, stick_len, tab)


# ---------------------- orchestrator -----------------------------------------

def optimise_cuts(workbook: Path, kerf: float | None = None) -> Path:
    """
    Run optimisation for all (tab, Material, diameter_in) groups and write the result.

    Args:
        workbook: Path to the source Excel workbook.
        kerf: Saw blade width in inches. Defaults to 0.125 if not provided.

    Returns:
        Path to the actual output file written (Cut_sheet.xlsx or timestamped fallback).
    """
    
    active_kerf = kerf if kerf is not None else DEFAULT_KERF_INCHES
    tidy = load_cut_demand(workbook)

    outputs = []
    for (tab, mat, dia), grp in tidy.groupby(["tab", "Material", "diameter_in"]):
        print(f"Optimising group: {mat} Ø{dia} with kerf={active_kerf}")
        try:
            result = optimise_group(grp, str(mat).strip(), float(dia), tab, kerf=active_kerf)
            outputs.append(result)
        except Exception as exc:
            print(f"Skipping {mat} Ø{dia}: {exc}")

    if not outputs:
        raise RuntimeError("No groups optimised – check input data and stock length definitions.")

    result_df = pd.concat(outputs, ignore_index=True)
    result_df = result_df.sort_values(by=["tab", "Material", "Diameter_in"])

    # --- File Output ---
    out_dir = workbook.parent
    fixed_name = out_dir / "Cut_sheet.xlsx"

    try:
        result_df.to_excel(fixed_name, index=False)
        out_path = fixed_name
        print(f"Cut sheet written to: {out_path}")
    except PermissionError:
        # File is likely open/locked; write a timestamped copy instead
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback = out_dir / f"Cut_sheet_{ts}.xlsx"
        result_df.to_excel(fallback, index=False)
        out_path = fallback
        print(f"'Cut_sheet.xlsx' was locked. Wrote fallback: {out_path}")

    return out_path


# ---------------------- CLI ---------------------------------------------------

def main() -> None:
    parser = argparse.ArgumentParser(description="Optimise saw cut sheet.")
    parser.add_argument("source", help="Workbook path or folder")
    parser.add_argument("--kerf", type=float, default=DEFAULT_KERF_INCHES, help="Saw blade width in inches.")
    args = parser.parse_args()

    src = Path(args.source)
    workbook = find_latest_workbook(src) if src.is_dir() else src
    print(f"Using workbook: {workbook}")

    t0 = datetime.now()
    out_path = optimise_cuts(workbook, kerf=args.kerf)
    print("Elapsed:", datetime.now() - t0)
    print(f"Output: {out_path}")

if __name__ == "__main__":
    # PuLP is not used in this version, so no guard is needed here.
    # If you re-introduce PuLP, add the dependency check back.
    main()
