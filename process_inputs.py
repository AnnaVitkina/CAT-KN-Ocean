"""
Interactive terminal tool: map input files to rate types, convert to DataFrames,
and save as xlsx in the processing folder with `_processed` in the name.

Prompts for:
  - main rates (single file; includes Inland-outport rates for on/pre-carriage)
  - BAF fee (one or more files)
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd

import paths as project_paths

ROOT = project_paths.ROOT
INPUT_DIR = project_paths.INPUT_DIR
PROCESSING_DIR = project_paths.PROCESSING_DIR


def _sync_dirs() -> None:
    """Pick up Colab/Drive paths refreshed by run_pipeline."""
    global ROOT, INPUT_DIR, PROCESSING_DIR
    ROOT = project_paths.ROOT
    INPUT_DIR = project_paths.INPUT_DIR
    PROCESSING_DIR = project_paths.PROCESSING_DIR

COST_TYPES = [
    ("main_rates", "main rates"),
    ("baf_fee", "BAF fee"),
]

# Cost types that accept multiple input files
MULTI_FILE_COST_TYPES = {"baf_fee"}

SUPPORTED_EXTENSIONS = {".xlsx", ".xls", ".csv"}


def list_input_files() -> list[Path]:
    _sync_dirs()
    if not INPUT_DIR.exists():
        raise FileNotFoundError(f"Input folder not found: {INPUT_DIR}")

    files = sorted(
        p
        for p in INPUT_DIR.iterdir()
        if p.is_file()
        and not p.name.startswith("~$")
        and p.suffix.lower() in SUPPORTED_EXTENSIONS
    )
    return files


def prompt_file_selection(
    files: list[Path],
) -> dict[str, Path | list[Path] | None]:
    print("\nFiles available in input/:\n")
    for i, path in enumerate(files, start=1):
        print(f"  [{i}] {path.name}")
    print("\nEnter a file number to assign, or press Enter to skip that cost.")
    print("For BAF fee you may select multiple files: e.g. 1,3 or 1 3\n")

    selections: dict[str, Path | list[Path] | None] = {}
    used_indexes: set[int] = set()

    for key, label in COST_TYPES:
        allow_multi = key in MULTI_FILE_COST_TYPES
        while True:
            hint = "numbers" if allow_multi else "number"
            raw = input(
                f"Select file{'s' if allow_multi else ''} for {label} "
                f"(or Enter to skip): "
            ).strip()
            if not raw:
                selections[key] = None
                print(f"  -> skipped {label}\n")
                break

            if allow_multi:
                tokens = [t for t in raw.replace(",", " ").split() if t]
                if not tokens or not all(t.isdigit() for t in tokens):
                    print(f"  Please enter {hint} like 1,3 or press Enter to skip.")
                    continue
                indexes = [int(t) for t in tokens]
            else:
                if not raw.isdigit():
                    print(f"  Please enter a valid {hint} or press Enter to skip.")
                    continue
                indexes = [int(raw)]

            if any(i < 1 or i > len(files) for i in indexes):
                print(f"  Each number must be between 1 and {len(files)}.")
                continue

            if any(i in used_indexes for i in indexes):
                print("  One or more files are already assigned to another cost type.")
                continue

            if len(indexes) != len(set(indexes)):
                print("  Duplicate numbers are not allowed.")
                continue

            chosen = [files[i - 1] for i in indexes]
            for i in indexes:
                used_indexes.add(i)

            if allow_multi:
                selections[key] = chosen
                print("  -> " + ", ".join(p.name for p in chosen) + "\n")
            else:
                selections[key] = chosen[0]
                print(f"  -> {chosen[0].name}\n")
            break

    return selections


def read_file_to_frames(path: Path) -> dict[str, pd.DataFrame]:
    suffix = path.suffix.lower()

    if suffix == ".csv":
        return {"Sheet1": pd.read_csv(path)}

    if suffix == ".xls":
        workbook = pd.read_excel(path, sheet_name=None, engine="xlrd")
        return workbook

    if suffix == ".xlsx":
        workbook = pd.read_excel(path, sheet_name=None, engine="openpyxl")
        return workbook

    raise ValueError(f"Unsupported file type: {path.suffix}")


def processed_output_path(source: Path) -> Path:
    _sync_dirs()
    return PROCESSING_DIR / f"{source.stem}_processed.xlsx"


def _normalize_header(name) -> str:
    return " ".join(str(name).strip().split())


def _find_inland_sheet_name(sheet_names: list[str]) -> str | None:
    for name in sheet_names:
        key = name.lower().replace(" ", "")
        if "inland" in key and "outport" in key:
            return name
    return None


def _load_inland_outport_table(path: Path, sheet_name: str) -> pd.DataFrame:
    """Load Inland-outport rates with the real header row detected."""
    suffix = path.suffix.lower()
    engine = "xlrd" if suffix == ".xls" else "openpyxl"
    raw = pd.read_excel(path, sheet_name=sheet_name, header=None, engine=engine)

    header_row = None
    for i in range(min(20, len(raw))):
        values = [_normalize_header(v) for v in raw.iloc[i].tolist()]
        if "From Service" in values and "To Service" in values:
            header_row = i
            break
    if header_row is None:
        raise ValueError(
            f"Could not find From Service / To Service headers in '{sheet_name}'"
        )

    headers = [_normalize_header(v) for v in raw.iloc[header_row].tolist()]
    df = raw.iloc[header_row + 1 :].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)
    return df


ALLOWED_SERVICES = {"D", "CFS-CY", "CFS-P"}


def clean_and_split_inland_outport(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    Clean Inland-outport rates and split into pre-carriage / on-carriage.

    Remove rows where From Service or To Service is not in {D, CFS-CY, CFS-P}.
    Remove D => D lanes.
    Remove Base Rate Basis 'per unit' only for on-carriage lanes where
    From Country or To Country is MX.

    - pre-carriage: From Service == 'D' (and not D => D)
    - on-carriage:  To Service == 'D' (and not D => D)
    """
    work = df.copy()
    work.columns = [_normalize_header(c) for c in work.columns]

    if "From Service" not in work.columns or "To Service" not in work.columns:
        raise ValueError("Inland sheet missing From Service / To Service columns")

    from_svc = work["From Service"].astype(str).str.strip()
    to_svc = work["To Service"].astype(str).str.strip()
    service_ok = from_svc.isin(ALLOWED_SERVICES) & to_svc.isin(ALLOWED_SERVICES)
    not_d_to_d = ~((from_svc == "D") & (to_svc == "D"))

    cleaned = work.loc[service_ok & not_d_to_d].copy().reset_index(drop=True)
    removed = len(work) - len(cleaned)

    from_is_d = cleaned["From Service"].astype(str).str.strip() == "D"
    to_is_d = cleaned["To Service"].astype(str).str.strip() == "D"

    pre = cleaned.loc[from_is_d].copy().reset_index(drop=True)
    on_ = cleaned.loc[to_is_d].copy().reset_index(drop=True)

    # per unit: drop only on-carriage MX lanes (From Country or To Country)
    if "Base Rate Basis" in on_.columns:
        basis = on_["Base Rate Basis"].astype(str).str.strip().str.lower()
        from_country = (
            on_["From Country"].astype(str).str.strip().str.upper()
            if "From Country" in on_.columns
            else pd.Series("", index=on_.index)
        )
        to_country = (
            on_["To Country"].astype(str).str.strip().str.upper()
            if "To Country" in on_.columns
            else pd.Series("", index=on_.index)
        )
        mx_lane = (from_country == "MX") | (to_country == "MX")
        drop_per_unit = (basis == "per unit") & mx_lane
        removed += int(drop_per_unit.sum())
        on_ = on_.loc[~drop_per_unit].copy().reset_index(drop=True)

    return {
        "pre-carriage": pre,
        "on-carriage": on_,
        "_removed": removed,
        "_cleaned_total": len(pre) + len(on_),
    }


def prepare_main_rates_frames(path: Path, frames: dict[str, pd.DataFrame]) -> dict[str, pd.DataFrame]:
    """
    For main rates: replace Inland-outport rates with cleaned
    pre-carriage and on-carriage sheets.
    """
    inland_name = _find_inland_sheet_name(list(frames.keys()))
    if inland_name is None:
        print("  warning: Inland-outport rates sheet not found — skipped inland split")
        return frames

    inland_df = _load_inland_outport_table(path, inland_name)
    split = clean_and_split_inland_outport(inland_df)

    out: dict[str, pd.DataFrame] = {}
    for name, df in frames.items():
        if name == inland_name:
            continue
        out[name] = df
    out["pre-carriage"] = split["pre-carriage"]
    out["on-carriage"] = split["on-carriage"]

    print(
        f"  inland cleaned -> pre-carriage: {len(split['pre-carriage'])} rows, "
        f"on-carriage: {len(split['on-carriage'])} rows "
        f"(kept {split['_cleaned_total']} of {len(inland_df)}, "
        f"removed {split['_removed']})"
    )
    return out


def save_frames(frames: dict[str, pd.DataFrame], output_path: Path) -> None:
    _sync_dirs()
    PROCESSING_DIR.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        for sheet_name, df in frames.items():
            safe_name = str(sheet_name)[:31] or "Sheet1"
            df.to_excel(writer, sheet_name=safe_name, index=False)


def _as_source_list(value: Path | list[Path] | None) -> list[Path]:
    if value is None:
        return []
    if isinstance(value, list):
        return value
    return [value]


def process_selections(
    selections: dict[str, Path | list[Path] | None],
) -> dict[str, Path | list[Path]]:
    """Process selected files. Returns cost_key -> path or list of paths."""
    saved: dict[str, Path | list[Path]] = {}

    for key, label in COST_TYPES:
        sources = _as_source_list(selections.get(key))
        if not sources:
            continue

        if key in MULTI_FILE_COST_TYPES:
            outputs: list[Path] = []
            for source in sources:
                print(f"Processing {label}: {source.name} ...")
                frames = read_file_to_frames(source)
                output_path = processed_output_path(source)
                save_frames(frames, output_path)
                outputs.append(output_path)
                print(f"  saved -> {output_path.name} ({len(frames)} sheet(s))")
            saved[key] = outputs
        else:
            source = sources[0]
            print(f"Processing {label}: {source.name} ...")
            frames = read_file_to_frames(source)
            if key == "main_rates":
                frames = prepare_main_rates_frames(source, frames)
            output_path = processed_output_path(source)
            save_frames(frames, output_path)
            saved[key] = output_path
            print(f"  saved -> {output_path.name} ({len(frames)} sheet(s))")

    return saved


def run() -> dict[str, Path | list[Path]]:
    """Interactive entry: prompt, process, return mapping of cost type -> output path."""
    files = list_input_files()
    if not files:
        print(f"No supported files found in {INPUT_DIR}")
        print(f"Supported types: {', '.join(sorted(SUPPORTED_EXTENSIONS))}")
        return {}

    selections = prompt_file_selection(files)

    if all(
        (v is None or v == [])
        for v in selections.values()
    ):
        print("Nothing selected. Exiting.")
        return {}

    print("\n--- Processing ---")
    saved = process_selections(selections)

    print("\nDone.")
    labels = dict(COST_TYPES)
    for key, path in saved.items():
        if isinstance(path, list):
            print(f"  {labels[key]}: {[p.name for p in path]}")
        else:
            print(f"  {labels[key]}: {path}")

    return saved


if __name__ == "__main__":
    run()
