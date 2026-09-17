"""
End-to-end CAT Ocean K&N pipeline.

Steps:
  1) Ask which input files are main rates / BAF
     (BAF fee accepts multiple files)
  2) Convert chosen files -> processing/*_processed.xlsx
  3) Build result matrix:
       - Result tab (main rates + missing BAF UIDs in green)
       - BAF tab (all BAF files, costs ordered by validity)
       - Missing lanes tab (non-GB missing BAF UIDs)
       - Accessorial costs tab
     -> output/{main_file}_result.xlsx

On/pre-carriage data comes from the main file sheet 'Inland-outport rates'
(no separate input files).

------------------------------------------------------------------------------
Google Colab
------------------------------------------------------------------------------
Code lives in:  /content/CAT-KN-Ocean
Data (Drive):   .../RMT_CAT/RMT_CAT_Ocean_KN/{input,processing,output}

  !pip install -q pandas openpyxl xlrd
  from google.colab import drive
  drive.mount('/content/drive')

  import os
  exec(open('/content/CAT-KN-Ocean/run_pipeline.py').read())

------------------------------------------------------------------------------
Local machine
------------------------------------------------------------------------------
  python run_pipeline.py

Optional overrides:
  CAT_KN_DATA_ROOT  - folder that contains input/ processing/ output/
  CAT_KN_CODE_ROOT  - folder with the Python modules
  CAT_KN_FORCE_LOCAL=1 - force local paths even if Colab is detected
"""

from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _bootstrap() -> Path:
    """Ensure code root is on sys.path (needed when launched via exec())."""
    code_dir = Path("/content/CAT-KN-Ocean")
    if not code_dir.is_dir():
        try:
            code_dir = Path(__file__).resolve().parent
        except NameError:
            code_dir = Path.cwd()

    code_s = str(code_dir)
    if code_s not in sys.path:
        sys.path.insert(0, code_s)
    return code_dir


_bootstrap()

import paths as project_paths  # noqa: E402
import build_result_matrix as matrix  # noqa: E402
import process_inputs as inputs  # noqa: E402


def _ensure_dependencies() -> None:
    """Install runtime deps on Colab if missing."""
    if not project_paths.running_on_colab():
        return
    missing: list[str] = []
    for pkg, import_name in (("pandas", "pandas"), ("openpyxl", "openpyxl"), ("xlrd", "xlrd")):
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pkg)
    if not missing:
        return
    print(f"Installing: {', '.join(missing)}")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "-q", *missing])


def run_pipeline() -> dict:
    """
    Run the full interactive pipeline.

    Returns a summary dict with selections, processed paths, and result path.
    """
    _ensure_dependencies()
    resolved = project_paths.refresh_module_paths()
    inputs._sync_dirs()
    matrix._sync_dirs()

    print("=" * 60)
    print("CAT Ocean K&N — end-to-end pipeline")
    print("=" * 60)
    print(f"  mode:      {'Colab / Drive' if project_paths.running_on_colab() else 'local'}")
    print(f"  code:      {resolved['code_root']}")
    print(f"  data:      {resolved['root']}")
    print(f"  input:     {resolved['input']}")
    print(f"  processing:{resolved['processing']}")
    print(f"  output:    {resolved['output']}")

    # ------------------------------------------------------------------
    # Step 1: select input files
    # ------------------------------------------------------------------
    print("\n[1/3] Select input files")
    files = inputs.list_input_files()
    if not files:
        print(f"No supported files found in {inputs.INPUT_DIR}")
        print(f"Supported types: {', '.join(sorted(inputs.SUPPORTED_EXTENSIONS))}")
        return {"processed": {}, "result": None}

    selections = inputs.prompt_file_selection(files)
    if all(v is None or v == [] for v in selections.values()):
        print("Nothing selected. Exiting.")
        return {"processed": {}, "result": None, "selections": selections}

    # ------------------------------------------------------------------
    # Step 2: process selected files into processing/
    # ------------------------------------------------------------------
    print("\n[2/3] Process selected files -> processing/")
    processed = inputs.process_selections(selections)

    labels = dict(inputs.COST_TYPES)
    print("\nProcessed:")
    for key, path in processed.items():
        if isinstance(path, list):
            print(f"  {labels[key]}: {[p.name for p in path]}")
        else:
            print(f"  {labels[key]}: {path.name}")

    # ------------------------------------------------------------------
    # Step 3: build result matrix from main rates (+ BAF tab)
    # ------------------------------------------------------------------
    print("\n[3/3] Build result matrix -> output/")
    result_path: Path | None = None

    main_processed = processed.get("main_rates")
    if main_processed is None:
        print("  Main rates was skipped — result matrix not built.")
    else:
        source = selections.get("main_rates")
        rate_card = source.name if isinstance(source, Path) else None
        baf_paths = processed.get("baf_fee")
        result_path = matrix.build_result_matrix(
            main_processed,
            rate_card=rate_card,
            baf_path=baf_paths,
        )
        print(f"  Wrote {result_path}")

    # ------------------------------------------------------------------
    # Summary
    # ------------------------------------------------------------------
    print("\n" + "=" * 60)
    print("Pipeline complete")
    print("=" * 60)
    for key, label in inputs.COST_TYPES:
        src = selections.get(key)
        out = processed.get(key)
        if src is None or src == []:
            print(f"  {label}: skipped")
        elif isinstance(src, list):
            names = ", ".join(p.name for p in src)
            out_names = (
                ", ".join(p.name for p in out)
                if isinstance(out, list)
                else (out.name if out else "?")
            )
            print(f"  {label}: {names} -> {out_names}")
        else:
            print(f"  {label}: {src.name} -> {out.name if out else '?'}")
    if result_path is not None:
        print(f"  result matrix: {result_path}")
    else:
        print("  result matrix: not created")

    return {
        "selections": selections,
        "processed": processed,
        "result": result_path,
        "paths": resolved,
    }


# Runs for `python run_pipeline.py` and for Colab:
#   exec(open('/content/CAT-KN-Ocean/run_pipeline.py').read())
if __name__ == "__main__":
    run_pipeline()
else:
    # exec() in a notebook often keeps __name__ as "__main__", but if not,
    # still start when this file is the entry executed via open().read().
    try:
        _ = __file__
    except NameError:
        run_pipeline()
