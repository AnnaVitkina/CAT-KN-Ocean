"""
Project paths for local and Google Colab / Drive.

Data root (input / processing / output):
  - Colab: Google Shared Drive folder (see COLAB_DATA_ROOT)
  - Local: same folder as the code (project root)
  - Override: CAT_KN_DATA_ROOT env var

Code root (Python modules):
  - Colab: /content/CAT-KN-Ocean
  - Local: this package directory
  - Override: CAT_KN_CODE_ROOT env var
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

# Code on Colab (repo / uploaded project)
COLAB_CODE_ROOT = Path("/content/CAT-KN-Ocean")

# Data on Google Drive (Shared drive)
COLAB_DATA_ROOT = Path(
    "/content/drive/Shareddrives/FA Ops Europe: Rate Maintenance Team "
    "/Documents/AI Adoption RMT/RMT_CAT/RMT_CAT_Ocean_KN"
)


def running_on_colab() -> bool:
    if os.environ.get("CAT_KN_FORCE_LOCAL", "").strip().lower() in {"1", "true", "yes"}:
        return False
    try:
        import google.colab  # noqa: F401

        return True
    except ImportError:
        return Path("/content/drive").exists() and COLAB_CODE_ROOT.exists()


def code_root() -> Path:
    env = os.environ.get("CAT_KN_CODE_ROOT", "").strip()
    if env:
        return Path(env).expanduser().resolve()

    if COLAB_CODE_ROOT.is_dir():
        return COLAB_CODE_ROOT.resolve()

    try:
        return Path(__file__).resolve().parent
    except NameError:
        return Path.cwd().resolve()


def data_root() -> Path:
    env = os.environ.get("CAT_KN_DATA_ROOT", "").strip()
    if env:
        return Path(env).expanduser().resolve()

    if running_on_colab():
        return COLAB_DATA_ROOT

    return code_root()


def ensure_sys_path() -> Path:
    """Put code root on sys.path so `import process_inputs` works under exec()."""
    root = code_root()
    root_s = str(root)
    if root_s not in sys.path:
        sys.path.insert(0, root_s)
    return root


def mount_drive_if_colab() -> None:
    """Mount Google Drive when running in Colab (no-op locally)."""
    if not running_on_colab():
        return
    mount_point = Path("/content/drive")
    if mount_point.exists() and any(mount_point.iterdir()):
        return
    try:
        from google.colab import drive

        drive.mount("/content/drive")
    except Exception as exc:  # noqa: BLE001
        print(f"warning: could not mount Google Drive: {exc}")


def ensure_data_folders() -> dict[str, Path]:
    """
    Create input / processing / output under the data root.
    Returns dict of folder paths.
    """
    root = data_root()
    folders = {
        "root": root,
        "input": root / "input",
        "processing": root / "processing",
        "output": root / "output",
    }
    for key, path in folders.items():
        if key == "root":
            continue
        path.mkdir(parents=True, exist_ok=True)
    return folders


def resolve_paths(*, mount_drive: bool = True) -> dict[str, Path]:
    """
    Resolve and optionally prepare all project paths.
    Call once at pipeline start.
    """
    ensure_sys_path()
    if mount_drive:
        mount_drive_if_colab()
    folders = ensure_data_folders()
    return {
        "code_root": code_root(),
        **folders,
    }


# Defaults used by modules at import time (local-safe).
# run_pipeline() refreshes these after Colab mount via refresh_module_paths().
ROOT = code_root()
DATA_ROOT = data_root()
INPUT_DIR = DATA_ROOT / "input"
PROCESSING_DIR = DATA_ROOT / "processing"
OUTPUT_DIR = DATA_ROOT / "output"


def refresh_module_paths() -> dict[str, Path]:
    """Recompute path globals after Drive mount / env changes."""
    global ROOT, DATA_ROOT, INPUT_DIR, PROCESSING_DIR, OUTPUT_DIR
    resolved = resolve_paths(mount_drive=True)
    ROOT = resolved["code_root"]
    DATA_ROOT = resolved["root"]
    INPUT_DIR = resolved["input"]
    PROCESSING_DIR = resolved["processing"]
    OUTPUT_DIR = resolved["output"]
    return resolved
