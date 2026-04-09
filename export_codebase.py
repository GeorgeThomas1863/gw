"""
export_codebase.py — Export all source files to a single text file.

Output location is resolved from config.EXPORT_OUTPUT_PATH; if that is
absent or blank the script prompts for a path at runtime.
"""
from __future__ import annotations

import sys
from pathlib import Path

# ---------------------------------------------------------------------------
# Resolve output path
# ---------------------------------------------------------------------------
_REPO_ROOT = Path(__file__).resolve().parent

try:
    import config
    _config_path: str = getattr(config, "EXPORT_OUTPUT_PATH", "")
except Exception:
    _config_path = ""

if _config_path:
    OUTPUT_PATH = Path(_config_path)
else:
    raw = input("Export output path (e.g. C:\\exports\\codebase_export.txt): ").strip()
    if not raw:
        print("No path provided. Aborting.")
        sys.exit(1)
    OUTPUT_PATH = Path(raw)

# ---------------------------------------------------------------------------
# Files to include
# ---------------------------------------------------------------------------
SKIP_DIRS = {".venv", "__pycache__", ".git", ".pytest_cache"}

OPTIONAL_EXTRAS = [
    "requirements.txt",
    "SETUP.txt",
    "CLAUDE.md",
    "current_packages.txt",
]

DELIMITER = "+" * 50


def _collect_py_files() -> list[Path]:
    """Return all .py files under the repo root, sorted by path, skipping noise dirs."""
    results: list[Path] = []
    for path in sorted(_REPO_ROOT.rglob("*.py")):
        if any(part in SKIP_DIRS for part in path.parts):
            continue
        if path.name == Path(__file__).name:
            continue  # exclude this script itself
        results.append(path)
    return results


def _write_file_block(out, path: Path, label: str) -> None:
    out.write(f"{DELIMITER}\n")
    out.write(f"FILE: {label}\n")
    out.write(f"{DELIMITER}\n\n")
    try:
        out.write(path.read_text(encoding="utf-8", errors="replace"))
    except Exception as exc:
        out.write(f"[ERROR reading file: {exc}]\n")
    out.write("\n\n")


def main() -> None:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    extras: list[Path] = []
    for name in OPTIONAL_EXTRAS:
        p = _REPO_ROOT / name
        if p.exists():
            extras.append(p)

    py_files = _collect_py_files()
    all_files = extras + py_files

    with OUTPUT_PATH.open("w", encoding="utf-8") as out:
        out.write(f"GrayWolfe Codebase Export\n")
        out.write(f"Files: {len(all_files)}\n")
        out.write(f"{DELIMITER}\n\n")

        for path in all_files:
            label = path.relative_to(_REPO_ROOT).as_posix()
            _write_file_block(out, path, label)

    print(f"Exported {len(all_files)} files -> {OUTPUT_PATH}")
    print(f"  {len(extras)} config/doc files, {len(py_files)} .py files")


if __name__ == "__main__":
    main()
