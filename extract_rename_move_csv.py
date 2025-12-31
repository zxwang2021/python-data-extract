"""Extract all .zip files in the current folder.

For each zip:
- Extracts into a sibling folder named <zip_stem>__extracted
- Finds all .csv files inside (recursively)
- Renames them to <zip_stem>.csv (or <zip_stem>_2.csv, ... if multiple)
- Moves the renamed CSVs into the current folder

Usage (PowerShell):
    python .\extract_rename_move_csv.py

Optional flags:
  --keep-extracted   Keep extracted folders (no cleanup)
  --overwrite        Allow overwriting existing destination CSVs
"""

from __future__ import annotations

import argparse
import shutil
import sys
import zipfile
from pathlib import Path


def _safe_dest_name(zip_stem: str, index: int, total: int) -> str:
    if total <= 1 and index == 1:
        return f"{zip_stem}.csv"
    return f"{zip_stem}_{index}.csv"


def _ensure_unique_path(path: Path) -> Path:
    """Return a non-existing path by appending (n) before suffix if needed."""
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent

    i = 2
    while True:
        candidate = parent / f"{stem}({i}){suffix}"
        if not candidate.exists():
            return candidate
        i += 1


def extract_rename_move_csvs(*, base_dir: Path, keep_extracted: bool, overwrite: bool) -> int:
    zips = sorted([p for p in base_dir.iterdir() if p.is_file() and p.suffix.lower() == ".zip"])
    if not zips:
        print("No .zip files found in:", str(base_dir))
        return 0

    moved_count = 0
    for zip_path in zips:
        zip_stem = zip_path.stem
        extract_dir = base_dir / f"{zip_stem}__extracted"

        # Make extraction idempotent-ish: clear old folder unless user asked to keep it.
        if extract_dir.exists() and not keep_extracted:
            shutil.rmtree(extract_dir)
        extract_dir.mkdir(parents=True, exist_ok=True)

        print(f"Processing: {zip_path.name}")

        try:
            with zipfile.ZipFile(zip_path, "r") as zf:
                zf.extractall(extract_dir)
        except zipfile.BadZipFile:
            print(f"  ERROR: Bad zip file, skipping: {zip_path.name}")
            continue

        csv_files = sorted([p for p in extract_dir.rglob("*.csv") if p.is_file()])
        if not csv_files:
            print("  No CSV files found inside extracted contents.")
            if not keep_extracted:
                shutil.rmtree(extract_dir, ignore_errors=True)
            continue

        total = len(csv_files)
        for i, src_csv in enumerate(csv_files, start=1):
            new_name = _safe_dest_name(zip_stem, i, total)
            renamed_src = src_csv.with_name(new_name)

            # Avoid clobbering a CSV within extraction folder (rare but possible)
            if renamed_src.exists() and renamed_src != src_csv:
                renamed_src = _ensure_unique_path(renamed_src)

            if renamed_src != src_csv:
                src_csv.rename(renamed_src)

            dest_path = base_dir / renamed_src.name
            if dest_path.exists():
                if overwrite:
                    dest_path.unlink()
                else:
                    dest_path = _ensure_unique_path(dest_path)

            shutil.move(str(renamed_src), str(dest_path))
            moved_count += 1
            print(f"  Moved: {dest_path.name}")

        if not keep_extracted:
            shutil.rmtree(extract_dir, ignore_errors=True)

    print(f"Done. Moved {moved_count} CSV file(s) to: {base_dir}")
    return 0


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Extract zips, rename CSVs to zip name, move to current folder")
    parser.add_argument("--keep-extracted", action="store_true", help="Keep extracted folders")
    parser.add_argument("--overwrite", action="store_true", help="Overwrite destination CSVs if they already exist")
    args = parser.parse_args(argv)

    base_dir = Path.cwd()
    return extract_rename_move_csvs(base_dir=base_dir, keep_extracted=args.keep_extracted, overwrite=args.overwrite)


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
