from __future__ import annotations

import argparse
import csv
import datetime as _dt
import os
import shutil
from pathlib import Path
from typing import Iterable, List, Optional, Tuple


def _normalize_cell(cell: str) -> str:
    """Normalize a cell for matching.

    Handles common Excel-export patterns like: ="主要人员 2" or "  ".
    """
    value = cell.strip()
    if value.startswith("="):
        value = value[1:].lstrip()

    # Strip surrounding quotes if present.
    if len(value) >= 2 and value[0] == '"' and value[-1] == '"':
        value = value[1:-1]

    return value.strip()


def _unwrap_excel_text_formula(cell: str) -> str:
    """Convert common Excel-export 'text as formula' patterns into plain text.

    Excel (and some export tools) will emit text as formulas like: ="00123" so
    that leading zeros are preserved. When saved back to CSV, we want the actual
    value (00123) rather than the formula string.

    Only unwraps the very specific pattern ="..." (or ='...') to avoid
    accidentally changing genuine formulas.
    """

    raw = cell.strip()
    if not raw.startswith("="):
        return cell

    expr = raw[1:].lstrip()
    if len(expr) >= 2 and ((expr[0] == '"' and expr[-1] == '"') or (expr[0] == "'" and expr[-1] == "'")):
        return expr[1:-1]

    return cell


def _parse_csv_line(line: str) -> Optional[List[str]]:
    """Parse a single CSV line into cells.

    Returns None if parsing fails (in which case caller may keep the line).
    """
    try:
        # newline='' is important for csv module; using a list with one line is fine.
        reader = csv.reader([line], delimiter=",", quotechar='"')
        return next(reader)
    except Exception:
        return None


def _row_is_effectively_empty(cells: Iterable[str]) -> bool:
    for cell in cells:
        if _normalize_cell(cell) != "":
            return False
    return True


def _first_nonempty_cell(cells: Iterable[str]) -> str:
    for cell in cells:
        normalized = _normalize_cell(cell)
        if normalized != "":
            return normalized
    return ""


def _should_drop_line(line: str, *, drop_prefix: str) -> Tuple[bool, str]:
    """Return (drop?, reason)."""
    if line.strip() == "":
        return True, "blank-line"

    cells = _parse_csv_line(line)
    if cells is None:
        # If the line can't be parsed, keep it rather than risking data loss.
        return False, "unparsed"

    if _row_is_effectively_empty(cells):
        return True, "empty-row"

    first_value = _first_nonempty_cell(cells)
    if first_value.startswith(drop_prefix):
        return True, f"prefix:{drop_prefix}"

    return False, "keep"


def _iter_target_csv_files(folder: Path) -> List[Path]:
    return sorted(
        [p for p in folder.glob("*.csv") if p.is_file()],
        key=lambda p: p.name,
    )


def clean_one_file(
    csv_path: Path,
    *,
    drop_prefix: str,
    encoding: str,
) -> Tuple[int, int]:
    """Clean a single CSV file in-place.

    Returns (kept_lines, dropped_lines).
    """
    kept = 0
    dropped = 0

    tmp_path = csv_path.with_suffix(csv_path.suffix + ".tmp")

    with csv_path.open("r", encoding=encoding, errors="replace", newline="") as fin, tmp_path.open(
        "w", encoding=encoding, newline=""
    ) as fout:
        writer = csv.writer(fout, delimiter=",", quotechar='"')
        for line_number, line in enumerate(fin, start=1):
            if line_number == 1:
                dropped += 1
                continue
            drop, _reason = _should_drop_line(line, drop_prefix=drop_prefix)
            if drop:
                dropped += 1
                continue

            cells = _parse_csv_line(line)
            if cells is None:
                # Preserve original line if parsing fails.
                fout.write(line)
            else:
                cleaned_cells = [_unwrap_excel_text_formula(c) for c in cells]
                writer.writerow(cleaned_cells)
            kept += 1

    os.replace(tmp_path, csv_path)
    return kept, dropped


def _choose_encoding_for_file(path: Path) -> str:
    """Try a few common encodings for Chinese CSV exports."""
    candidates = ["utf-8-sig", "utf-8", "gb18030", "gbk"]
    for enc in candidates:
        try:
            with path.open("r", encoding=enc, errors="strict") as f:
                f.read(8192)
            return enc
        except UnicodeDecodeError:
            continue
    # Fall back to a forgiving read.
    return "utf-8-sig"


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Remove empty rows and rows whose first non-empty cell starts with a prefix "
            "(default: 主要人员) from all .csv files in a folder. Also removes the first row "
            "(first line) of each file."
        )
    )
    parser.add_argument(
        "--folder",
        default=".",
        help="Folder containing CSV files (default: current directory).",
    )
    parser.add_argument(
        "--prefix",
        default="主要人员",
        help="Drop rows whose first non-empty cell starts with this prefix (default: 主要人员).",
    )
    parser.add_argument(
        "--backup-dir",
        default="backup",
        help="Backup directory name/path (default: backup).",
    )
    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Do not create a backup before overwriting files.",
    )
    parser.add_argument(
        "--encoding",
        default="auto",
        help="File encoding: auto | utf-8-sig | utf-8 | gb18030 | gbk (default: auto).",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would change without modifying files.",
    )

    args = parser.parse_args()

    folder = Path(args.folder).resolve()
    if not folder.exists() or not folder.is_dir():
        raise SystemExit(f"Folder not found: {folder}")

    csv_files = _iter_target_csv_files(folder)
    if not csv_files:
        print(f"No CSV files found in: {folder}")
        return 0

    timestamp = _dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_root = Path(args.backup_dir)
    backup_dir = (folder / backup_root / timestamp).resolve()

    total_kept = 0
    total_dropped = 0

    for csv_path in csv_files:
        if csv_path.parent.name == backup_root.name:
            # Only relevant if user points --folder at backup itself; otherwise harmless.
            continue

        encoding = args.encoding
        if encoding == "auto":
            encoding = _choose_encoding_for_file(csv_path)

        if args.dry_run:
            kept = 0
            dropped = 0
            with csv_path.open("r", encoding=encoding, errors="replace", newline="") as fin:
                for line_number, line in enumerate(fin, start=1):
                    if line_number == 1:
                        dropped += 1
                        continue
                    drop, _reason = _should_drop_line(line, drop_prefix=args.prefix)
                    if drop:
                        dropped += 1
                    else:
                        kept += 1
            print(f"[DRY RUN] {csv_path.name}: keep={kept}, drop={dropped}, encoding={encoding}")
            total_kept += kept
            total_dropped += dropped
            continue

        if not args.no_backup:
            backup_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(csv_path, backup_dir / csv_path.name)

        kept, dropped = clean_one_file(csv_path, drop_prefix=args.prefix, encoding=encoding)
        print(f"{csv_path.name}: keep={kept}, drop={dropped}, encoding={encoding}")
        total_kept += kept
        total_dropped += dropped

    print(f"Done. Total keep={total_kept}, total drop={total_dropped}.")
    if not args.no_backup and not args.dry_run:
        print(f"Backups saved to: {backup_dir}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
