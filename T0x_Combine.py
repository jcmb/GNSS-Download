#!/usr/bin/env python3

import argparse
import os
import re
import shutil
import sys
from collections import defaultdict
from pathlib import Path

VERSION = "1.2"

# YYYYMMDDHHMM — e.g. VRS-EB81__202504251200.T04
T0X_HHMM_PATTERN = re.compile(
    r"^(?P<base>.+)(?P<year_month>\d{6})(?P<day>\d{2})(?P<time>\d{4})\.(?P<ext>T0[24])$",
    re.IGNORECASE,
)

# YYYYMMDDHH — e.g. Cessnock-Base_2026092323.T04
T0X_HH_PATTERN = re.compile(
    r"^(?P<base>.+)(?P<year_month>\d{6})(?P<day>\d{2})(?P<hour>\d{2})\.(?P<ext>T0[24])$",
    re.IGNORECASE,
)


def get_args():
    parser = argparse.ArgumentParser(
        fromfile_prefix_chars="@",
        description="Combine Trimble T02/T04 files by base name and time order.",
    )

    parser.add_argument(
        "Input",
        nargs="+",
        help="T02/T04 files and/or directories to process",
    )
    parser.add_argument(
        "--Recursive",
        "-R",
        help="Search directories recursively for T02/T04 files",
        action="store_true",
    )
    parser.add_argument(
        "--CrossDay",
        "-X",
        help="Combine files across days within the same month (day is the 2 digits before time)",
        action="store_true",
    )
    parser.add_argument(
        "--Clobber",
        "-C",
        help="Overwrite existing combined output files",
        action="store_true",
    )
    parser.add_argument(
        "--DryRun",
        help="Show what would be combined without writing files",
        action="store_true",
    )
    parser.add_argument("--Verbose", "-V", help="Verbose", action="store_true")
    parser.add_argument("--Tell", "-T", help="Show settings", action="store_true")

    args = parser.parse_args()

    if args.Tell:
        sys.stderr.write(f"Version:   {VERSION}\n")
        sys.stderr.write(f"Input:     {args.Input}\n")
        sys.stderr.write(f"Recursive: {args.Recursive}\n")
        sys.stderr.write(f"CrossDay:  {args.CrossDay}\n")
        sys.stderr.write(f"Clobber:   {args.Clobber}\n")
        sys.stderr.write(f"DryRun:    {args.DryRun}\n")
        sys.stderr.write(f"Verbose:   {args.Verbose}\n")

    return vars(args)


def valid_hhmm(hhmm: str) -> bool:
    hour = int(hhmm[:2])
    minute = int(hhmm[2:])
    return 0 <= hour <= 23 and 0 <= minute <= 59


def valid_hour(hour: str) -> bool:
    return hour.isdigit() and 0 <= int(hour) <= 23


def valid_day(day: str) -> bool:
    return day.isdigit() and 1 <= int(day) <= 31


def parse_t0x_filename(name: str, cross_day: bool = False):
    # Prefer the longer HHMM suffix so ...1200 is not treated as hour-only.
    match = T0X_HHMM_PATTERN.match(name)
    if match:
        time = match.group("time")
        day = match.group("day")
        if not valid_hhmm(time) or not valid_day(day):
            return None
        sort_time = time
        label_time = time
    else:
        match = T0X_HH_PATTERN.match(name)
        if not match:
            return None
        hour = match.group("hour")
        day = match.group("day")
        if not valid_hour(hour) or not valid_day(day):
            return None
        # Normalize to HHMM for consistent ordering with mixed groups.
        sort_time = f"{hour}00"
        label_time = hour

    base = match.group("base")
    year_month = match.group("year_month")
    ext = match.group("ext").upper()

    if cross_day:
        group_key = f"{base}{year_month}"
        sort_key = f"{day}{sort_time}"
        label = f"{day} {label_time}"
    else:
        group_key = f"{base}{year_month}{day}"
        sort_key = sort_time
        label = label_time

    return group_key, sort_key, label, ext


def is_t0x_file(path: Path) -> bool:
    return path.suffix.upper() in (".T02", ".T04")


def find_t0x_files(inputs, recursive: bool) -> list[Path]:
    files = []

    for item in inputs:
        path = Path(item)
        if not path.exists():
            sys.stderr.write(f"Warning: path does not exist, skipping: {item}\n")
            continue

        if path.is_file():
            if is_t0x_file(path):
                files.append(path.resolve())
            else:
                sys.stderr.write(
                    f"Warning: not a T02/T04 file, skipping: {path.name}\n"
                )
            continue

        if recursive:
            for root, _dirs, filenames in os.walk(path):
                for filename in filenames:
                    file_path = Path(root) / filename
                    if is_t0x_file(file_path):
                        files.append(file_path.resolve())
        else:
            for file_path in path.iterdir():
                if file_path.is_file() and is_t0x_file(file_path):
                    files.append(file_path.resolve())

    return sorted(set(files))


def group_files(file_paths: list[Path], cross_day: bool, verbose: bool):
    groups = defaultdict(list)
    skipped = 0

    for path in file_paths:
        parsed = parse_t0x_filename(path.name, cross_day)
        if parsed is None:
            skipped += 1
            sys.stderr.write(
                f"Warning: invalid T0x filename, skipping: {path.name}\n"
            )
            continue

        group_key, sort_key, label, ext = parsed
        groups[(path.parent, group_key, ext)].append((sort_key, label, path))

    if verbose and skipped:
        sys.stderr.write(f"Skipped {skipped} file(s) with invalid names.\n")

    return groups


def detect_order_issues(sort_keys: list[str]) -> list[str]:
    issues = []
    seen = set()
    for index, sort_key in enumerate(sort_keys):
        if index > 0 and sort_key < sort_keys[index - 1]:
            issues.append(
                f"{sort_key} is out of order after {sort_keys[index - 1]}"
            )
        if sort_key in seen:
            issues.append(f"duplicate sort key {sort_key}")
        seen.add(sort_key)
    return issues


def format_size(num_bytes: int) -> str:
    if num_bytes < 1024:
        return f"{num_bytes} B"
    if num_bytes < 1024 * 1024:
        return f"{num_bytes / 1024:.1f} KB"
    return f"{num_bytes / (1024 * 1024):.1f} MB"


def combine_group(
    directory: Path,
    group_key: str,
    ext: str,
    members: list[tuple[str, str, Path]],
    clobber: bool,
    dry_run: bool,
    verbose: bool,
):
    ordered = sorted(members, key=lambda item: (item[0], item[2].name))
    output_name = f"{group_key}.{ext}"
    output_path = directory / output_name

    input_paths = [path for _sort_key, _label, path in ordered]
    input_paths_set = set(input_paths)

    if output_path.resolve() in {path.resolve() for path in input_paths_set}:
        sys.stderr.write(
            f"Warning: output file is also an input, skipping group: {output_name}\n"
        )
        return "error"

    if output_path.exists() and not clobber:
        sys.stderr.write(
            f"Error: output exists (use --Clobber to overwrite): {output_path}\n"
        )
        return "error"

    if len(ordered) == 1 and ordered[0][2] == output_path:
        if verbose:
            print(f"Already named: {output_name}")
        return "skipped"

    count_label = "file" if len(ordered) == 1 else "files"
    print(f"Combining {len(ordered)} {count_label} -> {output_name}")

    for index, (_sort_key, label, path) in enumerate(ordered, start=1):
        print(f"  {index}. {path.name}  ({label})")
        if verbose:
            print(f"      {path}")
            print(f"      {format_size(path.stat().st_size)}")

    sort_keys = [sort_key for sort_key, _label, _path in ordered]
    for issue in detect_order_issues(sort_keys):
        sys.stderr.write(f"Warning: {issue} in group {output_name}\n")

    if dry_run:
        print(f"  -> would write {output_path}")
        return "combined"

    total_bytes = 0
    with open(output_path, "wb") as out_file:
        for _sort_key, _label, path in ordered:
            with open(path, "rb") as in_file:
                shutil.copyfileobj(in_file, out_file)
            total_bytes += path.stat().st_size

    print(f"  -> wrote {output_path} ({format_size(total_bytes)})")
    return "combined"


def main():
    args = get_args()
    file_paths = find_t0x_files(args["Input"], args["Recursive"])

    if not file_paths:
        print("No .T02 or .T04 files found.")
        return

    groups = group_files(file_paths, args["CrossDay"], args["Verbose"])

    if not groups:
        print("No valid T0x files found.")
        return

    combined = 0
    errors = 0

    for (directory, group_key, ext), members in sorted(groups.items()):
        result = combine_group(
            directory,
            group_key,
            ext,
            members,
            args["Clobber"],
            args["DryRun"],
            args["Verbose"],
        )
        if result == "combined":
            combined += 1
        elif result == "error":
            errors += 1
        print()

    print(
        f"Summary: {combined} combined, {errors} error(s), "
        f"{len(groups) - combined - errors} skipped"
    )

    if errors:
        sys.exit(1)


if __name__ == "__main__":
    main()
