"""CLI entrypoint for attendance report generation."""

import argparse
from datetime import datetime, time
from pathlib import Path
from typing import Optional, Sequence

from .aggregator import AttendanceAggregator
from .parsers import load_work_mode_mapping, parse_directory
from .report_writer import write_report


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    """Parse CLI arguments for report generation command."""
    parser = argparse.ArgumentParser(
        description="Build attendance report from mixed .xls/.xlsx files."
    )
    parser.add_argument(
        "--input-dir",
        required=True,
        type=Path,
        help="Path to directory with source files.",
    )
    parser.add_argument(
        "--output",
        required=True,
        type=Path,
        help="Path to generated report .xlsx file.",
    )
    parser.add_argument(
        "--official-time",
        default="09:00",
        help="Default official arrival time in HH:MM format (default: 09:00).",
    )
    return parser.parse_args(argv)


def parse_official_time(raw_value: str) -> time:
    """Parse official arrival time value from CLI argument."""
    try:
        return datetime.strptime(raw_value, "%H:%M").time()
    except ValueError as exc:
        raise ValueError(
            f"Invalid --official-time '{raw_value}'. Expected HH:MM format."
        ) from exc


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Run end-to-end flow from parsing source files to writing final report."""
    args = parse_args(argv)

    input_dir: Path = args.input_dir
    output_file: Path = args.output
    official_time = parse_official_time(args.official_time)

    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory does not exist: {input_dir}")

    events, summary = parse_directory(input_dir)
    aggregator = AttendanceAggregator()
    aggregator.add_events(events)
    calendar = aggregator.build_calendar()

    if not calendar:
        raise RuntimeError("No attendance data found in input files.")

    work_mode_by_fio = load_work_mode_mapping(input_dir)

    write_report(
        output_path=output_file,
        calendar=calendar,
        default_official_time=official_time,
        work_mode_by_fio=work_mode_by_fio,
    )

    print(f"Processed files: {len(summary.processed_files)}")
    print(f"Skipped files: {len(summary.skipped_files)}")
    print(f"Total parsed events: {summary.total_events}")
    print(f"Report written to: {output_file}")

    if summary.skipped_files:
        print("\nSkipped file details:")
        for skipped_path, reason in summary.skipped_files:
            print(f"- {skipped_path.name}: {reason}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())