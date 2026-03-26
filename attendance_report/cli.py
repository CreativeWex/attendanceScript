"""CLI entrypoint for attendance report generation."""

import argparse
from datetime import date, datetime, time, timedelta
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

from .aggregator import AttendanceAggregator
from .parsers import load_work_mode_mapping, parse_directory
from .report_writer import write_report

OFFICE_WORK_MODE = "офисный труд"


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

    _print_top_work_duration(calendar, work_mode_by_fio=work_mode_by_fio, top_n=30)
    _print_top_arrival_delta(
        calendar,
        official_time=official_time,
        work_mode_by_fio=work_mode_by_fio,
        top_n=20,
    )

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


def _compute_day_work_duration(day: date, arrival: time, departure: time) -> timedelta:
    start_dt = datetime.combine(day, arrival)
    end_dt = datetime.combine(day, departure)
    if end_dt < start_dt:
        end_dt += timedelta(days=1)
    return end_dt - start_dt


def _format_timedelta_hhmm(value: timedelta) -> str:
    total_seconds = int(value.total_seconds())
    sign = "-" if total_seconds < 0 else ""
    total_seconds = abs(total_seconds)
    total_minutes = total_seconds // 60
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{sign}{hours}:{minutes:02d}"

def _compute_arrival_delta(day: date, arrival: time, official_time: time) -> timedelta:
    """Compute arrival delta as фактический приход - плановый приход."""
    arrival_dt = datetime.combine(day, arrival)
    official_dt = datetime.combine(day, official_time)
    return arrival_dt - official_dt


def _print_top_work_duration(
    calendar,
    work_mode_by_fio: Optional[Dict[str, str]],
    top_n: int = 20,
) -> None:
    averages: Dict[str, timedelta] = {}
    for employee_name, days in calendar.items():
        if work_mode_by_fio is None:
            continue
        mode = work_mode_by_fio.get(employee_name)
        if mode is None or mode.strip().lower() != OFFICE_WORK_MODE:
            continue

        total = timedelta(0)
        day_count = 0
        for day, bounds in days.items():
            # Match report logic: averages are computed only for weekdays.
            if day.weekday() >= 5:
                continue
            duration = _compute_day_work_duration(
                day, bounds.arrival_time, bounds.departure_time
            )
            total += duration
            day_count += 1
        if day_count:
            averages[employee_name] = total / day_count

    if work_mode_by_fio is None:
        print(
            "\nТоп-20 по времени работы не построен: не найден файл режима работы "
            "(нужно учитывать только режим 'офисный труд')."
        )
        return

    if not averages:
        print(
            f"\nТоп-20 по времени работы не построен: не найдено сотрудников с режимом "
            f"'{OFFICE_WORK_MODE}'."
        )
        return

    ranked: List[Tuple[str, timedelta]] = sorted(
        averages.items(), key=lambda item: item[1]
    )

    print("\nTop-30 сотрудников с наибольшим временем работы (Среднее время работы):")
    for idx, (name, total) in enumerate(reversed(ranked[-top_n:]), start=1):
        print(f"{idx:>2}. {name}: {_format_timedelta_hhmm(total)}")

    print("\nTop-30 сотрудников с наименьшим временем работы (Среднее время работы):")
    for idx, (name, total) in enumerate(ranked[:top_n], start=1):
        print(f"{idx:>2}. {name}: {_format_timedelta_hhmm(total)}")


def _print_top_arrival_delta(
    calendar,
    official_time: time,
    work_mode_by_fio: Optional[Dict[str, str]],
    top_n: int = 20,
) -> None:
    """Print top employees with the smallest arrival delta (weekday only)."""
    if work_mode_by_fio is None:
        print(
            "\nТоп-20 по отклонению прихода не построен: не найден файл режима работы "
            "(нужно учитывать только режим 'офисный труд')."
        )
        return

    averages: Dict[str, timedelta] = {}
    for employee_name, days in calendar.items():
        mode = work_mode_by_fio.get(employee_name)
        if mode is None or mode.strip().lower() != OFFICE_WORK_MODE:
            continue

        total = timedelta(0)
        day_count = 0
        for day, bounds in days.items():
            # Match report logic: averages are computed only for weekdays.
            if day.weekday() >= 5:
                continue
            total += _compute_arrival_delta(day, bounds.arrival_time, official_time)
            day_count += 1

        if day_count:
            averages[employee_name] = total / day_count

    if not averages:
        print(
            f"\nТоп-20 по отклонению прихода не построен: не найдено сотрудников с режимом "
            f"'{OFFICE_WORK_MODE}'."
        )
        return

    ranked = sorted(averages.items(), key=lambda item: item[1])
    print(
        "\nTop-20 сотрудников с наименьшим отклонением по времени прихода "
        "(Среднее отклонение прихода):"
    )
    for idx, (name, delta) in enumerate(ranked[:top_n], start=1):
        print(f"{idx:>2}. {name}: {_format_timedelta_hhmm(delta)}")


if __name__ == "__main__":
    raise SystemExit(main())