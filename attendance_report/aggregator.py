"""Aggregation logic for turning event streams into daily min/max bounds."""

from collections import defaultdict
from datetime import date, datetime, time, timedelta
from typing import DefaultDict, Iterable, List

from .models import DayBounds, EmployeeCalendar, EventRecord


class AttendanceAggregator:
    """Collects employee events and computes per-day attendance metrics."""

    def __init__(self) -> None:
        """Initialize empty in-memory event storage."""
        self._events: DefaultDict[str, DefaultDict[date, List[EventRecord]]] = (
            defaultdict(lambda: defaultdict(list))
        )

    def add_event(self, event: EventRecord) -> None:
        """Store a single event in internal storage."""
        day = event.occurred_at.date()
        self._events[event.employee_name][day].append(event)

    def add_events(self, events: Iterable[EventRecord]) -> None:
        """Store multiple events from any parser source."""
        for event in events:
            self.add_event(event)

    def build_calendar(self) -> EmployeeCalendar:
        """Build an employee -> date -> day bounds structure from collected events.

        If direction information (`is_entry`) is available for a day, use it
        to determine absence intervals as (OUT -> next IN). Otherwise, fall
        back to the legacy alternating assumption: вход, выход, вход, выход...
        """
        calendar: EmployeeCalendar = {}
        for employee_name, day_map in self._events.items():
            calendar[employee_name] = {}
            for day, events in day_map.items():
                if not events:
                    continue

                # Sort events by time within the day
                day_events = sorted(events, key=lambda e: e.occurred_at.time())
                times = [e.occurred_at.time() for e in day_events]

                has_direction = any(e.is_entry is not None for e in day_events)

                if has_direction:
                    # Arrival = first known entry, or earliest time if none
                    entry_times = [
                        e.occurred_at.time()
                        for e in day_events
                        if e.is_entry is True
                    ]
                    exit_times = [
                        e.occurred_at.time()
                        for e in day_events
                        if e.is_entry is False
                    ]

                    arrival_time = min(entry_times) if entry_times else min(times)
                    departure_time = max(exit_times) if exit_times else max(times)

                    total_absence = timedelta(0)
                    last_exit: datetime | None = None
                    for event in day_events:
                        current_dt = event.occurred_at
                        if event.is_entry is False:
                            # Remember last exit time
                            last_exit = current_dt
                        elif event.is_entry is True and last_exit is not None:
                            # Entry after the last exit -> absence interval
                            if current_dt > last_exit:
                                total_absence += current_dt - last_exit
                            last_exit = None
                else:
                    # Legacy behaviour: assume alternating вход/выход, starting with вход
                    day_times = sorted(times)
                    arrival_time = day_times[0]
                    departure_time = day_times[-1]

                    total_absence = timedelta(0)
                    for idx in range(1, len(day_times) - 1, 2):
                        out_time = day_times[idx]
                        in_time = day_times[idx + 1]
                        out_dt = datetime.combine(day, out_time)
                        in_dt = datetime.combine(day, in_time)
                        if in_dt > out_dt:
                            total_absence += in_dt - out_dt

                calendar[employee_name][day] = DayBounds(
                    arrival_time=arrival_time,
                    departure_time=departure_time,
                    absence_duration=total_absence,
                )
        return calendar