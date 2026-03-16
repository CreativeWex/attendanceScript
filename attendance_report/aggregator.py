"""Aggregation logic for turning event streams into daily min/max bounds."""

from collections import defaultdict
from datetime import date, datetime, time
from typing import DefaultDict, Dict, Iterable, List

from .models import DayBounds, EmployeeCalendar, EventRecord


class AttendanceAggregator:
    """Collects employee events and computes per-day arrival/departure bounds."""

    def __init__(self) -> None:
        """Initialize empty in-memory event storage."""
        self._events: DefaultDict[str, DefaultDict[date, List[time]]] = defaultdict(
            lambda: defaultdict(list)
        )

    def add_event(self, event: EventRecord) -> None:
        """Store a single event in internal storage."""
        day = event.occurred_at.date()
        self._events[event.employee_name][day].append(event.occurred_at.time())

    def add_events(self, events: Iterable[EventRecord]) -> None:
        """Store multiple events from any parser source."""
        for event in events:
            self.add_event(event)

    def build_calendar(self) -> EmployeeCalendar:
        """Build an employee -> date -> day bounds structure from collected events."""
        calendar: EmployeeCalendar = {}
        for employee_name, day_map in self._events.items():
            calendar[employee_name] = {}
            for day, times in day_map.items():
                if not times:
                    continue
                calendar[employee_name][day] = DayBounds(
                    arrival_time=min(times),
                    departure_time=max(times),
                )
        return calendar