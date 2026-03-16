"""Aggregation logic for turning event streams into daily min/max bounds."""

from collections import defaultdict
from datetime import date, datetime, time, timedelta
from typing import DefaultDict, Iterable, List

from .models import DayBounds, EmployeeCalendar, EventRecord


class AttendanceAggregator:
    """Collects employee events and computes per-day attendance metrics."""

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
        """Build an employee -> date -> day bounds structure from collected events.

        For each day we assume the event times represent alternating
        вход/выход сотрудника, начиная с входа. Тогда периоды отсутствия
        внутри дня — это интервалы между выходом и следующим входом.
        Последний выход без последующего входа в рамках того же дня
        в расчёт «отсутствия в течение дня» не включается.
        """
        calendar: EmployeeCalendar = {}
        for employee_name, day_map in self._events.items():
            calendar[employee_name] = {}
            for day, times in day_map.items():
                if not times:
                    continue

                # Сортируем события по времени в течение дня
                day_times = sorted(times)
                arrival_time = day_times[0]
                departure_time = day_times[-1]

                # Считаем суммарное отсутствие: сумма интервалов (выход -> следующий вход)
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