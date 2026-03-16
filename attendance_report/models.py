
"""Data models shared between parser, aggregator and report writer."""

from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from typing import Dict


@dataclass(frozen=True)
class EventRecord:
    """Represents one raw attendance event for a specific employee."""

    employee_name: str
    occurred_at: datetime


@dataclass(frozen=True)
class DayBounds:
    """Stores daily attendance metrics for one calendar day.

    - arrival_time: earliest arrival for the day
    - departure_time: latest departure for the day
    - absence_duration: total time spent outside office between first arrival
      and final departure (sum of all exit -> next entry intervals).
    """

    arrival_time: time
    departure_time: time
    absence_duration: timedelta


EmployeeCalendar = Dict[str, Dict[date, DayBounds]]