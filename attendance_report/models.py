
"""Data models shared between parser, aggregator and report writer."""

from dataclasses import dataclass
from datetime import date, datetime, time
from typing import Dict


@dataclass(frozen=True)
class EventRecord:
    """Represents one raw attendance event for a specific employee."""

    employee_name: str
    occurred_at: datetime


@dataclass(frozen=True)
class DayBounds:
    """Stores earliest arrival and latest departure for one calendar day."""

    arrival_time: time
    departure_time: time


EmployeeCalendar = Dict[str, Dict[date, DayBounds]]