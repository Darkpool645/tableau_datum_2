from __future__ import annotations

from pathlib import Path
from dataclasses import dataclass
from datetime import date, timedelta

MODE_UI = "ui"
MODE_URL = "url"
DOWNLOAD_DIR = Path("downloads")


@dataclass
class DownloadedDay:
    start: date
    end: date
    file: Path
    already_exists: bool

    @property
    def tag(self) -> str:
        return self.start.strftime("%Y-%m-%d")


def day_chunks(start: date, end: date) -> list[tuple[date, date]]:
    cursor = start
    sections = []
    while cursor <= end:
        sections.append((cursor, cursor))
        cursor += timedelta(days=1)

    return sections


def file_name(start: date, end: date) -> str:
    if start == end:
        return f"reporte_ventas_{start.strftime('%Y-%m-%d')}.xlsx"
    return (f"reporte_ventas_{start.strftime('%Y-%m-%d')}"
            f"_a_{end.strftime('%Y-%m-%d')}.xlsx")
