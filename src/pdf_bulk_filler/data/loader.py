"""Utilities for loading tabular data sources into pandas data frames."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional

import pandas as pd


@dataclass
class DataSample:
    """Container describing a loaded dataset."""

    source_path: Path
    dataframe: pd.DataFrame
    sheet_name: Optional[str] = None
    available_sheets: list[str] = field(default_factory=list)
    header_row: Optional[int] = None
    data_row: Optional[int] = None
    column_offset: int = 0

    def columns(self) -> list[str]:
        """Return the ordered list of column names."""
        return list(self.dataframe.columns)

    def head_records(self, rows: int = 20) -> pd.DataFrame:
        """Return the top `rows` rows for previewing."""
        return self.dataframe.head(rows)


class DataLoader:
    """Load CSV and Excel files using pandas/openpyxl."""

    SUPPORTED_SUFFIXES: tuple[str, ...] = (".csv", ".tsv", ".xls", ".xlsx")

    def load(
        self,
        path: Path,
        *,
        sheet: Optional[str] = None,
        header_row: Optional[int] = None,
        data_row: Optional[int] = None,
        column_offset: int = 0,
    ) -> DataSample:
        """Load a tabular file and return a :class:`DataSample`."""
        normalized = path.expanduser().resolve()
        if not normalized.exists():
            raise FileNotFoundError(f"Data source not found: {normalized}")

        if normalized.suffix.lower() not in self.SUPPORTED_SUFFIXES:
            allowed = ", ".join(self.SUPPORTED_SUFFIXES)
            raise ValueError(f"Unsupported file type {normalized.suffix!r}. Allowed: {allowed}")

        available_sheets: list[str] = []
        sheet_name: Optional[str] = None

        if normalized.suffix.lower() == ".csv":
            frame = pd.read_csv(normalized, header=None)
        elif normalized.suffix.lower() == ".tsv":
            frame = pd.read_csv(normalized, sep="\t", header=None)
        else:
            with pd.ExcelFile(normalized, engine="openpyxl") as workbook:
                available_sheets = list(workbook.sheet_names)
                if not available_sheets:
                    raise ValueError(f"Workbook '{normalized.name}' contains no worksheets.")
                if sheet is None:
                    sheet_name = available_sheets[0]
                else:
                    if sheet not in available_sheets:
                        available = ", ".join(available_sheets)
                        raise ValueError(
                            f"Worksheet '{sheet}' not found. Available sheets: {available}"
                        )
                    sheet_name = sheet
                frame = workbook.parse(sheet_name, header=None)

        total_rows, total_cols = frame.shape
        col_offset = max(0, column_offset)
        if col_offset >= total_cols:
            raise ValueError("Column offset exceeds available columns.")
        if col_offset:
            frame = frame.iloc[:, col_offset:]

        header_idx = header_row - 1 if header_row and header_row > 0 else 0
        data_idx = data_row - 1 if data_row and data_row > 0 else header_idx + 1
        header_idx = max(0, header_idx)
        data_idx = max(header_idx + 1, data_idx)

        if header_idx >= len(frame.index):
            raise ValueError("Header row index exceeds available rows.")
        if data_idx > len(frame.index):
            raise ValueError("Data start row exceeds available rows.")

        header_values = frame.iloc[header_idx]
        cleaned_columns = [
            self._normalize_column(value) or f"Column {idx + 1}"
            for idx, value in enumerate(header_values)
        ]
        unique_columns = self._deduplicate_columns(cleaned_columns)

        frame = frame.iloc[data_idx:]
        if frame.empty:
            raise ValueError("Loaded dataset is empty.")
        frame = frame.reset_index(drop=True)
        frame.columns = unique_columns

        return DataSample(
            source_path=normalized,
            dataframe=frame,
            sheet_name=sheet_name,
            available_sheets=available_sheets,
            header_row=header_idx + 1,
            data_row=data_idx + 1,
            column_offset=col_offset,
        )

    @staticmethod
    def _normalize_column(column: str) -> str:
        """Trim whitespace and collapse repeated spaces in column names."""
        collapsed = " ".join(str(column).split())
        return collapsed.strip()

    @staticmethod
    def _deduplicate_columns(columns: list[str]) -> list[str]:
        """Append numeric suffixes when column names repeat."""
        seen: dict[str, int] = {}
        unique: list[str] = []
        for name in columns:
            count = seen.get(name, 0)
            seen[name] = count + 1
            if count == 0:
                unique.append(name)
            else:
                unique.append(f"{name} ({count + 1})")
        return unique

    @classmethod
    def supported_filters(cls) -> Iterable[str]:
        """Return file dialog filters for supported formats."""
        return (
            "All Supported (*.csv *.tsv *.xls *.xlsx)",
            "CSV Files (*.csv)",
            "Excel Files (*.xls *.xlsx)",
            "Delimited Text (*.tsv)",
        )

