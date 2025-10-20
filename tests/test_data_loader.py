from pathlib import Path

import pandas as pd
import pytest

from pdf_bulk_filler.data.loader import DataLoader


def test_load_csv_normalizes_columns(tmp_path):
    csv_path = tmp_path / "sample.csv"
    csv_path.write_text("Full Name ,Age\nAlice,30\nBob,31\n", encoding="utf-8")

    loader = DataLoader()
    sample = loader.load(csv_path)

    assert sample.source_path == csv_path.resolve()
    assert sample.columns() == ["Full Name", "Age"]
    assert isinstance(sample.dataframe, pd.DataFrame)
    assert len(sample.dataframe) == 2
    assert sample.sheet_name is None
    assert sample.available_sheets == []
    assert sample.header_row == 1
    assert sample.data_row == 2
    assert sample.column_offset == 0


def test_loader_rejects_missing_file(tmp_path):
    loader = DataLoader()
    with pytest.raises(FileNotFoundError):
        loader.load(tmp_path / "unknown.csv")


def test_load_excel_returns_sheet_metadata(tmp_path):
    path = tmp_path / "multi.xlsx"
    contacts = pd.DataFrame({"ContactID": [1, 2], "Name": ["A", "B"]})
    invoices = pd.DataFrame({"ContactID": [1, 2], "Total": [100, 200]})
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        contacts.to_excel(writer, sheet_name="Contacts", index=False)
        invoices.to_excel(writer, sheet_name="Invoices", index=False)

    loader = DataLoader()
    sample = loader.load(path)
    assert sample.sheet_name == "Contacts"
    assert set(sample.available_sheets) == {"Contacts", "Invoices"}
    assert sample.header_row == 1
    assert sample.data_row == 2
    assert sample.column_offset == 0

    invoices_sample = loader.load(path, sheet="Invoices")
    assert invoices_sample.sheet_name == "Invoices"
    assert invoices_sample.available_sheets == sample.available_sheets
    assert invoices_sample.dataframe.iloc[0]["Total"] == 100


def test_load_with_custom_offsets(tmp_path):
    csv_path = tmp_path / "offset.csv"
    csv_path.write_text(
        "\n".join(
            [
                "meta1,meta2,meta3,meta4,meta5",
                "x,x,x,x,x",
                ",,,,",
                ",,,Name,Score",
                ",,,Alice,10",
                ",,,Bob,20",
            ]
        ),
        encoding="utf-8",
    )

    loader = DataLoader()
    sample = loader.load(csv_path, header_row=4, data_row=5, column_offset=3)

    assert sample.columns() == ["Name", "Score"]
    assert sample.header_row == 4
    assert sample.data_row == 5
    assert sample.column_offset == 3
    assert list(sample.dataframe["Name"]) == ["Alice", "Bob"]


def test_duplicate_headers_get_suffixes(tmp_path):
    csv_path = tmp_path / "dupe_headers.csv"
    csv_path.write_text(
        "Name,Name,Name\nAlice,Alicia,Aly\nBob,Robert,Rob\n",
        encoding="utf-8",
    )

    loader = DataLoader()
    sample = loader.load(csv_path)

    assert sample.columns() == ["Name", "Name (2)", "Name (3)"]
    assert list(sample.dataframe["Name (2)"]) == ["Alicia", "Robert"]
