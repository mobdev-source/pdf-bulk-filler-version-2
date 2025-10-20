from pathlib import Path

from pdf_bulk_filler.mapping.manager import MappingManager, MappingModel


def test_mapping_roundtrip(tmp_path):
    mapping = MappingModel(
        source_data=Path("/data/source.csv"),
        pdf_template=Path("/templates/form.pdf"),
        data_sheet="Contacts",
        header_row=4,
        data_row=5,
        column_offset=3,
        assignments={"FieldA": "column_a", "FieldB": "column_b"},
    )
    manager = MappingManager()

    destination = tmp_path / "mapping.json"
    manager.save(destination, mapping)

    loaded = manager.load(destination)
    assert loaded.source_data == mapping.source_data
    assert loaded.pdf_template == mapping.pdf_template
    assert loaded.data_sheet == mapping.data_sheet
    assert loaded.header_row == mapping.header_row
    assert loaded.data_row == mapping.data_row
    assert loaded.column_offset == mapping.column_offset
    assert loaded.assignments == mapping.assignments
