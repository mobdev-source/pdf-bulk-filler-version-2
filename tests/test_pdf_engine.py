from pathlib import Path

import fitz
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import (
    ArrayObject,
    BooleanObject,
    DictionaryObject,
    FloatObject,
    NameObject,
    NumberObject,
    TextStringObject,
)

from pdf_bulk_filler.pdf.engine import PdfEngine


def test_fill_rows_flatten(tmp_path):
    engine = PdfEngine()
    template_path = Path("assets/templates/sample_invoice.pdf").resolve()
    template = engine.open_template(template_path)

    mapping = {
        "full_name": "FullName",
        "email": "Email",
        "amount_due": "AmountDue",
    }
    rows = [
        {"FullName": "Alice Example", "Email": "alice@example.com", "AmountDue": "42.50", "id": "1001"},
    ]

    outputs = engine.fill_rows(
        template.path,
        tmp_path,
        mapping,
        rows,
        flatten=True,
        template_metadata=template,
    )

    assert len(outputs) == 1
    output_path = outputs[0]
    assert output_path.exists()

    reader = PdfReader(str(output_path))
    assert "/AcroForm" not in reader.trailer["/Root"]

    doc = fitz.open(output_path)
    try:
        page = doc.load_page(0)
        widgets = list(page.widgets() or [])
        assert widgets == []
        text = page.get_text()
        assert "Alice Example" in text
        assert "alice@example.com" in text
        assert "42.50" in text
    finally:
        doc.close()
        template.close()


def _create_checkbox_template(path: Path) -> None:
    writer = PdfWriter()
    writer.add_blank_page(width=200, height=200)
    page = writer.pages[0]

    checkbox = DictionaryObject()
    checkbox.update(
        {
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/FT"): NameObject("/Btn"),
            NameObject("/Ff"): NumberObject(0),
            NameObject("/Rect"): ArrayObject(
                [FloatObject(50), FloatObject(140), FloatObject(70), FloatObject(160)]
            ),
            NameObject("/T"): TextStringObject("Agree"),
            NameObject("/V"): NameObject("/Off"),
            NameObject("/AS"): NameObject("/Off"),
        }
    )
    appearance_states = DictionaryObject()
    appearance_states[NameObject("/Off")] = NameObject("/Off")
    appearance_states[NameObject("/Yes")] = NameObject("/Yes")
    checkbox[NameObject("/AP")] = DictionaryObject({NameObject("/N"): appearance_states})

    checkbox_ref = writer._add_object(checkbox)
    page[NameObject("/Annots")] = ArrayObject([checkbox_ref])

    acro_form = DictionaryObject()
    acro_form[NameObject("/Fields")] = ArrayObject([checkbox_ref])
    acro_form[NameObject("/NeedAppearances")] = BooleanObject(True)
    writer._root_object.update({NameObject("/AcroForm"): acro_form})

    with path.open("wb") as handle:
        writer.write(handle)


def test_fill_rows_checkbox_states(tmp_path):
    template_path = tmp_path / "checkbox.pdf"
    _create_checkbox_template(template_path)

    engine = PdfEngine()
    rows = [{"Agree": "/Yes"}]
    outputs = engine.fill_rows(
        template_path,
        tmp_path,
        {"Agree": "Agree"},
        rows,
        flatten=False,
        read_only=True,
    )

    output_path = outputs[0]
    reader = PdfReader(str(output_path))

    annotation_ref = reader.pages[0]["/Annots"][0]
    annotation = annotation_ref.get_object()
    assert str(annotation.get("/V")) == "/Yes"
    assert str(annotation.get("/AS")) == "/Yes"
    flags = int(annotation.get("/Ff", 0))
    assert flags & 1 == 1

    acro_form = reader.trailer["/Root"].get("/AcroForm")
    if acro_form is not None and hasattr(acro_form, "get_object"):
        acro_form = acro_form.get_object()
    if acro_form is not None:
        assert bool(acro_form.get("/NeedAppearances", False))
