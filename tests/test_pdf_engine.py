from pathlib import Path

import fitz
from PyPDF2 import PdfReader

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
