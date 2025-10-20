from pdf_bulk_filler.main import run_cli


def test_run_cli_version(capsys):
    exit_code = run_cli(["--version"])
    assert exit_code == 0
    captured = capsys.readouterr()
    assert "pdf-bulk-filler 0.1.0" in captured.out


def test_run_cli_with_no_gui(capsys):
    exit_code = run_cli(["--no-gui"])
    assert exit_code == 0
    captured = capsys.readouterr()
    assert "GUI launch suppressed" in captured.out
