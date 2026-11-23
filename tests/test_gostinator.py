from contextlib import contextmanager
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock

import pytest

from gostinator import solidworks
from gostinator.cli import run_batch


def test_batch_publishes_only_pairs_and_reports_file_progress(tmp_path, monkeypatch):
    source = tmp_path / "in"
    source.mkdir()
    for name in ("Вал", "Деталь"):
        (source / (name + ".SLDPRT")).touch()
    (source / "~$Draft.SLDPRT").touch()
    (source / "unrelated.pdf").touch()
    target = tmp_path / "out"
    seen = []
    staging_dirs = []
    app = object()
    template = Path("test.drwdot")

    @contextmanager
    def session():
        yield app

    def generate(connected, part, selected_template, stage):
        assert connected is app
        assert selected_template == template
        staging_dirs.append(stage)
        drawing = stage / (part.stem + ".SLDDRW")
        pdf = drawing.with_suffix(".pdf")
        drawing.write_bytes(b"drawing")
        pdf.write_bytes(b"%PDF")
        return drawing, pdf

    monkeypatch.setattr(solidworks, "session", session)
    monkeypatch.setattr(solidworks, "drawing_template", lambda _: template)
    monkeypatch.setattr(solidworks, "convert_part", generate)
    assert run_batch(source, target, on_progress=lambda *p: seen.append(p)) == 0
    assert seen == [(1, 2), (2, 2)]
    assert all(not folder.exists() for folder in staging_dirs)
    assert {p.name for p in target.iterdir()} == {
        stem + ext for stem in ("Вал", "Деталь") for ext in (".SLDDRW", ".pdf")
    }


def test_missing_pdf_leaves_no_drawing_and_collision_preserves_file(tmp_path, monkeypatch):
    source = tmp_path / "in"
    source.mkdir()
    (source / "part.SLDPRT").touch()
    target = tmp_path / "out"

    @contextmanager
    def session():
        yield object()

    def generate(app, part, template, stage):
        drawing = stage / "temporary.SLDDRW"
        drawing.write_bytes(b"drawing")
        return drawing, drawing.with_suffix(".pdf")

    monkeypatch.setattr(solidworks, "session", session)
    monkeypatch.setattr(solidworks, "drawing_template", lambda app: Path("part.drwdot"))
    monkeypatch.setattr(solidworks, "convert_part", generate)
    assert run_batch(source, target) == 2
    assert list(target.iterdir()) == []
    existing = target / "part.pdf"
    existing.write_bytes(b"original")
    with pytest.raises(FileExistsError):
        run_batch(source, target)
    assert existing.read_bytes() == b"original"


def test_save_uses_legacy_method_when_solidworks_rejects_saveas(tmp_path):
    def unsupported(*args):
        raise solidworks.com_error(-2147352571, "Type mismatch", None, 4)

    def legacy(path, version, options):
        Path(path).write_bytes(b"%PDF")
        return 0

    drawing = SimpleNamespace(Extension=SimpleNamespace(SaveAs=unsupported), SaveAs3=legacy)
    pdf = tmp_path / "part.pdf"
    solidworks._save(drawing, pdf)
    assert pdf.read_bytes() == b"%PDF"


@pytest.mark.parametrize("first_angle,already_open", [(True, False), (False, True)])
def test_convert_uses_native_projection_and_keeps_user_document_open(
        tmp_path, first_angle, already_open, monkeypatch):
    from gostinator import sheet_layout

    monkeypatch.setattr(sheet_layout, "installed_formats", lambda _: {})
    part = tmp_path / "part.SLDPRT"
    part.touch()
    sheet = SimpleNamespace(GetProperties2=lambda: (0, 0, 1, 1, first_angle, .21, .297))

    def save(path, version, flags, export_data, errors, warnings):
        Path(path).write_bytes(b"export")
        return True

    drawing = SimpleNamespace(
        GetCurrentSheet=lambda: sheet, GetTitle=lambda: "drawing",
        GetViews=lambda: (),
        Create1stAngleViews2=Mock(spec=lambda path: None, return_value=True),
        Create3rdAngleViews2=Mock(spec=lambda path: None, return_value=True),
        InsertModelAnnotations3=Mock(spec=lambda *args: None),
        ClearSelection2=Mock(spec=lambda value: None),
        Extension=SimpleNamespace(SaveAs=Mock(spec=save, side_effect=save)),
    )
    model = SimpleNamespace(GetPathName=lambda: str(part), GetTitle=lambda: "part",
                            GetBodies2=lambda *_: ())
    spec = SimpleNamespace()
    app = SimpleNamespace(
        GetOpenDocumentByName=lambda _: model if already_open else None,
        GetOpenDocSpec=Mock(spec=lambda: None, return_value=spec),
        OpenDoc7=Mock(spec=lambda: None, return_value=model),
        NewDocument=Mock(spec=lambda: None, return_value=drawing), CloseDoc=Mock(spec=lambda: None),
    )
    paths = solidworks.convert_part(app, part, Path("example.drwdot"), tmp_path)
    assert all(path.exists() for path in paths)
    assert [Path(args.args[0]).suffix for args in drawing.Extension.SaveAs.call_args_list] == [".SLDDRW", ".pdf"]
    expected = drawing.Create1stAngleViews2 if first_angle else drawing.Create3rdAngleViews2
    expected.assert_called_once_with(str(part))
    drawing.InsertModelAnnotations3.assert_called_once_with(0, 32768, True, False, False, False)
    assert [args.args[0] for args in app.CloseDoc.call_args_list] == (
        ["drawing"] if already_open else ["drawing", "part"])
    if already_open:
        app.OpenDoc7.assert_not_called()
    else:
        assert spec.Silent and spec.ReadOnly
