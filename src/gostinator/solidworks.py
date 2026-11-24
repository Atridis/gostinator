"""Create a drawing and export a PDF using the installed SolidWorks."""
from contextlib import contextmanager
import os
from pathlib import Path
from shutil import copy2
from uuid import uuid4

from pythoncom import CoInitialize, CoUninitialize, IID_IDispatch, TypeIIDs, VT_BYREF, VT_I4, com_error
from win32com.client import Dispatch, VARIANT


SW_DOC_PART = 1
SW_DOC_DRAWING = 3
SW_DEFAULT_TEMPLATE_DRAWING = 10
SW_FILE_LOCATIONS_DOCUMENT_TEMPLATES = 6
SW_SAVE_CURRENT_VERSION = 0
SW_SAVE_SILENT = 1
SW_IMPORT_FROM_ENTIRE_MODEL = 0
SW_DIMENSIONS_MARKED_FOR_DRAWING = 32768
SW_DISPLAY_ORIGINS = 6
SW_DISPLAY_COORD_SYSTEMS = 13
_RAW_DISPATCH = TypeIIDs[IID_IDispatch]


def as_dispatch(value):
    # SAFEARRAYs return raw PyIDispatch entries on SolidWorks 2017.
    return Dispatch(value) if isinstance(value, _RAW_DISPATCH) else value


def get_value(obj, name, *args):
    member = getattr(as_dispatch(obj), name)
    # A COM object returned by a property also looks callable to Python.
    value = member(*args) if callable(member) and not hasattr(member, "_oleobj_") else member
    return as_dispatch(value)


@contextmanager
def session():
    CoInitialize()
    try:
        try:
            app = Dispatch("SldWorks.Application")
        except com_error as exc:
            raise RuntimeError(f"Не удалось подключиться к SolidWorks — {exc}") from exc
        yield app
    finally:
        CoUninitialize()


def drawing_template(app):
    configured = get_value(app, "GetUserPreferenceStringValue", SW_DEFAULT_TEMPLATE_DRAWING)
    if configured:
        path = Path(configured.strip().strip('"'))
        if path.suffix.lower() == ".drwdot" and path.is_file():
            return path

    # Some installations ship a template before a default has been assigned.
    locations = get_value(app, "GetUserPreferenceStringValue", SW_FILE_LOCATIONS_DOCUMENT_TEMPLATES)
    folders = [Path(os.path.expandvars(value.strip().strip('"')))
               for value in (locations or "").split(";") if value.strip()]
    if program_data := os.environ.get("PROGRAMDATA"):
        root = Path(program_data) / "SOLIDWORKS"
        if root.is_dir():
            folders.extend(release / "templates"
                           for release in sorted(root.glob("SOLIDWORKS 20[0-9][0-9]"), reverse=True))
    for folder in folders:
        for name in ("gost-part drw.drwdot", "Drawing.drwdot", "Чертеж.drwdot", "Чертёж.drwdot"):
            path = folder / name
            if path.is_file():
                return path
    raise FileNotFoundError("Укажите шаблон чертежа .drwdot в настройках SolidWorks "
                            "«Шаблоны по умолчанию»")


def _same_path(doc, path):
    opened = get_value(doc, "GetPathName")
    return bool(opened) and Path(opened).resolve() == path


@contextmanager
def _open_document(app, path, *, read_only=True):
    existing = get_value(app, "GetOpenDocumentByName", str(path))
    if existing is not None and _same_path(existing, path):
        yield existing
        return

    spec = get_value(app, "GetOpenDocSpec", str(path))
    spec.DocumentType = SW_DOC_DRAWING if path.suffix.casefold() == ".slddrw" else SW_DOC_PART
    spec.ReadOnly = read_only
    spec.Silent = True
    doc = get_value(app, "OpenDoc7", spec)
    if doc is None:
        raise RuntimeError(f"SolidWorks не открыл файл {path.name}")
    if not _same_path(doc, path):
        raise RuntimeError("В SolidWorks уже открыт другой документ с таким именем")
    title = get_value(doc, "GetTitle")
    try:
        yield doc
    finally:
        _close(app, title)


def _close(app, title):
    try:
        get_value(app, "CloseDoc", title)
    except (AttributeError, com_error) as exc:
        print(f"[WARN] SolidWorks не смог закрыть документ — {exc}", flush=True)


def _save(drawing, target):
    extension = get_value(drawing, "Extension")
    try:
        get_value(extension, "SaveAs", str(target), SW_SAVE_CURRENT_VERSION,
                  SW_SAVE_SILENT, None, 0, 0)
    except com_error:
        # Some versions of the late-bound COM interface only expose SaveAs3.
        pass
    if target.is_file() and target.stat().st_size:
        return
    result = get_value(drawing, "SaveAs3", str(target), SW_SAVE_CURRENT_VERSION, SW_SAVE_SILENT)
    if not target.is_file() or not target.stat().st_size:
        raise RuntimeError(f"SolidWorks не сохранил {target.suffix} (код {result})")


def _hide_reference_markers(drawing):
    changed = False
    for setting in (SW_DISPLAY_ORIGINS, SW_DISPLAY_COORD_SYSTEMS):
        if get_value(drawing, "GetUserPreferenceToggle", setting):
            if not get_value(drawing, "SetUserPreferenceToggle", setting, False):
                raise RuntimeError("SolidWorks не скрыл обозначения координат на чертеже")
            changed = True
    if changed:
        get_value(drawing, "GraphicsRedraw2")


def _export_saved_flat_pdf(app, target, pdf):
    from .annotation_layout import geometry_outline, view_strokes
    from .drawing_views import drawing_views

    # Export from the persisted drawing after the writable copy of the
    # part has closed. A freshly created flat view can have its annotations
    # available before SolidWorks has a complete printable contour.
    with _open_document(app, target) as drawing:
        get_value(drawing, "ViewZoomtofit2")
        get_value(drawing, "GraphicsRedraw2")
        for view in drawing_views(drawing):
            if not get_value(view, "IsFlatPatternView"):
                continue
            outline = get_value(view, "GetOutline")[:4]
            contour = geometry_outline(view_strokes(view, outline), None)
            if contour is not None and contour[2] > contour[0] and contour[3] > contour[1]:
                _save(drawing, pdf)
                return
        raise RuntimeError("В сохранённом чертеже SolidWorks не подтвердил контур развёртки")


def convert_part(app, part, template, staging):
    """Generate native views and publishable files without reading the PDF."""
    from .model_geometry import read_model

    with _open_document(app, part) as model_doc:
        try:
            model = read_model(model_doc)
        except (AttributeError, com_error) as exc:
            print(f"[WARN] Геометрия через COM недоступна — {exc}", flush=True)
            model = None
        if model and model.sheet_metal:
            cache = Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local")
            root = cache / "Gostinator" / "models"
            cached_model = root / f"{uuid4().hex}.SLDPRT"
            pair = None
            try:
                root.mkdir(parents=True, exist_ok=True)
                copy2(part, cached_model)
                with _open_document(app, cached_model, read_only=False) as writable:
                    pair = _generate_drawing(app, cached_model, template, staging,
                                             model, writable_flat_source=writable)
                _export_saved_flat_pdf(app, *pair)
                print(f"    Связанная модель развёртки — {cached_model}", flush=True)
                return pair
            except (AttributeError, OSError, RuntimeError, com_error) as exc:
                try:
                    cached_model.unlink(missing_ok=True)
                except OSError:
                    pass
                if pair is not None:
                    raise RuntimeError(f"Не удалось выгрузить PDF с развёрткой — {exc}") from exc
                print(f"[WARN] Развёртка пропущена — {exc}", flush=True)
        try:
            return _generate_drawing(app, part, template, staging, model)
        except NativeDraftError as exc:
            if model is None:
                raise
            print(f"[WARN] Выбраны стандартные виды — {exc}", flush=True)
            return _generate_drawing(app, part, template, staging, None)


class NativeDraftError(RuntimeError):
    """The optional drawing geometry could not be safely applied."""


def _generate_drawing(app, part, template, staging, model, *, writable_flat_source=None):
    from .annotation_layout import arrange_dimensions, omit_repeated_section_scales, set_black_dimension_color
    from .drawing_dimensions import add_dimensions, prune_duplicate_bore_diameters, prune_model_dimensions
    from .drawing_views import add_view_centerlines, create_flat_pattern, create_views, sheet_frame
    from .hole_annotations import add_hole_center_marks, add_hole_quantities
    from .sheet_layout import (
        Layout, PAPERS, Paper, choose_layout, configure_sheet, installed_formats,
        synchronize_sheet_scale,
    )
    from .title_block import fit_title_block

    target = staging / (uuid4().hex + ".SLDDRW")
    pdf = target.with_suffix(".pdf")
    drawing = get_value(app, "NewDocument", str(template), 0, 0.0, 0.0)
    if drawing is None:
        raise RuntimeError("SolidWorks не создал чертёж по шаблону")
    try:
        frame = sheet_frame(drawing)
        layout = None
        include_flat = writable_flat_source is not None
        formats = installed_formats(app)
        if model is not None:
            if formats:
                layout = choose_layout(model, frame.first_angle, include_flat,
                                       papers=tuple(p for p in PAPERS if p in formats))
                configure_sheet(drawing, layout, formats[layout.paper], frame.first_angle)
                frame = sheet_frame(drawing)
                scale_text = f"{layout.scale:g}:1" if layout.scale >= 1 else f"1:{1/layout.scale:g}"
                print(f"    Лист {layout.paper.name}, масштаб {scale_text}", flush=True)
            else:
                paper = Paper("Шаблон", frame.width, frame.height, 12)
                layout = choose_layout(model, frame.first_angle, include_flat, papers=(paper,))
                print("[WARN] Установленный формат листа ГОСТ не найден, "
                      "используется текущий шаблон", flush=True)
        elif formats:
            # Surface models and unsupported feature sets still need a
            # native frame around SolidWorks' own standard views.
            paper = min(formats, key=lambda item: abs(item.width-frame.width) + abs(item.height-frame.height))
            configure_sheet(drawing, Layout(paper, 1., (), ()), formats[paper], frame.first_angle)
            frame = sheet_frame(drawing)
        try:
            views = create_views(drawing, part, model, frame, layout)
        except (AttributeError, RuntimeError, com_error) as exc:
            raise NativeDraftError(str(exc)) from exc
        if layout is not None:
            synchronize_sheet_scale(drawing, layout.scale)
            omit_repeated_section_scales(drawing)
        try:
            get_value(drawing, "InsertModelAnnotations3", SW_IMPORT_FROM_ENTIRE_MODEL,
                      SW_DIMENSIONS_MARKED_FOR_DRAWING, True, False, False, False)
        except com_error as exc:
            print(f"[WARN] Не удалось импортировать размеры модели — {exc}", flush=True)
        flat_model = None
        if include_flat:
            # The flat blank must not take ownership of any imported model
            # dimensions. Its envelope is dimensioned separately below.
            flat_name, flat_model = create_flat_pattern(drawing, part, layout, writable_flat_source)
            views.append(flat_name)
        if views:
            print(f"    Видов — {len(views)}", flush=True)
        if model is not None:
            try:
                skipped = prune_model_dimensions(drawing)
                count, cut_diameters, section_axes, hole_views = add_dimensions(
                    drawing, model, frame, flat_model=flat_model)
                visible, moved, hidden = arrange_dimensions(drawing, frame)
                duplicates = prune_duplicate_bore_diameters(drawing, model, cut_diameters)
            except (AttributeError, RuntimeError, com_error) as exc:
                raise NativeDraftError(str(exc)) from exc
            quantities = 0
            try:
                quantities, repeated = add_hole_quantities(drawing, frame, hole_views, cut_diameters)
                duplicates += repeated
            except (AttributeError, com_error) as exc:
                print(f"[WARN] Не удалось сгруппировать обозначения отверстий — {exc}", flush=True)
            visible -= duplicates
            skipped += duplicates
            print(f"    Размеров по геометрии — {count}, размеров скрыто — {skipped}",
                  flush=True)
            print(f"    Оставлено размеров — {visible}, перемещено — {moved}, скрыто при компоновке — {hidden}",
                  flush=True)
            if quantities:
                print(f"    Групповых обозначений отверстий — {quantities}", flush=True)
            try:
                centered = sum(add_view_centerlines(drawing, view, axes) for view, axes in section_axes)
                if centered:
                    print(f"    Осевых в разрезе — {centered}", flush=True)
            except (AttributeError, com_error) as exc:
                print(f"[WARN] Не удалось добавить осевые в разрез — {exc}", flush=True)
            try:
                marks = links = 0
                for view, holes, matrix, _ in hole_views:
                    made, joined = add_hole_center_marks(drawing, view, holes, matrix[12])
                    marks += made
                    links += joined
                if marks:
                    print(f"    Центровых обозначений — {marks}, соединений осей — {links}", flush=True)
            except (AttributeError, com_error) as exc:
                print(f"[WARN] Не удалось добавить центровые обозначения — {exc}", flush=True)
        if writable_flat_source is not None and not get_value(
                writable_flat_source, "Save3", SW_SAVE_SILENT,
                VARIANT(VT_BYREF | VT_I4, 0), VARIANT(VT_BYREF | VT_I4, 0)):
            raise RuntimeError("Не удалось сохранить конфигурацию развёртки")
        get_value(drawing, "ClearSelection2", True)
        fit_title_block(drawing, frame)
        _hide_reference_markers(drawing)
        try:
            set_black_dimension_color(drawing)
        except (AttributeError, com_error) as exc:
            print(f"[WARN] Не удалось задать чёрный цвет размеров — {exc}", flush=True)
        _save(drawing, target)
        if not include_flat:
            _save(drawing, pdf)
    finally:
        try:
            _close(app, get_value(drawing, "GetTitle"))
        except (AttributeError, com_error) as exc:
            print(f"[WARN] SolidWorks не вернул название чертежа — {exc}", flush=True)
    return target, pdf
