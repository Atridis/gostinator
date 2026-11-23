"""Batch entry point shared by the console and Tkinter applications."""
from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import sys
from tempfile import TemporaryDirectory

from . import solidworks


def _input_parts(folder):
    return sorted(
        (part for part in folder.iterdir()
         if part.is_file() and part.suffix.lower() == ".sldprt" and not part.name.startswith("~$")),
        key=lambda part: part.name.casefold(),
    )


def _publish_pair(drawing, pdf, output, stem):
    destinations = (output / (stem + ".SLDDRW"), output / (stem + ".pdf"))
    # Stage on the target volume. Windows rename refuses to replace an
    # existing destination, including files created by a concurrent run.
    with TemporaryDirectory(prefix=".gostinator-", dir=output) as folder:
        staged = (Path(folder) / "drawing.SLDDRW", Path(folder) / "drawing.pdf")
        for source, temporary in zip((drawing, pdf), staged):
            if not source.is_file() or not source.stat().st_size:
                raise ValueError(f"Отсутствует итоговый файл {source.suffix}")
            shutil.copyfile(source, temporary)
        published = []
        try:
            for temporary, destination in zip(staged, destinations):
                temporary.rename(destination)
                published.append(destination)
        except BaseException:
            for destination in published:
                destination.unlink(missing_ok=True)
            raise
    return destinations


def run_batch(parts_dir: Path, out_dir: Path, *, on_progress=None):
    parts_dir, out_dir = Path(parts_dir).resolve(), Path(out_dir).resolve()
    if not parts_dir.is_dir():
        raise ValueError(f"Папка моделей не найдена — {parts_dir}")
    if parts_dir == out_dir:
        raise ValueError("Выберите разные папки для моделей и результатов")
    parts = _input_parts(parts_dir)
    if not parts:
        raise ValueError("В папке нет файлов .SLDPRT")
    for part in parts:
        for suffix in (".SLDDRW", ".pdf"):
            target = out_dir / (part.stem + suffix)
            if target.exists():
                raise FileExistsError(f"Файл уже существует — {target}")
    out_dir.mkdir(parents=True, exist_ok=True)

    published = 0
    with solidworks.session() as app, TemporaryDirectory(
            prefix="gostinator-", ignore_cleanup_errors=True) as folder:
        template = solidworks.drawing_template(app)
        print(f"Шаблон — {template.name}", flush=True)
        for index, part in enumerate(parts, 1):
            print(f"[{index}/{len(parts)}] {part.name}", flush=True)
            try:
                drawing, pdf = solidworks.convert_part(app, part, template, Path(folder))
                final_drawing, final_pdf = _publish_pair(drawing, pdf, out_dir, part.stem)
                print(f"    {final_drawing.name} · {final_pdf.name}", flush=True)
                published += 1
            except (OSError, ValueError, RuntimeError, solidworks.com_error) as exc:
                print(f"[ERROR] {part.name} — {exc}", flush=True)
            finally:
                if on_progress:
                    on_progress(index, len(parts))
    print(f"Готово — {published} из {len(parts)}", flush=True)
    return 0 if published == len(parts) else 2


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            stream.reconfigure(encoding="utf-8", errors="backslashreplace")
    parser = argparse.ArgumentParser(
        prog="gostinator", description="Пакетная генерация чертежей SolidWorks")
    parser.add_argument("--parts-dir", type=Path, required=True, help="Папка с моделями .SLDPRT")
    parser.add_argument("--out-dir", type=Path, required=True, help="Папка для чертежей и PDF")
    args = parser.parse_args(argv)
    try:
        return run_batch(args.parts_dir, args.out_dir)
    except Exception as exc:
        print(f"[ERROR] {exc}", file=sys.stderr, flush=True)
        return 2
