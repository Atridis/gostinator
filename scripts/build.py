"""Build two standalone Windows programs and check the bundled runtimes."""
from __future__ import annotations

import argparse
import os
from pathlib import Path
import struct
import subprocess
import sys
from tempfile import TemporaryDirectory
import tomllib
from zipfile import ZIP_STORED, ZipFile


ROOT = Path(__file__).resolve().parents[1]
PROGRAMS = ("gostinator.exe", "gostinator-ui.exe")


def build(output: Path, work: Path):
    if sys.platform != "win32" or struct.calcsize("P") != 8:
        raise RuntimeError("Build Gostinator using 64-bit Python on Windows")

    output.mkdir(parents=True, exist_ok=True)
    work.mkdir(parents=True, exist_ok=True)
    subprocess.run(
        [sys.executable, "-m", "PyInstaller", "--noconfirm", "--distpath", str(output),
         "--workpath", str(work), str(ROOT / "scripts" / "gostinator.spec")],
        cwd=ROOT, check=True,
    )
    smoke(output)
    archive = archive_programs(output)
    if github_output := os.environ.get("GITHUB_OUTPUT"):
        with Path(github_output).open("a", encoding="utf-8") as outputs:
            outputs.write(f"archive_path={archive}\n")


def smoke(output: Path):
    from PyInstaller.archive.readers import CArchiveReader

    cli, gui = (output / name for name in PROGRAMS)
    with TemporaryDirectory(prefix="gostinator-smoke-") as folder:
        # Launch outside the checkout so only bundled modules are reachable.
        env = os.environ.copy()
        env.pop("PYTHONPATH", None)
        env.pop("PYTHONHOME", None)
        subprocess.run([str(cli), "--help"], cwd=folder, env=env, check=True, timeout=180)
    # CI may run without an interactive desktop. Check the data that the
    # windowed executable needs before publishing its archive.
    for program in (cli, gui):
        executable = CArchiveReader(program)
        bundled = {name.replace(chr(92), "/").lower() for name in executable.toc}
        required = {"pywin32_system32/pythoncom312.dll", "pywin32_system32/pywintypes312.dll"}
        if program == gui:
            required.update({"ui", "_tkinter.pyd", "_tcl_data/init.tcl", "_tk_data/tk.tcl"})
        if missing := required - bundled:
            raise RuntimeError(f"{program.name} is missing {', '.join(sorted(missing))}")
        python_modules = executable.open_embedded_archive("PYZ.pyz").toc
        roots = {name.split("/")[0].split(".")[0] for name in bundled}
        roots.update(name.split(".")[0].lower() for name in python_modules)
        forbidden = {"cadquery", "pymupdf", "ocp", "vtkmodules", "numpy", "shapely"}
        if bundled_optional := forbidden & roots:
            raise RuntimeError(f"{program.name} contains {', '.join(sorted(bundled_optional))}")


def archive_programs(output: Path):
    # The one-file EXEs are already compressed.
    with (ROOT / "pyproject.toml").open("rb") as metadata:
        version = tomllib.load(metadata)["project"]["version"]
    archive = output / f"gostinator-{version}-windows-x64.zip"
    temporary_archive = archive.with_suffix(".tmp")
    with ZipFile(temporary_archive, "w", compression=ZIP_STORED) as package:
        for program in PROGRAMS:
            executable = output / program
            package.write(executable, arcname=executable.name)
    temporary_archive.replace(archive)
    print(archive)
    return archive


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output-dir", type=Path, default=ROOT / "dist")
    parser.add_argument("--work-dir", type=Path, default=ROOT / "build")
    args = parser.parse_args()
    build(args.output_dir.resolve(), args.work_dir.resolve())
