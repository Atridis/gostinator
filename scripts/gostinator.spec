# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import sys
import tkinter

project = Path(SPECPATH).parent
source = project / "src"

def build_program(name, entry, *, console, datas=(), hookspath=(), excludes=()):
    analysis = Analysis(
        [str(source / "gostinator" / entry)],
        pathex=[str(source)],
        datas=list(datas),
        hookspath=list(hookspath),
        excludes=list(excludes),
    )
    return EXE(
        PYZ(analysis.pure), analysis.scripts, analysis.binaries, analysis.datas,
        name=name, console=console, upx=False,
    )

cli = build_program("gostinator", "__main__.py", console=True, excludes=["tkinter"])

tcl_root = Path(sys.base_prefix) / "tcl"
tk_data = []
for library, version in (("tcl", tkinter.TclVersion), ("tk", tkinter.TkVersion)):
    folder = tcl_root / f"{library}{version}"
    if folder.is_dir():
        tk_data.append((str(folder), f"_{library}_data"))
gui = build_program(
    "gostinator-ui", "ui.py", console=False, datas=[*tk_data, (str(project / "pyproject.toml"), ".")],
    hookspath=[str(Path(SPECPATH) / "hooks")] if len(tk_data) == 2 else [],
)
