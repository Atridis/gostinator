"""Keep Tkinter when the spec already supplies both Tcl and Tk libraries."""


def pre_find_module_path(api):
    # A non-interactive build account can fail PyInstaller's Tcl runtime probe.
    # The spec includes the data, and the finished executable is smoke-tested.
    pass
