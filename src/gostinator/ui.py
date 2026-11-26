"""Tkinter front end for Gostinator."""
from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
from importlib.metadata import version
from pathlib import Path
from queue import Empty, Queue
import sys
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tomllib

from gostinator.cli import run_batch


DIRECTORY_EDITS = {
    ord("X"): ("Вырезать", "<<Cut>>"),
    ord("C"): ("Копировать", "<<Copy>>"),
    ord("V"): ("Вставить", "<<Paste>>"),
    ord("A"): ("Выделить всё", "<<SelectAll>>"),
}


def project_version():
    if getattr(sys, "frozen", False):
        config = Path(sys._MEIPASS) / "pyproject.toml"
    else:
        config = Path(__file__).resolve().parents[2] / "pyproject.toml"
        if not config.is_file():
            return version("gostinator")
    with config.open("rb") as project:
        return tomllib.load(project)["project"]["version"]


def path_text(variable):
    # Explorer's "Copy as path" includes double quotes.
    return variable.get().strip().strip('"')


class LogStream:
    encoding = "utf-8"

    def __init__(self, events):
        self.events = events

    def write(self, text):
        if text:
            self.events.put(("log", text))
        return len(text)

    def flush(self):
        pass


class GostinatorWindow:
    def __init__(self, root):
        self.root = root
        self.events = Queue()
        self.running = False
        self.status = tk.StringVar(value="Готово к запуску")
        self.parts = tk.StringVar(value=str(Path.cwd() / "in"))
        self.output = tk.StringVar(value=str(Path.cwd() / "out"))

        root.title(f"Gostinator-{project_version()}")
        root.minsize(610, 480)
        root.geometry("750x570")
        root.configure(background="#f5f7fa")
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.option_add("*Font", ("Segoe UI", 10))

        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f5f7fa")
        style.configure("TLabel", background="#f5f7fa", foreground="#172438")
        style.configure("Status.TLabel", foreground="#43556b")
        style.configure("Run.TButton", background="#184f9a", foreground="white",
                        padding=(16, 7), font=("Segoe UI", 10, "bold"))
        style.map("Run.TButton", background=[("disabled", "#677b96"), ("active", "#123d77")])
        style.configure("Horizontal.TProgressbar", background="#184f9a",
                        troughcolor="#d9e3ef", bordercolor="#d9e3ef")

        content = ttk.Frame(root, padding=(25, 18))
        content.pack(expand=True, fill="both")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(5, weight=1)
        self.directory_field(content, 0, "Папка моделей SolidWorks", self.parts)
        self.directory_field(content, 1, "Папка результатов", self.output)

        actions = ttk.Frame(content)
        actions.grid(row=2, sticky="ew", pady=(1, 12))
        self.run_button = ttk.Button(actions, text="Создать чертежи", style="Run.TButton", command=self.start)
        self.run_button.pack(side="left")

        self.progress = ttk.Progressbar(content, mode="determinate", maximum=1)
        self.progress.grid(row=3, sticky="ew")
        ttk.Label(content, textvariable=self.status, style="Status.TLabel").grid(row=4, sticky="w", pady=(4, 11))

        log_frame = ttk.LabelFrame(content, text="Журнал", padding=(8, 5))
        log_frame.grid(row=5, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side="right", fill="y")
        self.log = tk.Text(log_frame, wrap="word", state="disabled", height=8, borderwidth=0,
                           font=("Consolas", 10), background="white", foreground="#172438",
                           padx=8, pady=6, yscrollcommand=scrollbar.set)
        self.log.pack(side="left", fill="both", expand=True)
        scrollbar.configure(command=self.log.yview)
        root.after(60, self.poll)

    def directory_field(self, parent, row, caption, variable):
        frame = ttk.Frame(parent)
        frame.grid(row=row, sticky="ew", pady=(0, 10))
        frame.columnconfigure(0, weight=1)
        ttk.Label(frame, text=caption).grid(row=0, column=0, sticky="w", pady=(0, 3))
        entry = ttk.Entry(frame, textvariable=variable)
        entry.grid(row=1, column=0, sticky="ew")

        menu = tk.Menu(entry, tearoff=False)
        for key, (label, action) in DIRECTORY_EDITS.items():
            menu.add_command(label=label, accelerator=f"Ctrl+{chr(key)}",
                             command=lambda action=action: entry.event_generate(action))

        def show_edit_menu(event):
            entry.focus_set()
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        def edit_shortcut(event):
            # Tk versions differ in resolving Ctrl shortcuts on non-Latin
            # layouts. Windows virtual key codes keep these letters stable.
            if event.state & 0x20010:  # Alt or AltGr
                return None
            edit = DIRECTORY_EDITS.get(event.keycode)
            if edit:
                entry.event_generate(edit[1])
                return "break"
            return None

        entry.bind("<<ContextMenu>>", show_edit_menu)
        entry.bind("<Control-KeyPress>", edit_shortcut)

        def browse():
            current = Path(path_text(variable)).expanduser()
            initial = current if current.is_dir() else current.parent
            selected = filedialog.askdirectory(parent=self.root, title=caption, initialdir=initial)
            if selected:
                variable.set(selected)

        button = ttk.Button(frame, text="Обзор…", command=browse)
        button.grid(row=1, column=1, padx=(8, 0))

    def start(self):
        if self.running:
            return
        parts, output = path_text(self.parts), path_text(self.output)
        if not parts or not output:
            messagebox.showerror("Gostinator", "Укажите обе папки.", parent=self.root)
            return
        self.running = True
        self.status.set("Подключение к SolidWorks…")
        self.progress.configure(value=0, maximum=1)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.run_button.state(["disabled"])
        Thread(target=self.convert, args=(Path(parts), Path(output))).start()

    def convert(self, parts, output):
        stream = LogStream(self.events)
        with redirect_stdout(stream), redirect_stderr(stream):
            try:
                result = run_batch(parts, output, on_progress=self.update_progress)
            except Exception as exc:
                result = 2
                print(f"[ERROR] {exc}")
        self.events.put(("done", result))

    def update_progress(self, completed, total):
        self.events.put(("progress", (completed, total)))

    def poll(self):
        while True:
            try:
                event, payload = self.events.get_nowait()
            except Empty:
                break
            if event == "log":
                self.log.configure(state="normal")
                self.log.insert("end", payload)
                self.log.see("end")
                self.log.configure(state="disabled")
            elif event == "progress":
                completed, total = payload
                self.progress.configure(value=completed, maximum=total)
                self.status.set(f"Обработано {completed} из {total}")
            elif event == "done":
                self.running = False
                self.run_button.state(["!disabled"])
                self.status.set("Готово" if payload == 0 else "Есть ошибки. Подробности в журнале.")
        self.root.after(60, self.poll)

    def close(self):
        if self.running:
            messagebox.showinfo("Gostinator", "Дождитесь завершения обработки.", parent=self.root)
        else:
            self.root.destroy()


def main():
    root = tk.Tk()
    GostinatorWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
