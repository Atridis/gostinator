"""Tkinter front end for Gostinator."""
from __future__ import annotations

from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from queue import Empty, Queue
from threading import Thread
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from gostinator.cli import run_batch


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
        self.controls = []

        root.title("Gostinator")
        root.minsize(610, 480)
        root.geometry("750x570")
        root.configure(background="#f5f7fa")
        root.protocol("WM_DELETE_WINDOW", self.close)
        root.option_add("*Font", ("Segoe UI", 10))

        style = ttk.Style(root)
        style.theme_use("clam")
        style.configure("TFrame", background="#f5f7fa")
        style.configure("TLabel", background="#f5f7fa", foreground="#172438")
        style.configure("Heading.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Status.TLabel", foreground="#43556b")
        style.configure("Run.TButton", background="#184f9a", foreground="white",
                        padding=(16, 7), font=("Segoe UI", 10, "bold"))
        style.map("Run.TButton", background=[("disabled", "#677b96"), ("active", "#123d77")])
        style.configure("Horizontal.TProgressbar", background="#184f9a",
                        troughcolor="#d9e3ef", bordercolor="#d9e3ef")

        content = ttk.Frame(root, padding=(25, 18))
        content.pack(expand=True, fill="both")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(6, weight=1)
        ttk.Label(content, text="Gostinator", style="Heading.TLabel").grid(row=0, sticky="w", pady=(0, 12))
        self.directory_field(content, 1, "Папка моделей SolidWorks", self.parts)
        self.directory_field(content, 2, "Папка результатов", self.output)

        actions = ttk.Frame(content)
        actions.grid(row=3, sticky="ew", pady=(1, 12))
        run_button = ttk.Button(actions, text="Создать чертежи", style="Run.TButton", command=self.start)
        run_button.pack(side="left")
        self.controls.append(run_button)

        self.progress = ttk.Progressbar(content, mode="determinate", maximum=1)
        self.progress.grid(row=4, sticky="ew")
        ttk.Label(content, textvariable=self.status, style="Status.TLabel").grid(row=5, sticky="w", pady=(4, 11))

        log_frame = ttk.LabelFrame(content, text="Журнал", padding=(8, 5))
        log_frame.grid(row=6, sticky="nsew")
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

        def browse():
            current = Path(variable.get()).expanduser()
            initial = current if current.is_dir() else current.parent
            selected = filedialog.askdirectory(parent=self.root, title=caption, initialdir=initial)
            if selected:
                variable.set(selected)

        button = ttk.Button(frame, text="Обзор…", command=browse)
        button.grid(row=1, column=1, padx=(8, 0))
        self.controls.extend((entry, button))

    def start(self):
        if self.running:
            return
        parts, output = self.parts.get().strip(), self.output.get().strip()
        if not parts or not output:
            messagebox.showerror("Gostinator", "Укажите обе папки.", parent=self.root)
            return
        self.running = True
        self.status.set("Подключение к SolidWorks…")
        self.progress.configure(value=0, maximum=1)
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        for control in self.controls:
            control.state(["disabled"])
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
                for control in self.controls:
                    control.state(["!disabled"])
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
