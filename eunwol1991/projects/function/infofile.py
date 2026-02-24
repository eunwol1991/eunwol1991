import os
import threading
import tkinter as tk
from tkinter import filedialog, font, scrolledtext, ttk
from typing import Callable, Optional


def _is_wsl() -> bool:
    if os.name == "nt":
        return False
    if os.environ.get("WSL_DISTRO_NAME"):
        return True
    try:
        with open("/proc/version", "r", encoding="utf-8") as f:
            return "microsoft" in f.read().lower()
    except OSError:
        return False


def _platform_drive_root() -> str:
    if os.name == "nt":
        return "c:/"
    if _is_wsl():
        return "/mnt/c"
    return "/"


def _from_c(path_tail: str) -> str:
    tail = (path_tail or "").lstrip("/")
    root = _platform_drive_root()
    if root.endswith("/"):
        return f"{root}{tail}"
    return f"{root}/{tail}"


BASE_DIR = _from_c("Users/jhunj/Dropbox/DO & INV/DO & INV 2026")

TOKYONIGHT = {
    "bg": "#1a1b26",
    "surface": "#24283b",
    "surface_alt": "#1f2335",
    "text": "#c0caf5",
    "muted": "#9aa5ce",
    "border": "#414868",
    "accent": "#7aa2f7",
    "accent_hover": "#89b4fa",
    "success": "#9ece6a",
    "warning": "#e0af68",
    "danger": "#f7768e",
    "info": "#7dcfff",
}

MONTH_MAP = {
    1: "1. Jan",
    2: "2. Feb",
    3: "3. Mar",
    4: "4. Apr",
    5: "5. May",
    6: "6. Jun",
    7: "7. Jul",
    8: "8. Aug",
    9: "9. Sep",
    10: "10. Oct",
    11: "11. Nov",
    12: "12. Dec",
}


def search_month_folder(
    month: int,
    logger,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
):
    month_name = MONTH_MAP.get(month)
    if not month_name:
        logger("Invalid month number")
        return
    found = False
    walks = list(os.walk(BASE_DIR))
    total = max(1, len(walks))
    for idx, (root, dirs, _) in enumerate(walks, start=1):
        if progress_cb:
            progress_cb(total, idx, "Scanning folders...")
        for d in dirs:
            if d.lower() == month_name.lower():
                logger(f"Found: {os.path.join(root, d)}")
                found = True
    if not found:
        logger("Not found")


def clean_empty_month_folders(
    logger,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
):
    targets = []
    for root, _dirs, _ in os.walk(BASE_DIR):
        for name in MONTH_MAP.values():
            path = os.path.join(root, name)
            if os.path.isdir(path):
                targets.append(path)

    deleted = skipped = failed = 0
    total = max(1, len(targets))
    for idx, path in enumerate(targets, start=1):
        if progress_cb:
            progress_cb(total, idx, "Cleaning folders...")
        if not os.listdir(path):
            try:
                os.rmdir(path)
                logger(f"✅ Deleted: {path}", "success")
                deleted += 1
            except Exception as e:
                logger(f"❌ Delete failed: {path} -> {e}", "fail")
                failed += 1
        else:
            logger(f"⚠️ Kept: {path} (not empty)", "skip")
            skipped += 1
    logger(
        f"Summary: deleted {deleted}, kept {skipped}, failed {failed}",
        "info",
    )


def create_month_folders(
    month: int,
    logger,
    progress_cb: Optional[Callable[[int, int, str], None]] = None,
):
    month_name = MONTH_MAP.get(month)
    if not month_name:
        logger("Invalid month number")
        return
    created = skipped = failed = 0
    month_dirs_lower = {m.lower() for m in MONTH_MAP.values()}
    eligible_roots = []
    for root, dirs, _ in os.walk(BASE_DIR):
        if root == BASE_DIR:
            # 跳过根目录
            continue

        # 仅当当前目录已经包含任意月份文件夹时才补齐
        has_month = any(d.lower() in month_dirs_lower for d in dirs)
        if not has_month:
            continue

        eligible_roots.append(root)

    total = max(1, len(eligible_roots))
    for idx, root in enumerate(eligible_roots, start=1):
        if progress_cb:
            progress_cb(total, idx, "Creating folders...")

        month_path = os.path.join(root, month_name)
        if os.path.exists(month_path):
            logger(f"⚠️ Skipped: {month_path} (already exists)", "skip")
            skipped += 1
        else:
            try:
                os.makedirs(month_path)
                logger(f"✅ Created: {month_path}", "success")
                created += 1
            except Exception as e:
                logger(f"❌ Create failed: {month_path} -> {e}", "fail")
                failed += 1
    logger(
        f"Summary: created {created}, skipped {skipped}, failed {failed}",
        "info",
    )


class InfoApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.option = tk.IntVar(value=1)
        self.month_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self.title_font = font.Font(family="Segoe UI", size=20, weight="bold")
        self.subtitle_font = font.Font(family="Segoe UI", size=10)
        self.label_font = font.Font(family="Segoe UI", size=11)
        self.log_font = font.Font(family="Consolas", size=10)

        self.style = ttk.Style()
        self.style.theme_use("clam")
        self._configure_styles()

        master.title("Info File Utility")
        master.geometry("1080x680")
        master.minsize(960, 620)
        master.configure(bg=TOKYONIGHT["bg"])
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)

        shell = ttk.Frame(master, style="App.TFrame", padding=(18, 18, 18, 14))
        shell.grid(row=0, column=0, sticky="nsew")
        shell.columnconfigure(0, weight=1)
        shell.rowconfigure(1, weight=1)

        header = ttk.Frame(shell, style="App.TFrame")
        header.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        header.columnconfigure(0, weight=1)

        ttk.Label(header, text="Info File Utility", style="Title.TLabel").grid(
            row=0, column=0, sticky="w"
        )
        ttk.Label(
            header,
            text=f"Base: {BASE_DIR}",
            style="Muted.TLabel",
        ).grid(row=1, column=0, sticky="w", pady=(2, 0))

        body = ttk.Frame(shell, style="App.TFrame")
        body.grid(row=1, column=0, sticky="nsew")
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=1)
        body.rowconfigure(0, weight=1)

        control = ttk.Frame(body, style="Card.TFrame", padding=(14, 14, 14, 14))
        control.grid(row=0, column=0, sticky="nsw", padx=(0, 12))

        ttk.Label(control, text="Actions", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        ttk.Radiobutton(
            control,
            text="1. Smart subfolder search",
            variable=self.option,
            value=1,
            style="App.TRadiobutton",
        ).grid(row=1, column=0, sticky="w", pady=3)
        ttk.Radiobutton(
            control,
            text="2. Clean empty month folders",
            variable=self.option,
            value=2,
            style="App.TRadiobutton",
        ).grid(row=2, column=0, sticky="w", pady=3)
        ttk.Radiobutton(
            control,
            text="3. Smart create month folders",
            variable=self.option,
            value=3,
            style="App.TRadiobutton",
        ).grid(row=3, column=0, sticky="w", pady=3)

        ttk.Label(control, text="Month", style="Body.TLabel").grid(
            row=4, column=0, sticky="w", pady=(12, 4)
        )
        month_entry = ttk.Entry(
            control,
            textvariable=self.month_var,
            width=16,
            style="App.TEntry",
        )
        month_entry.grid(row=5, column=0, sticky="ew")

        self.run_btn = ttk.Button(
            control,
            text="Run",
            command=self.start,
            style="Accent.TButton",
        )
        self.run_btn.grid(row=6, column=0, sticky="ew", pady=(14, 6))

        self.export_btn = ttk.Button(
            control,
            text="Export Log",
            command=self.export_log,
            style="App.TButton",
        )
        self.export_btn.grid(row=7, column=0, sticky="ew")

        log_frame = ttk.Frame(body, style="Card.TFrame", padding=(12, 12, 12, 12))
        log_frame.grid(row=0, column=1, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)

        ttk.Label(log_frame, text="Log Output", style="Section.TLabel").grid(
            row=0, column=0, sticky="w", pady=(0, 8)
        )

        self.log_widget = scrolledtext.ScrolledText(
            log_frame,
            state=tk.DISABLED,
            font=self.log_font,
            background=TOKYONIGHT["surface"],
            foreground=TOKYONIGHT["text"],
            insertbackground=TOKYONIGHT["text"],
            highlightthickness=1,
            highlightbackground=TOKYONIGHT["border"],
            highlightcolor=TOKYONIGHT["accent"],
            borderwidth=0,
            relief=tk.FLAT,
        )
        self.log_widget.tag_config("success", foreground=TOKYONIGHT["success"])
        self.log_widget.tag_config("fail", foreground=TOKYONIGHT["danger"])
        self.log_widget.tag_config("skip", foreground=TOKYONIGHT["warning"])
        self.log_widget.tag_config("info", foreground=TOKYONIGHT["info"])
        self.log_widget.grid(row=1, column=0, sticky="nsew")

        footer = ttk.Frame(shell, style="App.TFrame")
        footer.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        footer.columnconfigure(0, weight=1)

        self.status_label = ttk.Label(
            footer,
            textvariable=self.status_var,
            style="Status.TLabel",
            anchor="w",
        )
        self.status_label.grid(row=0, column=0, sticky="ew")

        self.progress = ttk.Progressbar(
            footer,
            mode="determinate",
            style="App.Horizontal.TProgressbar",
            maximum=100,
            value=0,
        )
        self.progress.grid(row=1, column=0, sticky="ew", pady=(6, 0))

        master.bind("<Configure>", self.on_resize)

    def _configure_styles(self):
        self.style.configure("App.TFrame", background=TOKYONIGHT["bg"])
        self.style.configure(
            "Card.TFrame",
            background=TOKYONIGHT["surface_alt"],
            borderwidth=1,
            relief="solid",
            bordercolor=TOKYONIGHT["border"],
        )
        self.style.configure(
            "Title.TLabel",
            background=TOKYONIGHT["bg"],
            foreground=TOKYONIGHT["text"],
            font=self.title_font,
        )
        self.style.configure(
            "Section.TLabel",
            background=TOKYONIGHT["surface_alt"],
            foreground=TOKYONIGHT["text"],
            font=("Segoe UI", 12, "bold"),
        )
        self.style.configure(
            "Muted.TLabel",
            background=TOKYONIGHT["bg"],
            foreground=TOKYONIGHT["muted"],
            font=self.subtitle_font,
        )
        self.style.configure(
            "Body.TLabel",
            background=TOKYONIGHT["surface_alt"],
            foreground=TOKYONIGHT["text"],
            font=self.label_font,
        )
        self.style.configure(
            "Status.TLabel",
            background=TOKYONIGHT["bg"],
            foreground=TOKYONIGHT["muted"],
            font=self.label_font,
        )
        self.style.configure(
            "App.TRadiobutton",
            background=TOKYONIGHT["surface_alt"],
            foreground=TOKYONIGHT["text"],
            font=self.label_font,
            padding=2,
        )
        self.style.map(
            "App.TRadiobutton",
            background=[("active", TOKYONIGHT["surface_alt"])],
            foreground=[("active", TOKYONIGHT["text"])],
        )
        self.style.configure(
            "App.TEntry",
            fieldbackground=TOKYONIGHT["surface"],
            background=TOKYONIGHT["surface"],
            foreground=TOKYONIGHT["text"],
            bordercolor=TOKYONIGHT["border"],
            insertcolor=TOKYONIGHT["text"],
            relief="flat",
            padding=(8, 6),
        )
        self.style.map(
            "App.TEntry",
            fieldbackground=[("focus", TOKYONIGHT["surface"])],
            bordercolor=[("focus", TOKYONIGHT["accent"])],
        )
        self.style.configure(
            "App.TButton",
            background=TOKYONIGHT["surface"],
            foreground=TOKYONIGHT["text"],
            bordercolor=TOKYONIGHT["border"],
            font=self.label_font,
            padding=(10, 8),
            relief="flat",
        )
        self.style.map(
            "App.TButton",
            background=[("active", TOKYONIGHT["surface_alt"])],
            foreground=[("active", TOKYONIGHT["text"])],
        )
        self.style.configure(
            "Accent.TButton",
            background=TOKYONIGHT["accent"],
            foreground=TOKYONIGHT["bg"],
            bordercolor=TOKYONIGHT["accent"],
            font=self.label_font,
            padding=(10, 8),
            relief="flat",
        )
        self.style.map(
            "Accent.TButton",
            background=[("active", TOKYONIGHT["accent_hover"])],
            foreground=[("active", TOKYONIGHT["bg"])],
        )
        self.style.configure(
            "App.Horizontal.TProgressbar",
            troughcolor=TOKYONIGHT["surface"],
            background=TOKYONIGHT["accent"],
            bordercolor=TOKYONIGHT["border"],
            lightcolor=TOKYONIGHT["accent"],
            darkcolor=TOKYONIGHT["accent"],
        )

    def log(self, message: str, tag: Optional[str] = None):
        self.log_widget.after(0, self._append_log, message, tag)

    def _append_log(self, message: str, tag: Optional[str] = None):
        self.log_widget.config(state=tk.NORMAL)
        if tag:
            self.log_widget.insert(tk.END, message + "\n", tag)
        else:
            self.log_widget.insert(tk.END, message + "\n")
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)

    def start(self):
        self.status_var.set("Processing...")
        self.progress.configure(mode="determinate", maximum=100, value=0)
        self.run_btn.configure(state=tk.DISABLED)
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.config(state=tk.DISABLED)
        threading.Thread(target=self.execute, daemon=True).start()

    def _report_progress(self, total: int, current: int, note: str):
        self.master.after(0, self._apply_progress, total, current, note)

    def _apply_progress(self, total: int, current: int, note: str):
        safe_total = max(1, int(total))
        safe_current = max(0, min(int(current), safe_total))
        pct = int((safe_current / safe_total) * 100)
        self.progress.configure(value=pct)
        self.status_var.set(f"{note} {pct}%")

    def export_log(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Text Files", "*.txt")]
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.log_widget.get("1.0", tk.END))
            self.status_var.set("Log exported")

    def execute(self):
        try:
            month = int(self.month_var.get()) if self.month_var.get() else None
        except ValueError:
            month = None

        if self.option.get() in (1, 3) and month is None:
            self.log("Please enter a month number")
            self.master.after(0, self._finish_ui, "Waiting for input")
            return

        try:
            if self.option.get() == 1:
                if month is None:
                    self.master.after(0, self._finish_ui, "Waiting for input")
                    return
                search_month_folder(month, self.log, self._report_progress)
            elif self.option.get() == 2:
                clean_empty_month_folders(self.log, self._report_progress)
            else:
                if month is None:
                    self.master.after(0, self._finish_ui, "Waiting for input")
                    return
                create_month_folders(month, self.log, self._report_progress)
            self.master.after(0, self._finish_ui, "Done")
        except Exception as e:
            self.log(f"Error occurred: {e}", "fail")
            self.master.after(0, self._finish_ui, "Error")

    def _finish_ui(self, status_text: str):
        self.status_var.set(status_text)
        if status_text == "Done":
            self.progress.configure(value=100)
        elif status_text in ("Waiting for input", "Error"):
            self.progress.configure(value=0)
        self.run_btn.configure(state=tk.NORMAL)

    def on_resize(self, event):
        width = max(event.width, 600)
        base = max(9, int(width / 92))
        self.title_font.configure(size=base + 10)
        self.subtitle_font.configure(size=max(9, base - 1))
        self.label_font.configure(size=base)
        self.log_font.configure(size=max(8, base - 1))


def main():
    root = tk.Tk()
    InfoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
