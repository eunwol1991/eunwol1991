import os
import re
import shutil
from datetime import date, timedelta

import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from ttkbootstrap import Treeview
from tkinter import filedialog, messagebox
from tkinter import font as tkfont
from tkinter import BooleanVar, DoubleVar, StringVar, TclError


def _windows_path_to_wsl(path: str) -> str:
    text = (path or "").strip()
    match = re.match(r"^([A-Za-z]):\\(.*)$", text)
    if not match:
        return text
    drive = match.group(1).lower()
    rest = match.group(2).replace("\\", "/")
    return f"/mnt/{drive}/{rest}"


def _first_existing_path(candidates):
    for candidate in candidates:
        if candidate and os.path.isdir(candidate):
            return candidate
    return ""


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


def _default_dropbox_base() -> str:
    root = _platform_drive_root()
    if root.endswith("/"):
        return f"{root}Users/jhunj/Dropbox"
    return f"{root}/Users/jhunj/Dropbox"


def main():
    file_info_list = []
    selected_files = []
    current_source_dir = [""]

    dropbox_base = _default_dropbox_base()
    source_default = f"{dropbox_base}/DO & INV/DO & INV 2026"
    target_default = f"{dropbox_base}/for jj/Doc to print - JJ"

    default_source_dir = _first_existing_path(
        [
            source_default,
            os.getcwd(),
        ]
    )
    last_directory = [default_source_dir]

    base_width, base_height = 920, 740
    target_dir = _first_existing_path(
        [
            target_default,
        ]
    )

    app = ttk.Window(themename="darkly")
    app.title("File Selector")
    app.geometry(f"{base_width}x{base_height}")
    app.resizable(True, True)

    style = ttk.Style()

    base_body_size = 12
    base_small_size = 10
    base_section_size = 14
    base_title_size = 18

    ui_font = tkfont.Font(family="Arial", size=base_body_size)
    small_font = tkfont.Font(family="Arial", size=base_small_size)
    section_font = tkfont.Font(family="Arial", size=base_section_size)
    title_font = tkfont.Font(family="Arial", size=base_title_size)
    tree_font_main = tkfont.Font(family="Arial", size=base_body_size)
    tree_font_chosen = tkfont.Font(family="Arial", size=base_body_size)

    style.configure("TLabel", font=ui_font)
    style.configure("TButton", font=ui_font)
    style.configure("TCheckbutton", font=ui_font)
    style.configure("TEntry", font=ui_font)
    style.configure("Zoom.TLabel", foreground="#FFFFFF", font=small_font)

    style.configure("Main.Treeview", font=tree_font_main)
    style.configure("Chosen.Treeview", font=tree_font_chosen)

    def rowheight_for(font_obj, factor: float) -> int:
        try:
            line_space = int(font_obj.metrics("linespace"))
        except Exception:
            line_space = int(14 * factor)
        gap = max(2, int(4 * factor))
        return max(24, line_space + gap * 2 + 2)

    scale_var = DoubleVar(value=1.49)

    def apply_scale(value):
        try:
            factor = float(value)
        except Exception:
            factor = 1.0

        try:
            app.tk.call("tk", "scaling", factor)
        except Exception:
            pass

        body_size = max(8, int(round(base_body_size * factor)))
        small_size = max(8, int(round(base_small_size * factor)))
        section_size = max(9, int(round(base_section_size * factor)))
        title_size = max(10, int(round(base_title_size * factor)))

        ui_font.configure(size=body_size)
        small_font.configure(size=small_size)
        section_font.configure(size=section_size)
        title_font.configure(size=title_size)
        tree_font_main.configure(size=body_size)
        tree_font_chosen.configure(size=body_size)

        style.configure("TLabel", font=ui_font)
        style.configure("TButton", font=ui_font)
        style.configure("TCheckbutton", font=ui_font)
        style.configure("TEntry", font=ui_font)
        style.configure(
            "Main.Treeview",
            font=tree_font_main,
            rowheight=rowheight_for(tree_font_main, factor),
        )
        style.configure(
            "Chosen.Treeview",
            font=tree_font_chosen,
            rowheight=rowheight_for(tree_font_chosen, factor),
        )
        style.configure(
            "Treeview.Heading",
            font=ui_font,
            padding=(max(4, int(6 * factor)),),
        )
        app.geometry(f"{int(base_width * factor)}x{int(base_height * factor)}")
        zoom_label.configure(text=f"UI Scale: {int(round(factor * 100))}%")

    topbar = ttk.Frame(app)
    topbar.pack(fill=X, padx=10, pady=(10, 0))

    ttk.Label(
        topbar,
        text="Choose Source Folder and Match Files",
        font=title_font,
        bootstyle="info",
    ).pack(side=LEFT, padx=(0, 10))

    zoom_label = ttk.Label(topbar, text="UI Scale: 149%", style="Zoom.TLabel")
    zoom_label.pack(side=RIGHT, padx=(10, 0))

    ttk.Scale(
        topbar,
        from_=0.8,
        to=1.6,
        orient=HORIZONTAL,
        variable=scale_var,
        command=apply_scale,
        bootstyle=INFO,
        length=220,
    ).pack(side=RIGHT)

    file_frame = ttk.Frame(app)
    file_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

    file_tree = Treeview(
        file_frame,
        columns=("#1",),
        show="headings",
        height=10,
        style="Main.Treeview",
        bootstyle="info",
    )
    file_tree.heading("#1", text="Files")
    file_tree.column("#1", anchor="w")
    file_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))

    scrollbar = ttk.Scrollbar(file_frame, orient="vertical", command=file_tree.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    file_tree.config(yscrollcommand=scrollbar.set)

    selected_frame = ttk.Frame(app)
    selected_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

    ttk.Label(
        selected_frame,
        text="Selected Files Order",
        font=section_font,
        bootstyle="info",
    ).pack(anchor="w")

    selected_tree = Treeview(
        selected_frame,
        columns=("#1",),
        show="headings",
        height=5,
        style="Chosen.Treeview",
        bootstyle="warning",
    )
    selected_tree.heading("#1", text="Selected Files")
    selected_tree.column("#1", anchor="w")
    selected_tree.pack(fill=BOTH, expand=True)

    entry_frame = ttk.Frame(app)
    entry_frame.pack(fill=X, padx=10, pady=10)

    input_row = ttk.Frame(entry_frame)
    input_row.pack(anchor=CENTER)

    ttk.Label(
        input_row,
        text="Invoice Start:",
        font=ui_font,
    ).pack(side=LEFT)

    invoice_entry = ttk.Entry(input_row, font=ui_font, width=14)
    invoice_entry.pack(side=LEFT, padx=(6, 0))

    option_row = ttk.Frame(entry_frame)
    option_row.pack(anchor=CENTER, pady=(6, 0))

    auto_number_var = BooleanVar(value=False)
    ttk.Checkbutton(
        option_row,
        text="Auto number by MMYY",
        variable=auto_number_var,
        bootstyle="round-toggle",
    ).pack(side=LEFT)

    schedule_hint_var = StringVar(value="")
    ttk.Label(
        entry_frame,
        textvariable=schedule_hint_var,
        font=small_font,
        bootstyle="secondary",
    ).pack(anchor=CENTER, pady=(4, 0))

    file_pattern = re.compile(
        r"""^(?P<prefix>[A-Z0-9._ \-]+?)
            \s+xx26\s*[-\u2013\u2014]\s*00x
            (?:\s*[-\u2013\u2014]\s*DO\s*&\s*INV)?
            (?:\s*\((?P<name>.+)\))?
        """,
        re.IGNORECASE | re.VERBOSE,
    )

    wef_date_pattern = re.compile(
        r"\bWEF\b[^0-9]{0,12}(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b",
        re.IGNORECASE,
    )

    def get_reference_date() -> date:
        return date.today()

    def get_activation_cutoff(ref_date: date) -> date:
        # Weekday rule:
        # - Mon-Thu: allow next-day WEF.
        # - Fri: allow next workweek docs (up to +4 calendar days).
        # - Weekend: align to next Monday.
        if ref_date.weekday() == 4:
            return ref_date + timedelta(days=4)
        if ref_date.weekday() == 5:
            return ref_date + timedelta(days=2)
        if ref_date.weekday() == 6:
            return ref_date + timedelta(days=1)
        return ref_date + timedelta(days=1)

    def classify_wef_status(
        wef_date: date | None, ref_date: date, cutoff_date: date
    ) -> str:
        if wef_date is None:
            return "no-wef"
        if wef_date < ref_date:
            return "past"
        if wef_date <= cutoff_date:
            return "current"
        return "future"

    def update_schedule_hint():
        ref_date = get_reference_date()
        cutoff = get_activation_cutoff(ref_date)
        weekday = ref_date.weekday()
        if weekday == 4:
            week_text = "Fri +4 days"
        elif weekday in (5, 6):
            week_text = "Weekend -> next Monday"
        else:
            week_text = "Mon-Thu +1 day"
        schedule_hint_var.set(
            f"Auto schedule: today {ref_date:%d-%m-%Y}, CURRENT allowed until {cutoff:%d-%m-%Y} ({week_text})."
        )

    update_schedule_hint()

    def extract_wef_date_from_path(path):
        latest = None
        folder_path = os.path.dirname(path or "")
        for segment in os.path.normpath(folder_path).split(os.sep):
            if not segment:
                continue
            for m in wef_date_pattern.finditer(segment):
                day = int(m.group(1))
                month = int(m.group(2))
                year = int(m.group(3))
                if year < 100:
                    year += 2000
                try:
                    parsed = date(year, month, day)
                except ValueError:
                    continue
                if latest is None or parsed > latest:
                    latest = parsed
        return latest

    def get_next_invoice_number(search_dir, invoice_prefix):
        pattern = re.compile(
            rf"(?<!\d){re.escape(invoice_prefix)}\s*-\s*(\d{{3}})(?!\d)",
            re.IGNORECASE,
        )
        max_number = 0
        for root_dir, dirs, files in os.walk(search_dir):
            dirs[:] = [d for d in dirs if d.lower() != "history"]
            for name in files:
                for m in pattern.finditer(name):
                    number = int(m.group(1))
                    if number > max_number:
                        max_number = number
        return max_number + 1

    def resolve_month_folder(base_dir, invoice_prefix):
        month_value = int(invoice_prefix[:2])
        month_names = {
            1: "jan",
            2: "feb",
            3: "mar",
            4: "apr",
            5: "may",
            6: "jun",
            7: "jul",
            8: "aug",
            9: "sep",
            10: "oct",
            11: "nov",
            12: "dec",
        }
        month_token = month_names[month_value]
        month_pattern = re.compile(rf"^0?{month_value}\s*[\.\-_ ]", re.IGNORECASE)

        def is_target_month_folder(folder_name):
            lowered = folder_name.lower()
            return month_token in lowered and bool(month_pattern.search(lowered))

        current_name = os.path.basename(os.path.normpath(base_dir)).lower()
        if is_target_month_folder(current_name):
            return base_dir

        try:
            child_dirs = [
                os.path.join(base_dir, d)
                for d in os.listdir(base_dir)
                if os.path.isdir(os.path.join(base_dir, d))
            ]
        except OSError:
            return ""

        for folder_path in sorted(child_dirs):
            if is_target_month_folder(os.path.basename(folder_path)):
                return folder_path

        candidates = []
        for root_dir, dirs, _ in os.walk(base_dir):
            dirs[:] = [d for d in dirs if d.lower() != "history"]
            for folder_name in dirs:
                if not is_target_month_folder(folder_name):
                    continue
                folder_path = os.path.join(root_dir, folder_name)
                relative = os.path.relpath(folder_path, base_dir)
                depth = relative.count(os.sep)
                candidates.append((depth, relative.lower(), folder_path))

        if candidates:
            candidates.sort(key=lambda x: (x[0], x[1]))
            return candidates[0][2]

        return ""

    def ask_directory_quick(initial_dir):
        selected = {"path": ""}

        base_dir = os.path.abspath(initial_dir or os.getcwd())
        if not os.path.isdir(base_dir):
            base_dir = os.getcwd()

        dialog = ttk.Toplevel(app)
        dialog.title("Quick Folder Picker")
        dialog.geometry("860x560")
        dialog.minsize(720, 420)
        dialog.transient(app)
        dialog.grab_set()

        root_var = StringVar(value=base_dir)
        keyword_var = StringVar()
        status_var = StringVar(value="Waiting to build folder index...")
        indexed_dirs = []
        filtered_dirs = []

        top = ttk.Frame(dialog)
        top.pack(fill=X, padx=12, pady=(12, 8))
        ttk.Label(top, text="Index root folder:", font=small_font).pack(side=LEFT)
        root_entry = ttk.Entry(top, textvariable=root_var, width=72)
        root_entry.pack(side=LEFT, padx=(8, 8), fill=X, expand=True)

        def browse_root():
            picked = filedialog.askdirectory(
                initialdir=root_var.get().strip() or base_dir
            )
            if picked:
                root_var.set(picked)
                rebuild_index()

        ttk.Button(top, text="Change Root", command=browse_root, bootstyle=INFO).pack(
            side=LEFT
        )

        search_bar = ttk.Frame(dialog)
        search_bar.pack(fill=X, padx=12, pady=(0, 8))
        ttk.Label(search_bar, text="Keyword filter:", font=small_font).pack(side=LEFT)
        keyword_entry = ttk.Entry(search_bar, textvariable=keyword_var, width=42)
        keyword_entry.pack(side=LEFT, padx=(8, 8), fill=X, expand=True)

        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill=BOTH, expand=True, padx=12, pady=(0, 8))

        folder_list = ttk.Treeview(
            list_frame,
            columns=("folder", "relative"),
            show="headings",
            height=16,
            bootstyle="info",
        )
        folder_list.heading("folder", text="Folder")
        folder_list.heading("relative", text="Relative Path")
        folder_list.column("folder", width=260, anchor="w")
        folder_list.column("relative", width=560, anchor="w")
        folder_list.pack(side=LEFT, fill=BOTH, expand=True)

        y_scroll = ttk.Scrollbar(list_frame, orient=VERTICAL, command=folder_list.yview)
        y_scroll.pack(side=RIGHT, fill=Y)
        folder_list.configure(yscrollcommand=y_scroll.set)

        ttk.Label(dialog, textvariable=status_var, bootstyle="secondary").pack(
            anchor="w", padx=12, pady=(0, 8)
        )

        btn_row = ttk.Frame(dialog)
        btn_row.pack(fill=X, padx=12, pady=(0, 12))

        def get_selected_index():
            selected_item = folder_list.selection()
            if not selected_item:
                return None
            iid = selected_item[0]
            try:
                return int(folder_list.item(iid, "text"))
            except (TypeError, ValueError, TclError):
                return None

        def confirm_selection(_event=None):
            idx = get_selected_index()
            if idx is None or idx < 0 or idx >= len(filtered_dirs):
                messagebox.showwarning("Warning", "Please choose one folder first.")
                return
            selected["path"] = filtered_dirs[idx]["path"]
            dialog.destroy()

        def use_system_dialog():
            picked = filedialog.askdirectory(
                initialdir=root_var.get().strip() or base_dir
            )
            if picked:
                selected["path"] = picked
                dialog.destroy()

        def fill_list(rows):
            folder_list.delete(*folder_list.get_children())
            for idx, item in enumerate(rows):
                folder_name = os.path.basename(item["path"]) or item["path"]
                iid = folder_list.insert(
                    "",
                    END,
                    text=str(idx),
                    values=(folder_name, item["relative"]),
                )
                folder_list.item(
                    iid, tags=("oddrow",) if idx % 2 == 0 else ("evenrow",)
                )
            folder_list.tag_configure("oddrow", background="#2B2B2B")
            folder_list.tag_configure("evenrow", background="#242424")
            if rows:
                first = folder_list.get_children()[0]
                folder_list.selection_set(first)
                folder_list.focus(first)

        def move_selection(step):
            children = folder_list.get_children()
            if not children:
                return "break"
            current = folder_list.selection()
            current_idx = max(0, children.index(current[0])) if current else 0
            target_idx = max(0, min(len(children) - 1, current_idx + step))
            target = children[target_idx]
            folder_list.selection_set(target)
            folder_list.focus(target)
            folder_list.see(target)
            return "break"

        def apply_filter(*_):
            query = keyword_var.get().strip().lower()
            terms = [term for term in query.split() if term]
            filtered_dirs.clear()
            if not terms:
                filtered_dirs.extend(indexed_dirs)
            else:
                for item in indexed_dirs:
                    haystack = (
                        f"{item['relative']} {os.path.basename(item['path'])}".lower()
                    )
                    if all(term in haystack for term in terms):
                        filtered_dirs.append(item)
            fill_list(filtered_dirs)
            status_var.set(
                f"Indexed {len(indexed_dirs)} folder(s), matched {len(filtered_dirs)}."
            )

        def rebuild_index(_event=None):
            selected_root = os.path.abspath(root_var.get().strip() or base_dir)
            if not os.path.isdir(selected_root):
                messagebox.showwarning("Warning", "Root folder does not exist.")
                return
            status_var.set("Scanning folders, please wait...")
            dialog.update_idletasks()

            indexed_dirs.clear()
            indexed_dirs.append({"path": selected_root, "relative": "."})
            for root_dir, dirs, _ in os.walk(selected_root):
                dirs[:] = [d for d in dirs if d.lower() != "history"]
                for name in dirs:
                    full_path = os.path.join(root_dir, name)
                    relative = os.path.relpath(full_path, selected_root)
                    indexed_dirs.append({"path": full_path, "relative": relative})
            indexed_dirs.sort(key=lambda x: x["relative"].lower())
            apply_filter()

        ttk.Button(
            btn_row,
            text="System Picker",
            command=use_system_dialog,
            bootstyle="secondary",
        ).pack(side=LEFT)
        ttk.Button(
            btn_row, text="Cancel", command=dialog.destroy, bootstyle="warning"
        ).pack(side=RIGHT, padx=(8, 0))
        ttk.Button(
            btn_row, text="Confirm", command=confirm_selection, bootstyle="success"
        ).pack(side=RIGHT)

        keyword_var.trace_add("write", apply_filter)
        root_entry.bind("<Return>", rebuild_index)
        keyword_entry.bind("<Return>", confirm_selection)
        keyword_entry.bind("<Down>", lambda e: move_selection(1))
        keyword_entry.bind("<Up>", lambda e: move_selection(-1))
        folder_list.bind("<Double-1>", confirm_selection)
        folder_list.bind("<Return>", confirm_selection)
        folder_list.bind("<Escape>", lambda e: dialog.destroy())
        dialog.bind("<Escape>", lambda e: dialog.destroy())

        rebuild_index()
        keyword_entry.focus_set()
        app.wait_window(dialog)
        return selected["path"]

    def browse_source():
        source_dir = ask_directory_quick(last_directory[0])
        if not source_dir:
            return

        effective_date = get_reference_date()
        activation_cutoff = get_activation_cutoff(effective_date)
        update_schedule_hint()
        current_source_dir[0] = source_dir
        last_directory[0] = os.path.dirname(source_dir)

        with open("last_dir.txt", "w", encoding="utf-8") as f:
            f.write(source_dir)

        file_info_list.clear()
        file_tree.delete(*file_tree.get_children())
        selected_files.clear()
        selected_tree.delete(*selected_tree.get_children())

        idx = 0
        hidden_future_count = 0
        status_counts = {"past": 0, "current": 0, "future": 0, "no-wef": 0}
        for root_dir, _, files in os.walk(source_dir):
            if "history" in root_dir.lower():
                continue
            for name in files:
                full_path = os.path.join(root_dir, name)
                if not os.path.isfile(full_path):
                    continue
                filename_no_ext = os.path.splitext(name)[0]
                match = file_pattern.match(filename_no_ext)
                if not match:
                    continue

                wef_date = extract_wef_date_from_path(full_path)
                wef_status = classify_wef_status(
                    wef_date, effective_date, activation_cutoff
                )
                status_counts[wef_status] = status_counts.get(wef_status, 0) + 1
                if wef_status == "future":
                    hidden_future_count += 1
                    continue

                friendly_name = (match.group("name") or match.group("prefix")).strip()
                if wef_date:
                    wef_text = wef_date.strftime("%d-%m-%Y")
                    label = wef_status.upper()
                    friendly_name = f"{friendly_name} [WEF {wef_text} | {label}]"
                else:
                    friendly_name = f"{friendly_name} [NO WEF]"

                display = f"{len(file_info_list) + 1}. {friendly_name}"
                file_info_list.append(
                    {
                        "display_name": display,
                        "file_path": full_path,
                        "wef_date": wef_date,
                        "wef_status": wef_status,
                    }
                )
                tag = ("oddrow",) if idx % 2 == 0 else ("evenrow",)
                file_tree.insert("", END, values=(display,), tags=tag)
                idx += 1

        file_tree.tag_configure("oddrow", background="#2B2B2B")
        file_tree.tag_configure("evenrow", background="#242424")

        if not file_info_list:
            messagebox.showinfo("Info", "No matching files found.")
        elif hidden_future_count > 0:
            messagebox.showinfo(
                "Info",
                f"Hidden {hidden_future_count} FUTURE file(s). "
                f"Found PAST={status_counts.get('past', 0)}, CURRENT={status_counts.get('current', 0)}, "
                f"FUTURE={status_counts.get('future', 0)}.",
            )

    def add_to_selected():
        selected_item = file_tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select a file first.")
            return
        index = file_tree.index(selected_item)
        if index in [item["index"] for item in selected_files]:
            return
        selected_files.append({"index": index, "file_info": file_info_list[index]})
        update_selected_listbox()

    def update_selected_listbox():
        selected_tree.delete(*selected_tree.get_children())
        for i, item in enumerate(selected_files):
            tag = ("oddrow",) if i % 2 == 0 else ("evenrow",)
            selected_tree.insert(
                "",
                END,
                values=(f"{i + 1}. {item['file_info']['display_name']}",),
                tags=tag,
            )
        selected_tree.tag_configure("oddrow", background="#2B2B2B")
        selected_tree.tag_configure("evenrow", background="#242424")

    def delete_selected():
        selected_item = selected_tree.selection()
        if not selected_item:
            return
        index = selected_tree.index(selected_item)
        del selected_files[index]
        update_selected_listbox()

    def clear_selected():
        selected_files.clear()
        update_selected_listbox()

    def copy_files():
        if not target_dir:
            messagebox.showerror("Error", "Target folder is not configured.")
            return
        try:
            os.makedirs(target_dir, exist_ok=True)
        except OSError as exc:
            messagebox.showerror("Error", f"Cannot access target folder: {exc}")
            return
        if not selected_files:
            messagebox.showwarning("Warning", "Please choose file(s) to copy first.")
            return

        invoice_input = invoice_entry.get().strip()

        if auto_number_var.get():
            if not re.fullmatch(r"\d{4}", invoice_input):
                messagebox.showwarning(
                    "Warning", "Auto mode expects 4-digit prefix, example: 0226."
                )
                return
            invoice_prefix = invoice_input
            month_value = int(invoice_prefix[:2])
            if month_value < 1 or month_value > 12:
                messagebox.showwarning(
                    "Warning", "Invalid month in prefix. Use MMYY format."
                )
                return
            if not current_source_dir[0] or not os.path.isdir(current_source_dir[0]):
                messagebox.showwarning(
                    "Warning", "Please browse and choose source folder first."
                )
                return
            search_dir = resolve_month_folder(current_source_dir[0], invoice_prefix)
            if not search_dir:
                messagebox.showwarning(
                    "Warning",
                    f"Cannot find month folder for {invoice_prefix} under: {current_source_dir[0]}",
                )
                return
            invoice_number = get_next_invoice_number(search_dir, invoice_prefix)
        else:
            if not re.fullmatch(r"\d{4}\s*-\s*\d{3}", invoice_input):
                messagebox.showwarning(
                    "Warning",
                    "Invalid invoice start number. Example: 0326 - 001",
                )
                return
            invoice_prefix, invoice_number_text = invoice_input.split("-")
            invoice_prefix = invoice_prefix.strip()
            invoice_number = int(invoice_number_text.strip())

        if invoice_number > 999:
            messagebox.showerror(
                "Error",
                f"Prefix {invoice_prefix} already reached {invoice_number:03d}.",
            )
            return

        for item in selected_files:
            src_path = item["file_info"]["file_path"]
            filename = os.path.basename(src_path)
            new_filename = re.sub(
                r"xx26\s*[-\u2013\u2014]\s*00x",
                f"{invoice_prefix} - {invoice_number:03d}",
                filename,
                flags=re.IGNORECASE,
            )
            invoice_number += 1

            dst_path = os.path.join(target_dir, new_filename)
            if os.path.exists(dst_path):
                base_name, ext = os.path.splitext(new_filename)
                count = 1
                while os.path.exists(dst_path):
                    new_filename = f"{base_name}_{count}{ext}"
                    dst_path = os.path.join(target_dir, new_filename)
                    count += 1

            shutil.copy2(src_path, dst_path)

        messagebox.showinfo("Success", "Files copied and renamed.")

    file_tree.bind("<Double-1>", lambda e: add_to_selected())

    btn_frame = ttk.Frame(app)
    btn_frame.pack(pady=10)
    ttk.Button(
        btn_frame, text="Choose Source", command=browse_source, bootstyle=INFO
    ).grid(row=0, column=0, padx=10)
    ttk.Button(btn_frame, text="Add", command=add_to_selected, bootstyle=SUCCESS).grid(
        row=0, column=1, padx=10
    )
    ttk.Button(btn_frame, text="Clear", command=clear_selected, bootstyle=WARNING).grid(
        row=0, column=2, padx=10
    )
    ttk.Button(
        btn_frame, text="Delete", command=delete_selected, bootstyle=DANGER
    ).grid(row=0, column=3, padx=10)
    ttk.Button(btn_frame, text="Copy", command=copy_files, bootstyle=PRIMARY).grid(
        row=0, column=4, padx=10
    )
    ttk.Button(btn_frame, text="Exit", command=app.quit, bootstyle=SECONDARY).grid(
        row=0, column=5, padx=10
    )

    apply_scale(scale_var.get())

    try:
        app.state("zoomed")
    except TclError:
        try:
            app.attributes("-zoomed", True)
        except TclError:
            pass

    app.mainloop()


if __name__ == "__main__":
    main()
