import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .service import apply_insert, load_context, preview_insert, suggest


class DeliveryAssistantApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Delivery Details Entry Assistant")
        self.geometry("980x680")

        self.file_var = tk.StringVar()
        self.customer_var = tk.StringVar()
        self.outlet_var = tk.StringVar()
        self.desc_var = tk.StringVar()
        self.qty_pcs_var = tk.StringVar()
        self.qty_ctn_var = tk.StringVar()
        self.invoice_var = tk.StringVar()
        self._context = None
        self._last_plan = None

        self._build_ui()

    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)
        ttk.Label(top, text="Excel file:").pack(side=tk.LEFT)
        ttk.Entry(top, textvariable=self.file_var, width=80).pack(side=tk.LEFT, padx=6)
        ttk.Button(top, text="Browse...", command=self._browse).pack(side=tk.LEFT)

        form = ttk.LabelFrame(self, text="Input", padding=10)
        form.pack(fill=tk.X, padx=10, pady=8)
        fields = [
            ("Description", self.desc_var),
            ("Customer", self.customer_var),
            ("Outlet", self.outlet_var),
            ("Qty in Pcs", self.qty_pcs_var),
            ("Qty in Ctns", self.qty_ctn_var),
            ("Invoice #", self.invoice_var),
        ]
        for idx, (label, var) in enumerate(fields):
            ttk.Label(form, text=label).grid(row=idx, column=0, sticky=tk.W, pady=3)
            ttk.Entry(form, textvariable=var, width=45).grid(
                row=idx, column=1, sticky=tk.W, padx=6
            )

        ops = ttk.Frame(self, padding=(10, 2))
        ops.pack(fill=tk.X)
        ttk.Button(ops, text="Suggest", command=self._suggest).pack(side=tk.LEFT)
        ttk.Button(ops, text="Preview", command=self._preview).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(ops, text="Insert", command=self._insert).pack(side=tk.LEFT)

        self.suggest_tree = ttk.Treeview(
            self,
            columns=("date", "customer", "outlet", "score"),
            show="headings",
        )
        for col, text, width in (
            ("date", "Matched Date", 140),
            ("customer", "Customer", 250),
            ("outlet", "Outlet", 250),
            ("score", "Score", 100),
        ):
            self.suggest_tree.heading(col, text=text)
            self.suggest_tree.column(col, width=width)
        self.suggest_tree.pack(fill=tk.BOTH, expand=True, padx=10, pady=(4, 10))

    def _browse(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.file_var.set(path)

    def _suggest(self):
        if not self.file_var.get().strip():
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
        self._ensure_context()
        if self._context is None:
            return
        for item in self.suggest_tree.get_children():
            self.suggest_tree.delete(item)
        ranked = suggest(
            self._context,
            {
                "description": self.desc_var.get().strip(),
                "customer": self.customer_var.get().strip(),
                "outlet": self.outlet_var.get().strip(),
            },
            limit=8,
        )
        for rec in ranked:
            data = rec["record"]
            rec_date = data.get("record_date")
            date_text = rec_date.strftime("%d/%m/%Y") if rec_date else ""
            self.suggest_tree.insert(
                "",
                tk.END,
                values=(
                    date_text,
                    data.get("customer", ""),
                    data.get("outlet", ""),
                    f"{rec['score']:.1f}",
                ),
            )

    def _preview(self):
        if not self.file_var.get().strip():
            messagebox.showerror("Error", "Please select an Excel file first.")
            return
        self._ensure_context()
        if self._context is None:
            return
        try:
            self._last_plan = preview_insert(
                self._context,
                {
                    "description": self.desc_var.get().strip(),
                    "customer": self.customer_var.get().strip(),
                    "outlet": self.outlet_var.get().strip(),
                    "qty_pcs": int(self.qty_pcs_var.get().strip() or "0"),
                    "qty_ctns": int(self.qty_ctn_var.get().strip() or "0"),
                    "invoice": self.invoice_var.get().strip(),
                },
            )
        except ValueError:
            messagebox.showerror(
                "Error", "Qty in Pcs and Qty in Ctns must be integers."
            )
            return

        lines = [f"Insert row: {self._last_plan['insert_row']}"]
        for col_idx, value in sorted(self._last_plan["user_values"].items()):
            lines.append(f"Column {col_idx}: {value}")
        messagebox.showinfo("Preview", "\n".join(lines))

    def _insert(self):
        if self._last_plan is None:
            self._preview()
            if self._last_plan is None:
                return
        if not messagebox.askyesno("Confirm", "Insert this row into workbook?"):
            return
        try:
            backup_path = apply_insert(self._context, self._last_plan)
        except Exception as exc:
            messagebox.showerror("Insert failed", str(exc))
            return
        messagebox.showinfo("Done", f"Row inserted. Backup created:\n{backup_path}")

    def _ensure_context(self):
        file_path = self.file_var.get().strip()
        if self._context and self._context.get("file_path") == file_path:
            return
        try:
            self._context = load_context(file_path)
        except Exception as exc:
            self._context = None
            messagebox.showerror("Load failed", str(exc))


def launch():
    app = DeliveryAssistantApp()
    app.mainloop()
