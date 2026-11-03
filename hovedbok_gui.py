#!/usr/bin/env python3
"""
hovedbok_gui.py - a simple window for hovedbok_compare.py and bq_download.py.

Put this file in the same folder as hovedbok_compare.py and bq_download.py and
run:  python hovedbok_gui.py   (or double-click it)

Uses only tkinter, which comes with Python on Windows.
"""

import os
import queue
import subprocess
import sys
import json
import re
import threading
import time
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

HERE = Path(__file__).resolve().parent
COMPARE = HERE / "hovedbok_compare.py"
DOWNLOAD = HERE / "bq_download.py"
SETTINGS = Path.home() / ".hovedbok_compare_gui.json"  # the window's settings, remembered between sessions


class NotLoaded(Exception):
    pass


GUESS = "(guess)"


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel vs BigQuery comparison")
        w, h = 860, 760
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"{min(w, sw - 60)}x{min(h, sh - 100)}")
        self.minsize(560, 380)
        self.resizable(True, True)
        self.proc = None
        self.output = queue.Queue()
        self.task_name = None
        self.task_start = 0.0
        self.action_buttons = []
        self.last_lines = []
        self.loaded_key = None         # (file, sheet, header row) of the loaded copy
        self.loaded_cols = []
        self.loaded_original = []      # column numbers before naming
        self.loaded_samples = []

        self.old = tk.StringVar()
        self.sheet = tk.StringVar()
        self.start_row = tk.StringVar(value="2")
        self.old_columns = tk.StringVar()
        self.old_names = tk.StringVar()
        self.source = tk.StringVar(value="bq")
        self.bq_table = tk.StringVar(value="")
        self.new_file = tk.StringVar()
        self.bq_where = tk.StringVar()
        self.use_cache = tk.BooleanVar(value=True)
        self.workdir = tk.StringVar(value=str(Path.home() / "excel_bigquery_check"))
        self.decimals = tk.StringVar(value="2")
        self.casefold = tk.BooleanVar(value=False)
        self.new_types = tk.StringVar()
        self.loaded_text = tk.StringVar(value="Not loaded")
        self.roles = {r: tk.StringVar(value=GUESS) for r in ("date", "voucher", "account", "amount")}

        self.restore_settings()
        self.build()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        for var in (self.old, self.sheet, self.start_row, self.old_columns, self.old_names):
            var.trace_add("write", lambda *_: self.old_changed())
        self.after(100, self.drain_output)

    # ------------------------------------------------------------- layout ---
    def build(self):
        pad = {"padx": 6, "pady": 3}

        # Status bar first, packed at the bottom, so it always stays visible.
        bar = ttk.Frame(self, padding=(8, 2, 8, 6))
        bar.pack(side="bottom", fill="x")
        self.status_dot = tk.Label(bar, text="\u25cf", fg="green", font=("", 12))
        self.status_dot.pack(side="left")
        self.status = ttk.Label(bar, text="Ready")
        self.status.pack(side="left", padx=6)
        self.progress = ttk.Progressbar(bar, mode="indeterminate", length=200)
        self.progress.pack(side="right")

        # Form on top, messages below; drag the divider to change the split.
        paned = ttk.PanedWindow(self, orient="vertical")
        paned.pack(fill="both", expand=True, padx=4, pady=(4, 0))

        form_outer = ttk.Frame(paned)
        canvas = tk.Canvas(form_outer, highlightthickness=0, borderwidth=0)
        vscroll = ttk.Scrollbar(form_outer, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)
        vscroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        f = ttk.Frame(canvas, padding=8)
        f_id = canvas.create_window((0, 0), window=f, anchor="nw")
        f.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(f_id, width=e.width))

        def wheel(e):
            if f.winfo_height() > canvas.winfo_height():
                canvas.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units")
        form_outer.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", wheel))
        form_outer.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        paned.add(form_outer, weight=1)

        log_frame = ttk.Frame(paned)
        paned.add(log_frame, weight=1)
        def place_divider(tries=0):
            # Wait until the window has its real size, then give the form the room it needs
            # (at most 70% of the window) and the messages the rest.
            total = paned.winfo_height()
            if total < 200 and tries < 30:
                return self.after(100, place_divider, tries + 1)
            want = f.winfo_reqheight() + 6
            paned.sashpos(0, max(150, min(want, int(total * 0.7))))
        self.after_idle(place_divider)

        f.columnconfigure(1, weight=1)
        r = 0

        ttk.Label(f, text="Old table (Excel)", font=("", 10, "bold")).grid(row=r, column=0, sticky="w", **pad); r += 1
        ttk.Label(f, text="Excel file").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.old).grid(row=r, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Browse...", command=self.pick_old).grid(row=r, column=2, **pad); r += 1
        ttk.Label(f, text="Sheet").grid(row=r, column=0, sticky="w", **pad)
        self.sheet_box = ttk.Combobox(f, textvariable=self.sheet, values=[])
        self.sheet_box.grid(row=r, column=1, sticky="ew", **pad)
        b = ttk.Button(f, text="Load", command=self.load)
        b.grid(row=r, column=2, **pad); r += 1
        self.action_buttons.append(b)
        ttk.Label(f, text="Data starts at row").grid(row=r, column=0, sticky="w", **pad)
        hr = ttk.Frame(f)
        hr.grid(row=r, column=1, sticky="w", **pad)
        ttk.Entry(hr, textvariable=self.start_row, width=6).pack(side="left")
        ttk.Label(hr, text="  first row with data, below the column names", foreground="gray").pack(side="left"); r += 1
        ttk.Label(f, text="Columns").grid(row=r, column=0, sticky="w", **pad)
        cr = ttk.Frame(f)
        cr.grid(row=r, column=1, sticky="ew", **pad)
        ttk.Entry(cr, textvariable=self.old_columns, width=14).pack(side="left")
        ttk.Label(cr, text="   Names").pack(side="left")
        ttk.Entry(cr, textvariable=self.old_names).pack(side="left", fill="x", expand=True, padx=(4, 0))
        b = ttk.Button(f, text="Name columns...", command=self.name_columns)
        b.grid(row=r, column=2, **pad); r += 1
        self.action_buttons.append(b)
        ttk.Label(f, text="Columns: e.g. 1-8,10 or A-H,J (empty = all).  Names: optional, in the same order, "
                  "separated by ; (blank keeps the number)", foreground="gray").grid(row=r, column=1, sticky="w",
                                                                                   padx=6); r += 1
        ttk.Label(f, text="Loaded").grid(row=r, column=0, sticky="w", **pad)
        self.loaded_label = tk.Label(f, textvariable=self.loaded_text, fg="#b06000", anchor="w")
        self.loaded_label.grid(row=r, column=1, sticky="w", **pad); r += 1

        ttk.Separator(f).grid(row=r, column=0, columnspan=3, sticky="ew", pady=6); r += 1
        ttk.Label(f, text="New table (BigQuery)", font=("", 10, "bold")).grid(row=r, column=0, sticky="w", **pad); r += 1
        ttk.Radiobutton(f, text="Read BigQuery table", variable=self.source, value="bq").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.bq_table).grid(row=r, column=1, sticky="ew", **pad)
        b = ttk.Button(f, text="Download to CSV", command=self.download)
        b.grid(row=r, column=2, **pad); r += 1
        self.action_buttons.append(b)
        ttk.Label(f, text="    SQL filter (WHERE)").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.bq_where).grid(row=r, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Examples", command=self.where_help).grid(row=r, column=2, **pad); r += 1
        ttk.Label(f, text="optional, e.g.  dato BETWEEN '2025-01-01' AND '2025-12-31'",
                  foreground="gray").grid(row=r, column=1, sticky="w", padx=6); r += 1
        ttk.Checkbutton(f, text="Reuse the downloaded data between runs (cache)", variable=self.use_cache).grid(
            row=r, column=1, sticky="w", **pad); r += 1
        ttk.Radiobutton(f, text="Use a file (CSV/Excel)", variable=self.source, value="file").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.new_file).grid(row=r, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Browse...", command=self.pick_new).grid(row=r, column=2, **pad); r += 1
        ttk.Label(f, text="Column types").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.new_types).grid(row=r, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Types...", command=self.types_dialog).grid(row=r, column=2, **pad); r += 1
        ttk.Label(f, text="optional, e.g.  kode=STRING; belop=NUMERIC; dato=DATE  (empty = guess from the values)",
                  foreground="gray").grid(row=r, column=1, sticky="w", padx=6); r += 1

        ttk.Separator(f).grid(row=r, column=0, columnspan=3, sticky="ew", pady=6); r += 1
        ttk.Label(f, text="Output folder").grid(row=r, column=0, sticky="w", **pad)
        ttk.Entry(f, textvariable=self.workdir).grid(row=r, column=1, sticky="ew", **pad)
        ttk.Button(f, text="Browse...", command=self.pick_workdir).grid(row=r, column=2, **pad); r += 1
        opts = ttk.Frame(f)
        opts.grid(row=r, column=1, sticky="w", **pad)
        ttk.Label(opts, text="Decimals").pack(side="left")
        ttk.Entry(opts, textvariable=self.decimals, width=4).pack(side="left", padx=(4, 16))
        ttk.Checkbutton(opts, text="Ignore upper/lower case", variable=self.casefold).pack(side="left"); r += 1


        btns = ttk.Frame(f)
        btns.grid(row=r, column=0, columnspan=3, sticky="ew", pady=8)
        groups = (
            ("1. Check the data", (("Preview both tables", self.preview_data, True),)),
            ("2. Match columns", (("Suggest automatically", self.suggest, True),
                                  ("Edit by hand...", self.map_columns, True),
                                  ("Check values", self.check_values, True),
                                  ("Open mapping file", lambda: self.open_path(self.mapping_path()), False))),
            ("3. Compare", (("Column roles...", self.roles_dialog, False),
                            ("Run comparison", self.compare, True),
                            ("Open Excel report", lambda: self.open_path(self.results_dir() / "results.xlsx"), False),
                            ("Open results folder", lambda: self.open_path(self.results_dir()), False))),
            ("Tools", (("Test the program", self.selftest, True),
                       ("Save settings for the command line", self.export_config, False))),
        )
        for i, (title, buttons) in enumerate(groups):
            box = ttk.LabelFrame(btns, text=title)
            box.grid(row=i // 2, column=i % 2, padx=(0, 8), pady=(0, 4), sticky="nsw")
            for text, cmd, busy in buttons:
                b = ttk.Button(box, text=text, command=cmd)
                b.pack(side="left", padx=4, pady=4)
                if busy:
                    self.action_buttons.append(b)
        self.stop_button = ttk.Button(btns, text="Stop", command=self.stop, state="disabled")
        self.stop_button.grid(row=1, column=2, padx=4, pady=(14, 4), sticky="se")
        btns.columnconfigure(2, weight=1); r += 1

        self.log = scrolledtext.ScrolledText(log_frame, height=8, font=("Consolas", 9), wrap="word")
        self.log.pack(fill="both", expand=True, padx=6, pady=(0, 2))
        self.write("1. Pick the Excel file and sheet, set the first data row and the columns, and press Load.\n"
                   "   The 'Loaded' line shows the rows and columns that were read; the first values of each\n"
                   "   column are listed here so you can check them.\n"
                   "   Optional: give the columns names (Names field or Name columns...) so they are easier to map.\n"
                   "2. Check the BigQuery table and filter.\n"
                   "3. Use the buttons at the bottom from left to right: check the data, match the columns\n"
                   "   (automatically, by hand, or both), then run the comparison.\n")

    # ----------------------------------------------------------- settings ---
    def setting_vars(self):
        v = {"old": self.old, "sheet": self.sheet, "start_row": self.start_row, "old_columns": self.old_columns,
             "old_names": self.old_names, "source": self.source, "bq_table": self.bq_table,
             "new_file": self.new_file, "bq_where": self.bq_where, "use_cache": self.use_cache,
             "workdir": self.workdir, "decimals": self.decimals, "casefold": self.casefold,
             "new_types": self.new_types}
        v.update({f"role_{r}": var for r, var in self.roles.items()})
        return v

    def restore_settings(self):
        try:
            saved = json.loads(SETTINGS.read_text(encoding="utf-8"))
        except (OSError, ValueError):
            return
        for name, var in self.setting_vars().items():
            if name in saved:
                try:
                    var.set(saved[name])
                except tk.TclError:
                    pass

    def save_settings(self):
        try:
            SETTINGS.write_text(json.dumps({n: v.get() for n, v in self.setting_vars().items()},
                                           ensure_ascii=False, indent=1), encoding="utf-8")
        except OSError:
            pass

    def on_close(self):
        self.save_settings()
        self.destroy()

    def export_config(self):
        """Write settings.json for hovedbok_compare.py --config, so the same check can run without the window."""
        try:
            self.old_args()
        except ValueError as e:
            return messagebox.showwarning("Missing input", str(e))
        cfg = {"old": self.old.get(), "old-start-row": int(self.start_row.get().strip() or "2")}
        if self.sheet.get():
            cfg["old-sheet"] = self.sheet.get()
        for key, var in (("old-columns", self.old_columns), ("old-names", self.old_names),
                         ("new-types", self.new_types)):
            if var.get().strip():
                cfg[key] = var.get().strip()
        if self.source.get() == "bq":
            cfg["new-bq"] = self.bq_table.get().strip()
            if self.bq_where.get().strip():
                cfg["new-where"] = self.clean_where()
            if self.use_cache.get():
                cfg["cache"] = str(self.work() / "bq_cache.pkl")
        else:
            cfg["new"] = self.new_file.get()
        cfg["decimals"] = int(self.decimals.get().strip() or "2")
        if self.casefold.get():
            cfg["casefold"] = True
        cfg["mapping"] = str(self.mapping_path())
        cfg["out-dir"] = str(self.results_dir())
        mapped = {oc for oc, nc in self.read_mapping_file().items() if nc} if self.mapping_path().exists() else set()
        for role, var in self.roles.items():
            if var.get() != GUESS and var.get() in mapped:
                cfg[f"{role}-col"] = var.get()
        path = self.work() / "settings.json"
        path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        self.write(f"\nSaved {path}. Run the same check without this window with:\n"
                   f'  python "{COMPARE}" run --config "{path}"\n'
                   f"(add --json for output that other programs can read; see CLI_README.md)\n")

    def selftest(self):
        self.run([COMPARE, "selftest"], name="Testing the program")

    # -------------------------------------------------------------- paths ---
    def work(self):
        p = Path(self.workdir.get().strip() or ".")
        p.mkdir(parents=True, exist_ok=True)
        return p

    def mapping_path(self):
        return self.work() / "mapping.csv"

    def results_dir(self):
        return self.work() / "results"

    def pick_old(self):
        path = filedialog.askopenfilename(filetypes=[("Excel or CSV", "*.xlsx *.xlsm *.xls *.csv"), ("All files", "*.*")])
        if not path:
            return
        self.old.set(path)
        try:
            import openpyxl
            wb = openpyxl.load_workbook(path, read_only=True)
            self.sheet_box["values"] = wb.sheetnames
            self.sheet.set(wb.sheetnames[0])
            wb.close()
        except PermissionError:
            self.sheet_box["values"] = []
            self.write("Can't open the file to list its sheets (permission denied). It is probably open in Excel or "
                       "still syncing in OneDrive. Close it in Excel, then press Load.\n")
            return
        except Exception as e:
            self.sheet_box["values"] = []
            self.write(f"Couldn't list the sheets ({e}). Type the sheet name and press Load.\n")
            return
        self.write("Check the sheet, the first data row and the columns, then press Load.\n")

    def pick_new(self):
        path = filedialog.askopenfilename(filetypes=[("CSV or Excel", "*.csv *.xlsx *.xls"), ("All files", "*.*")])
        if path:
            self.new_file.set(path)
            self.source.set("file")

    def pick_workdir(self):
        path = filedialog.askdirectory()
        if path:
            self.workdir.set(path)

    # ----------------------------------------------------------- commands ---
    def old_args(self):
        if not self.old.get():
            raise ValueError("Pick the old Excel file first.")
        args = ["--old", self.old.get()]
        if self.sheet.get():
            args += ["--old-sheet", self.sheet.get()]
        start = self.start_row.get().strip() or "2"
        if not start.isdigit() or int(start) < 1:
            raise ValueError("'Data starts at row' must be a row number, e.g. 2.")
        args += ["--old-start-row", start]
        if self.old_columns.get().strip():
            if not re.fullmatch(r"[0-9A-Za-z,\-\s]+", self.old_columns.get()):
                raise ValueError("Columns: use numbers or letters with commas and dashes, e.g. 1-8,10 or A-H,J.")
            args += ["--old-columns", self.old_columns.get().strip()]
        if self.old_names.get().strip():
            args += ["--old-names", self.old_names.get().strip()]
        return args

    def old_cache(self):
        return self.work() / "old_cache.pkl"

    def current_key(self):
        return (self.old.get(), self.sheet.get(), self.start_row.get().strip(), self.old_columns.get().strip(),
                self.old_names.get().strip())

    def is_loaded(self):
        return self.loaded_key == self.current_key() and self.old_cache().exists()

    def old_changed(self):
        if self.loaded_key is not None and self.loaded_key != self.current_key():
            self.loaded_text.set("Not loaded (settings changed): press Load")
            self.loaded_label.config(fg="#b06000")

    def load(self, then=None):
        try:
            args = ["list-columns"] + self.old_args() + ["--save", str(self.old_cache())]
        except ValueError as e:
            return messagebox.showwarning("Missing input", str(e))
        if self.proc and self.proc.poll() is None:
            return messagebox.showinfo("Busy", "Something is already running. Wait or press Stop.")
        key = self.current_key()
        self.loaded_key = None
        self.loaded_text.set("Loading ...")
        self.loaded_label.config(fg="#b06000")

        def done(ok):
            rows, cols = None, None
            self.loaded_original, self.loaded_samples = [], []
            for line in self.last_lines:
                if line.startswith("ROWS:"):
                    rows = int(line[5:])
                elif line.startswith("COLUMNS_JSON:"):
                    cols = json.loads(line[len("COLUMNS_JSON:"):])
                elif line.startswith("ORIGINAL_JSON:"):
                    self.loaded_original = json.loads(line[len("ORIGINAL_JSON:"):])
                elif line.startswith("SAMPLES_JSON:"):
                    self.loaded_samples = json.loads(line[len("SAMPLES_JSON:"):])
            if not ok or cols is None:
                self.loaded_text.set("Loading failed: see the messages below")
                self.loaded_label.config(fg="red")
                return
            self.loaded_key, self.loaded_cols = key, cols
            warn = False
            text = f"{rows:,} rows".replace(",", " ") + f", {len(cols)} columns: {', '.join(cols[:6])}{' ...' if len(cols) > 6 else ''}"
            self.loaded_text.set(text)
            self.loaded_label.config(fg="red" if warn else "green")
            if then:
                then()
        self.run([COMPARE] + args, name="Loading Excel file", on_done=done)

    def common_args(self):
        if not self.is_loaded():
            raise NotLoaded()
        args = ["--old", str(self.old_cache())]
        if self.source.get() == "bq":
            if not self.bq_table.get().strip():
                raise ValueError("Enter the BigQuery table as project.dataset.table.")
            args += ["--new-bq", self.bq_table.get().strip()]
            if self.bq_where.get().strip():
                args += ["--new-where", self.clean_where()]
        else:
            if not self.new_file.get():
                raise ValueError("Pick the new file, or choose 'Read BigQuery table'.")
            args += ["--new", self.new_file.get()]
        args += ["--decimals", self.decimals.get().strip() or "2"]
        if self.casefold.get():
            args.append("--casefold")
        return args

    def suggest(self):
        try:
            args = ["suggest"] + self.common_args() + ["--out", str(self.mapping_path())] + self.types_args()
        except NotLoaded:
            return self.load(then=self.suggest)
        except ValueError as e:
            return messagebox.showwarning("Missing input", str(e))
        self.run([COMPARE] + args, name="Suggesting mapping", on_done=lambda ok: ok and self.write(
            "\nCheck the suggestion with Edit by hand... (or Open mapping file), then press Run comparison.\n"))

    # ------------------------------------------------------------ preview ---
    def fetch_preview(self, then, retry):
        """Read columns, sample values and the top 10 rows of both tables, then call then(data)."""
        try:
            args = ["preview"] + self.common_args() + self.types_args()
        except NotLoaded:
            return self.load(then=retry)
        except ValueError as e:
            return messagebox.showwarning("Missing input", str(e))
        if self.source.get() == "bq" and self.use_cache.get():
            args += ["--cache", str(self.work() / "bq_cache.pkl")]

        def done(ok):
            data = None
            for line in self.last_lines:
                if line.startswith("PREVIEW_JSON:"):
                    data = json.loads(line[len("PREVIEW_JSON:"):])
            if ok and data:
                then(data)
        self.run([COMPARE] + args, name="Reading both tables", on_done=done)

    def preview_data(self):
        def show(data):
            win = tk.Toplevel(self)
            win.title("Preview: first non-empty values per column")
            win.geometry("1000x560")
            win.minsize(500, 300)
            self.preview_tables(win, data).pack(fill="both", expand=True, padx=8, pady=8)
            ttk.Button(win, text="Close", command=win.destroy).pack(anchor="e", padx=8, pady=(0, 8))
        self.fetch_preview(show, self.preview_data)

    def preview_tables(self, parent, data):
        """Excel table above, BigQuery table below, each with scrollbars."""
        frame = ttk.PanedWindow(parent, orient="vertical")
        for key, title in (("old_top", "Excel (old table): first 10 non-empty values of each column"),
                           ("new_top", "BigQuery / file (new table): first 10 non-empty values of each column")):
            part = data.get(key) or {"columns": [], "rows": []}
            box = ttk.Frame(frame)
            head = ttk.Frame(box)
            head.grid(row=0, column=0, sticky="w", pady=(4, 2))
            ttk.Label(head, text=title, font=("", 9, "bold")).pack(side="left")
            ttk.Label(head, text="   empty cells are skipped, so a line is not one row; values shown as compared",
                      foreground="gray").pack(side="left")
            cols = part["columns"]
            ids = [f"c{i}" for i in range(len(cols))]
            tree = ttk.Treeview(box, columns=ids, show="headings", height=10)
            for cid, name in zip(ids, cols):
                tree.heading(cid, text=self.col_label(name) if key == "old_top" else name)
                width = max([len(self.col_label(name))] + [len(r[ids.index(cid)]) for r in part["rows"]] + [6])
                tree.column(cid, width=min(240, 8 * width + 12), minwidth=40, stretch=False, anchor="w")
            for row in part["rows"]:
                tree.insert("", "end", values=row)
            ys = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
            xs = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
            tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
            tree.grid(row=1, column=0, sticky="nsew")
            ys.grid(row=1, column=1, sticky="ns")
            xs.grid(row=2, column=0, sticky="ew")
            box.rowconfigure(1, weight=1)
            box.columnconfigure(0, weight=1)
            frame.add(box, weight=1)
        return frame

    @staticmethod
    def col_label(col):
        try:
            n = int(col)
        except ValueError:
            return col
        letters = ""
        while n:
            n, r = divmod(n - 1, 26)
            letters = chr(65 + r) + letters
        return f"{col} ({letters})"

    # ------------------------------------------------------- column names ---
    def name_columns(self):
        """One entry per loaded Excel column, with its first values; fills the Names field."""
        if not self.is_loaded():
            return self.load(then=self.name_columns)
        numbers = self.loaded_original or self.loaded_cols
        current = [n.strip() for n in self.old_names.get().split(";" if ";" in self.old_names.get() else ",")] \
            if self.old_names.get().strip() else []
        win = tk.Toplevel(self)
        win.title("Name the Excel columns")
        win.geometry("760x600")
        win.minsize(500, 300)
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="Type a name for each column (blank keeps the number). The names are used in the "
                  "mapping and the results.", wraplength=720).pack(anchor="w", padx=8, pady=(8, 4))
        outer = ttk.Frame(win)
        outer.pack(fill="both", expand=True, padx=8)
        canvas = tk.Canvas(outer, highlightthickness=0)
        sb = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        inner = ttk.Frame(canvas)
        inner.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        canvas.bind("<Enter>", lambda e: canvas.bind_all(
            "<MouseWheel>", lambda ev: canvas.yview_scroll(int(-ev.delta / 120), "units")))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))
        ttk.Label(inner, text="Column", font=("", 9, "bold")).grid(row=0, column=0, sticky="w", padx=4)
        ttk.Label(inner, text="Name", font=("", 9, "bold")).grid(row=0, column=1, sticky="w", padx=4)
        ttk.Label(inner, text="First values", font=("", 9, "bold")).grid(row=0, column=2, sticky="w", padx=4)
        entries = []
        for i, num in enumerate(numbers):
            var = tk.StringVar(value=current[i] if i < len(current) else "")
            ttk.Label(inner, text=self.col_label(num)).grid(row=i + 1, column=0, sticky="w", padx=4, pady=1)
            ttk.Entry(inner, textvariable=var, width=24).grid(row=i + 1, column=1, sticky="w", padx=4, pady=1)
            sample = self.loaded_samples[i] if i < len(self.loaded_samples) else ""
            ttk.Label(inner, text=sample[:60], foreground="gray").grid(row=i + 1, column=2, sticky="w", padx=4)
            entries.append(var)

        def close():
            canvas.unbind_all("<MouseWheel>")
            win.grab_release()
            win.destroy()

        def save():
            names = [v.get().strip().replace(";", " ") for v in entries]
            used = [n for n in names if n]
            dupes = sorted({n for n in used if used.count(n) > 1})
            if dupes:
                return messagebox.showwarning("Duplicate names", "Each name can only be used once: "
                                              + ", ".join(dupes), parent=win)
            while names and not names[-1]:
                names.pop()
            close()
            self.old_names.set("; ".join(names) + (";" if len(names) == 1 else ""))
            self.write("Column names set. Loading again with the new names ...\n"
                       "(A mapping made before renaming must be made again.)\n")
            self.load()

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=8)
        ttk.Button(bar, text="Clear all", command=lambda: [v.set("") for v in entries]).pack(side="left")
        ttk.Button(bar, text="Cancel", command=close).pack(side="right")
        ttk.Button(bar, text="Save names and load", command=save).pack(side="right", padx=4)
        win.protocol("WM_DELETE_WINDOW", close)

    # ------------------------------------------------------ manual mapping ---
    def map_columns(self):
        self.fetch_preview(lambda data: self.mapping_dialog(data["old"], data["new"], data), self.map_columns)

    def read_mapping_file(self):
        import csv
        path = self.mapping_path()
        if not path.exists():
            return {}
        with open(path, encoding="utf-8-sig", newline="") as fh:
            text = fh.read()
        delim = ";" if text.count(";") >= text.count(",") else ","
        return {r["old_column"]: (r.get("new_column") or "").strip()
                for r in csv.DictReader(text.splitlines(), delimiter=delim)}

    def mapping_dialog(self, old, new, data=None):
        NONE = "(not compared)"
        win = tk.Toplevel(self)
        win.title("Map columns: Excel -> BigQuery")
        win.geometry("1260x820")
        win.minsize(600, 400)
        win.transient(self)
        win.grab_set()

        split = ttk.PanedWindow(win, orient="vertical")
        split.pack(fill="both", expand=True, padx=8, pady=(6, 0))
        if data:
            split.add(self.preview_tables(split, data), weight=1)
        lower = ttk.Frame(split)
        split.add(lower, weight=1)
        ttk.Label(lower, text="For each Excel column, pick the matching BigQuery column, or '(not compared)'. "
                              "Match: 'values' = share of the Excel column's values found in the BigQuery column "
                              "(green 90%+, orange 50%+, red below); 'name' = name similarity; score = 60% values + "
                              "30% name + 10% similar number of different values, as used by Suggest.", wraplength=1200).pack(anchor="w", pady=6)

        outer = ttk.Frame(lower)
        outer.pack(fill="both", expand=True)
        canvas = tk.Canvas(outer, highlightthickness=0)
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        grid = ttk.Frame(canvas)
        grid.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        gid = canvas.create_window((0, 0), window=grid, anchor="nw")
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(gid, width=e.width))
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        def wheel(e):
            if str(e.widget).startswith(str(canvas)):  # leave the preview tables' own scrolling alone
                canvas.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units")
        canvas.bind_all("<MouseWheel>", wheel)
        grid.columnconfigure(1, weight=1)
        grid.columnconfigure(4, weight=1)
        scores = (data or {}).get("scores", {})

        for c, text in enumerate(("Excel column", "Excel values", "BigQuery column", "Match score", "BigQuery values")):
            ttk.Label(grid, text=text, font=("", 9, "bold")).grid(row=0, column=c, sticky="w", padx=4, pady=2)

        def label(col):
            try:
                n = int(col)
                letters = ""
                while n:
                    n, r = divmod(n - 1, 26)
                    letters = chr(65 + r) + letters
                return f"{col} ({letters})"
            except ValueError:
                return col

        existing = self.read_mapping_file()
        choices = [NONE] + list(new)
        boxes = {}
        for i, oc in enumerate(old, start=1):
            ttk.Label(grid, text=label(oc)).grid(row=i, column=0, sticky="w", padx=4, pady=1)
            ttk.Label(grid, text=" | ".join(old[oc]), foreground="gray").grid(row=i, column=1, sticky="w", padx=4)
            var = tk.StringVar(value=existing.get(oc) if existing.get(oc) in new else NONE)
            box = ttk.Combobox(grid, textvariable=var, values=choices, state="readonly", width=28)
            box.grid(row=i, column=2, sticky="w", padx=4, pady=1)
            match = tk.Label(grid, text="", anchor="w")
            match.grid(row=i, column=3, sticky="w", padx=4)
            sample = ttk.Label(grid, text="", foreground="gray")
            sample.grid(row=i, column=4, sticky="w", padx=4)

            def show(*_, var=var, sample=sample, match=match, oc=oc):
                sample.config(text=" | ".join(new.get(var.get(), [])))
                row = scores.get(str(oc), {})
                best = max(row.items(), key=lambda kv: kv[1][0]) if row else None
                if var.get() == NONE or var.get() not in row:
                    text = f"best: {best[0]} {best[1][0]:.0%}" if best and best[1][0] >= 0.5 else ""
                    match.config(text=text, fg="gray")
                    return
                score, overlap, name_sim = row[var.get()]
                text = f"values {overlap:.0%}, name {name_sim:.0%}  (score {score:.0%})"
                if best and best[0] != var.get() and best[1][0] > score + 0.05:
                    text += f"   better: {best[0]} {best[1][0]:.0%}"
                match.config(text=text, fg="#1a7f37" if overlap >= 0.9 else "#b06000" if overlap >= 0.5 else "#c62828")
            var.trace_add("write", show)
            show()
            boxes[oc] = var

        def close():
            canvas.unbind_all("<MouseWheel>")
            win.destroy()

        def clear():
            for v in boxes.values():
                v.set(NONE)

        def save(check=False):
            chosen = {oc: v.get() for oc, v in boxes.items() if v.get() != NONE}
            if not chosen:
                return messagebox.showwarning("Nothing mapped", "Pick a BigQuery column for at least one Excel column.",
                                              parent=win)
            used = {}
            for oc, nc in chosen.items():
                used.setdefault(nc, []).append(oc)
            dupes = {nc: ocs for nc, ocs in used.items() if len(ocs) > 1}
            if dupes and not messagebox.askyesno(
                    "Same BigQuery column twice",
                    "These BigQuery columns are picked more than once:\n"
                    + "\n".join(f"{nc}: Excel {', '.join(ocs)}" for nc, ocs in dupes.items())
                    + "\n\nSave anyway?", parent=win):
                return
            import csv
            with open(self.mapping_path(), "w", newline="", encoding="utf-8-sig") as fh:
                w = csv.writer(fh, delimiter=";")
                w.writerow(["old_column", "new_column", "name_similarity", "value_overlap", "old_distinct_values",
                            "confidence"])
                for oc in old:
                    w.writerow([oc, chosen.get(oc, ""), "", "", "", "manual" if oc in chosen else ""])
            self.write(f"Saved mapping with {len(chosen)} column pair(s) to {self.mapping_path()}\n"
                       + "".join(f"  {label(oc)} -> {nc}\n" for oc, nc in chosen.items())
                       + "Press Run comparison when ready.\n")
            close()
            if check:
                self.check_values()

        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=8)
        ttk.Button(bar, text="Clear all", command=clear).pack(side="left")
        ttk.Label(bar, text="  Starts from the current mapping (e.g. from Suggest), if there is one.",
                  foreground="gray").pack(side="left")
        ttk.Button(bar, text="Cancel", command=close).pack(side="right")
        ttk.Button(bar, text="Save mapping", command=save).pack(side="right", padx=4)
        ttk.Button(bar, text="Save and check values", command=lambda: save(check=True)).pack(side="right", padx=4)
        win.protocol("WM_DELETE_WINDOW", close)

    # -------------------------------------------------------- value check ---
    def check_values(self):
        """For each matched pair, how many of the Excel column's non-empty values exist in the BigQuery column."""
        if not self.mapping_path().exists():
            return messagebox.showwarning("No mapping", "Match the columns first: press Suggest automatically "
                                          "or Edit by hand...")
        try:
            args = ["check-values"] + self.common_args() + ["--mapping", str(self.mapping_path()),
                                                            "--out-dir", str(self.results_dir())] + self.types_args()
        except NotLoaded:
            return self.load(then=self.check_values)
        except ValueError as e:
            return messagebox.showwarning("Missing input", str(e))
        if self.source.get() == "bq" and self.use_cache.get():
            args += ["--cache", str(self.work() / "bq_cache.pkl")]

        def done(ok):
            notin = None
            for line in self.last_lines:
                if line.startswith("NOTIN_FILE:"):
                    notin = line[len("NOTIN_FILE:"):]
            for line in self.last_lines:
                if line.startswith("CHECK_JSON:"):
                    self.check_window(json.loads(line[len("CHECK_JSON:"):]), notin)
        self.run([COMPARE] + args, name="Checking values", on_done=done)

    def check_window(self, results, notin=None):
        win = tk.Toplevel(self)
        win.title("Check values: are the Excel values in the matched BigQuery column?")
        win.geometry("1240x440")
        win.minsize(600, 250)
        ttk.Label(win, text="For each matched pair: how many of the different non-empty Excel values also occur "
                  "in the BigQuery column. Green is 99% or more, yellow 80% or more, red below. "
                  "Double-click a line to see every Excel value that is not in the BigQuery column.",
                  wraplength=1000).pack(anchor="w", padx=8, pady=(8, 4))
        box = ttk.Frame(win)
        box.pack(fill="both", expand=True, padx=8)
        cols = (("old", "Excel column", 130), ("new", "BigQuery column", 160), ("found", "Values found", 110),
                ("pct", "% found", 70), ("rows", "Excel rows not found", 130),
                ("cause", "Likely cause of the values not found", 330),
                ("ex", "Excel values not found (examples)", 220), ("nex", "BigQuery values look like", 180))
        tree = ttk.Treeview(box, columns=[c for c, _, _ in cols], show="headings")
        for cid, text, width in cols:
            tree.heading(cid, text=text)
            tree.column(cid, width=width, minwidth=50, anchor="w", stretch=cid in ("cause", "ex", "nex"))
        tree.tag_configure("good", background="#dff3df")
        tree.tag_configure("check", background="#fff3cd")
        tree.tag_configure("bad", background="#f8d7da")
        for r in results:
            pct = r["pct"]
            if pct is None:
                tag, pct_text, found = "", "", "Excel column empty"
            else:
                tag = "good" if pct >= 99 else "check" if pct >= 80 else "bad"
                pct_text, found = f"{pct:.1f}", f"{r['found']} of {r['distinct']}"
            tree.insert("", "end", tags=(tag,), values=(
                self.col_label(r["old"]), r["new"], found, pct_text, f"{r['rows_not_found']} of {r['rows']}",
                "; ".join(f"{c['rows']} {c['text']}" + (f" ('{c['old']}' vs '{c['new']}')" if c["new"] else "")
                          for c in r.get("causes", [])),
                " | ".join(r["examples"]), " | ".join(r["new_examples"])))
        ys = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        xs = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=8)
        ttk.Label(bar, text="Likely cause: 'different value' is a genuine difference; the others (case, rounding, "
                  "sign, date one day off ...) are format problems. Saved as value_check.csv in the results folder.",
                  foreground="gray", wraplength=900).pack(side="left")
        ttk.Button(bar, text="Close", command=win.destroy).pack(side="right")
        if notin:
            ttk.Button(bar, text="Open full list", command=lambda: self.open_path(Path(notin))).pack(
                side="right", padx=4)

            def on_double(_e):
                sel = tree.selection()
                if sel:
                    self.not_in_window(results[tree.index(sel[0])]["old"], notin)
            tree.bind("<Double-1>", on_double)
        ttk.Button(bar, text="Edit mapping...", command=lambda: (win.destroy(), self.map_columns())).pack(
            side="right", padx=4)

    def not_in_window(self, column, path):
        """Every value of one Excel column that the BigQuery column lacks, with the likely cause."""
        import csv
        with open(path, encoding="utf-8-sig", newline="") as fh:
            rows = [r for r in csv.reader(fh, delimiter=";")]
        head, rows = rows[0], [r for r in rows[1:] if r and r[0] == column]
        win = tk.Toplevel(self)
        win.title(f"Values in Excel column {self.col_label(column)} that are not in BigQuery")
        win.geometry("1100x520")
        win.minsize(500, 250)
        total = sum(int(r[4]) for r in rows) if rows else 0
        ttk.Label(win, text=f"{len(rows)} different values, in {total} Excel rows. REAL = probably genuinely absent; "
                  "FORMAT/CHECK = probably the same value written differently. 'Closest' is the nearest BigQuery "
                  "value, to help you judge.", wraplength=1060).pack(anchor="w", padx=8, pady=(8, 4))
        box = ttk.Frame(win)
        box.pack(fill="both", expand=True, padx=8)
        show = [(2, "Value (as compared)", 150), (3, "Value in Excel", 150), (4, "Rows", 50),
                (5, "Excel rows", 130), (6, "Category", 80), (7, "Likely cause", 260), (8, "Closest BigQuery value", 170)]
        tree = ttk.Treeview(box, columns=[str(i) for i, _, _ in show], show="headings")
        for i, text, width in show:
            tree.heading(str(i), text=text if i != 5 else head[5].replace("_", " "))
            tree.column(str(i), width=width, minwidth=40, anchor="w", stretch=i in (7, 8))
        tree.tag_configure("FORMAT", background="#fff3cd")
        tree.tag_configure("CHECK", background="#fff3cd")
        tree.tag_configure("REAL", background="#f8d7da")
        for r in rows[:5000]:
            tree.insert("", "end", tags=(r[6],), values=[r[i] for i, _, _ in show])
        ys = ttk.Scrollbar(box, orient="vertical", command=tree.yview)
        xs = ttk.Scrollbar(box, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ys.grid(row=0, column=1, sticky="ns")
        xs.grid(row=1, column=0, sticky="ew")
        box.rowconfigure(0, weight=1)
        box.columnconfigure(0, weight=1)
        bar = ttk.Frame(win)
        bar.pack(fill="x", padx=8, pady=8)
        if len(rows) > 5000:
            ttk.Label(bar, text=f"Showing the first 5000; open the full list for all {len(rows)}.",
                      foreground="gray").pack(side="left")
        ttk.Button(bar, text="Close", command=win.destroy).pack(side="right")
        ttk.Button(bar, text="Open full list", command=lambda: self.open_path(Path(path))).pack(side="right", padx=4)

    def compare(self):
        if not self.mapping_path().exists():
            return messagebox.showwarning("No mapping", "Match the columns first: press Suggest automatically "
                                          "or Edit by hand... under 2. Match columns.")
        try:
            args = ["compare"] + self.common_args() + ["--mapping", str(self.mapping_path()),
                                                       "--out-dir", str(self.results_dir())]
        except NotLoaded:
            return self.load(then=self.compare)
        except ValueError as e:
            return messagebox.showwarning("Missing input", str(e))
        if self.source.get() == "bq" and self.use_cache.get():
            args += ["--cache", str(self.work() / "bq_cache.pkl")]
        args += self.types_args()
        mapped = {oc for oc, nc in self.read_mapping_file().items() if nc}
        for role, var in self.roles.items():
            if var.get() != GUESS and var.get() in mapped:
                args += [f"--{role}-col", var.get()]
        self.run([COMPARE] + args, name="Comparing rows")

    def types_args(self):
        text = self.new_types.get().strip()
        return ["--new-types", text] if text else []

    def types_dialog(self):
        """Pick a data type for each mapped BigQuery column; fills the Column types field."""
        mapped = [nc for nc in self.read_mapping_file().values() if nc]
        if not mapped:
            return messagebox.showwarning("No mapping", "Match the columns first (2. Match columns); the types are "
                                          "set for the BigQuery columns in the mapping.")
        current = {}
        for part in self.new_types.get().split(";"):
            if "=" in part:
                col, typ = (x.strip() for x in part.rsplit("=", 1))
                current[col] = typ.upper()
        win = tk.Toplevel(self)
        win.title("Column types")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="How each BigQuery column should be read. The Excel column mapped to it is read the same "
                  "way.\nAUTO = guess from the values (the default). STRING keeps text exactly (e.g. leading zeros "
                  "in 0123). INTEGER drops decimals. DATE ignores the time of day. DATE_LOCAL first converts a UTC "
                  "timestamp to Norwegian time (2026-09-29 22:00 UTC -> 2026-09-30), then ignores the time; use it "
                  "for BigQuery TIMESTAMP columns. TIMESTAMP keeps the time. "
                  "BOOL treats 1/ja/true and 0/nei/false as the same.", wraplength=560, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w", padx=10, pady=8)
        options = ["AUTO", "STRING", "INTEGER", "NUMERIC", "DATE", "DATE_LOCAL", "TIMESTAMP", "BOOL"]
        vars_ = {}
        for r, col in enumerate(mapped, start=1):
            ttk.Label(win, text=col).grid(row=r, column=0, sticky="w", padx=10, pady=2)
            var = tk.StringVar(value=current.get(col, "AUTO") if current.get(col, "AUTO") in options else "AUTO")
            ttk.Combobox(win, textvariable=var, values=options, state="readonly", width=14).grid(
                row=r, column=1, sticky="w", padx=10, pady=2)
            vars_[col] = var

        def save():
            self.new_types.set("; ".join(f"{c}={v.get()}" for c, v in vars_.items() if v.get() != "AUTO"))
            win.destroy()
        bar = ttk.Frame(win)
        bar.grid(row=len(mapped) + 1, column=0, columnspan=2, sticky="e", padx=10, pady=10)
        ttk.Button(bar, text="Cancel", command=win.destroy).pack(side="right")
        ttk.Button(bar, text="OK", command=save).pack(side="right", padx=4)

    def roles_dialog(self):
        """Which Excel column is the date, group number, category and amount (used by the extra checks)."""
        mapped = [oc for oc, nc in self.read_mapping_file().items() if nc]
        if not mapped:
            return messagebox.showwarning("No mapping", "Match the columns first (2. Match columns).")
        win = tk.Toplevel(self)
        win.title("Column roles")
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="Optional. Used for the extra checks: dates outside the period, whole groups of rows missing "
                  "(e.g. a voucher or an order), reversals (+x/-x) and totals per month and category.\n"
                  "Leave a role on '(guess)' if your table has no such column. '(guess)' lets the program pick from "
                  "the names and values; the choice it made is shown in the messages when you compare.",
                  wraplength=520, justify="left").grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=8)
        texts = {"date": "Date", "voucher": "Group number (e.g. voucher, order)",
                 "account": "Category (e.g. account)", "amount": "Amount"}
        for r, (role, var) in enumerate(self.roles.items(), start=1):
            if var.get() not in mapped:
                var.set(GUESS)
            ttk.Label(win, text=texts[role]).grid(row=r, column=0, sticky="w", padx=10, pady=3)
            ttk.Combobox(win, textvariable=var, values=[GUESS] + mapped,
                         state="readonly", width=30).grid(row=r, column=1, sticky="w", padx=10, pady=3)
        ttk.Button(win, text="OK", command=win.destroy).grid(row=6, column=1, sticky="e", padx=10, pady=10)

    def download(self):
        table = self.bq_table.get().strip()
        if not table:
            return messagebox.showwarning("Missing input", "Enter the BigQuery table first.")
        out = filedialog.asksaveasfilename(defaultextension=".csv", initialdir=str(self.work()),
                                           initialfile=table.split(".")[-1] + ".csv", filetypes=[("CSV", "*.csv")])
        if not out:
            return

        def done(ok):
            if ok:
                self.new_file.set(out)
                self.source.set("file")
                self.write("\nThe CSV is selected as the new table.\n")
        cmd = [DOWNLOAD, "--table", table, "--out", out]
        if self.bq_where.get().strip():
            cmd += ["--where", self.clean_where()]
        self.run(cmd, name="Downloading from BigQuery", on_done=done)

    def clean_where(self):
        w = self.bq_where.get().strip()
        if w.lower().startswith("where "):
            w = w[6:].strip()
        return w.rstrip(";")

    def where_help(self):
        messagebox.showinfo("SQL filter examples", (
            "Only the condition, without the word WHERE. Use the BigQuery column names.\n\n"
            "One year:\n  dato BETWEEN '2025-01-01' AND '2025-12-31'\n\n"
            "From a date:\n  dato >= '2025-01-01'\n\n"
            "Some codes:\n  kode IN (3000, 3100)\n\n"
            "Combined:\n  dato >= '2025-01-01' AND selskap = 10\n\n"
            "Text values go in single quotes. Use the same range the old Excel file covers, "
            "so rows outside it don't slow the comparison down."))

    # ------------------------------------------------------------ columns ---
    # ------------------------------------------------------------ running ---
    def set_busy(self, name):
        self.task_name, self.task_start = name, time.time()
        for b in self.action_buttons:
            b.state(["disabled"])
        self.stop_button.state(["!disabled"])
        self.status_dot.config(fg="orange")
        self.progress.start(12)
        self.config(cursor="watch")
        self.tick()

    def set_idle(self, message, ok=True):
        self.task_name = None
        for b in self.action_buttons:
            b.state(["!disabled"])
        self.stop_button.state(["disabled"])
        self.progress.stop()
        self.status_dot.config(fg="green" if ok else "red")
        self.status.config(text=message)
        self.config(cursor="")

    def tick(self):
        if self.task_name:
            secs = int(time.time() - self.task_start)
            self.status.config(text=f"Working: {self.task_name} ...  {secs // 60}:{secs % 60:02d}")
            self.after(500, self.tick)

    def run(self, cmd, name="Working", on_done=None):
        if self.proc and self.proc.poll() is None:
            return messagebox.showinfo("Busy", "Something is already running. Wait or press Stop.")
        self.save_settings()
        script = cmd[0]
        if not Path(script).exists():
            return messagebox.showerror("Script missing", f"{Path(script).name} must be in the same folder as this window:\n{HERE}")
        full = [sys.executable, "-u", str(script)] + [str(c) for c in cmd[1:]]
        self.last_lines = []
        self.write("\n> " + " ".join(f'"{c}"' if " " in c else c for c in full[2:]) + "\n")
        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONHASHSEED="0")  # same input, same output
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        self.set_busy(name)
        try:
            self.proc = subprocess.Popen(full, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, cwd=str(self.work()),
                                         text=True, encoding="utf-8", errors="replace", env=env, creationflags=flags)
        except OSError as e:
            self.set_idle(f"Could not start: {e}", ok=False)
            return

        def reader(proc):
            for line in proc.stdout:
                self.output.put(line)
            code = proc.wait()
            self.output.put(("done", code, on_done))
        threading.Thread(target=reader, args=(self.proc,), daemon=True).start()

    def drain_output(self):
        try:
            while True:
                item = self.output.get_nowait()
                if isinstance(item, tuple):
                    _, code, cb = item
                    secs = int(time.time() - self.task_start)
                    took = f"{secs // 60}:{secs % 60:02d}"
                    name = self.task_name or "Task"
                    self.set_idle(f"Done: {name} ({took})" if code == 0 else f"Failed: {name} ({took}), see messages", code == 0)
                    self.write("Finished.\n" if code == 0 else f"Stopped with an error (code {code}). See the messages above.\n")
                    if cb:
                        cb(code == 0)
                else:
                    self.last_lines.append(item.rstrip("\n"))
                    if not item.startswith(("COLUMNS_JSON:", "ORIGINAL_JSON:", "SAMPLES_JSON:", "CHECK_JSON:", "NOTIN_FILE:")):
                        self.write(item)
        except queue.Empty:
            pass
        self.after(100, self.drain_output)

    def stop(self):
        if self.proc and self.proc.poll() is None:
            self.proc.terminate()

    def write(self, text):
        self.log.insert("end", text)
        self.log.see("end")

    def open_path(self, path):
        path = Path(path)
        if not path.exists():
            return messagebox.showinfo("Not there yet", f"{path} doesn't exist yet.")
        if sys.platform.startswith("win"):
            os.startfile(str(path))
        elif sys.platform == "darwin":
            subprocess.run(["open", str(path)])
        else:
            subprocess.run(["xdg-open", str(path)])


if __name__ == "__main__":
    if sys.platform.startswith("win"):
        try:  # sharp text on high-DPI screens
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()
