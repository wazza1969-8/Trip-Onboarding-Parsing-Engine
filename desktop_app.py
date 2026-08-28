"""
Trip Support Intake -- local desktop app (Tkinter)
=====================================================

A real desktop application window (no browser, no server, no hosting).
Each install is fully self-contained: submissions and results are saved
in a local SQLite file in this user's own app-data folder (see
desktop_backend.local_app_data_dir()) -- nothing is shared with anyone
else's install.

Requirements
------------
    pip install anthropic pymupdf openpyxl
    (tkinter ships with standard Python installs on Windows/Mac; on some
    Linux distros it's a separate OS package, e.g. `apt install python3-tk`)

Before building
----------------
    Fill in ANTHROPIC_API_KEY in config.py.

Run from source
----------------
    python desktop_app.py

Building a double-click installer
-----------------------------------
    See BUILD_INSTRUCTIONS.md for the one-time PyInstaller commands to
    produce a Windows .exe and a Mac .app. PyInstaller cannot cross
    build, so each has to be built on that actual OS.
"""

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import desktop_backend as backend

# The salesperson list itself lives in this user's local database
# (desktop_backend.py's salespeople table), editable via the "Manage..."
# button next to the dropdown -- see _open_manage_salespeople_dialog().
# desktop_backend.DEFAULT_SALESPEOPLE is only the one-time seed.
PAYMENT_METHODS = ["Credit Card", "Ok to Invoice"]
MAX_ACCT_NBR_LEN = 20
ALLOWED_FILETYPES = [
    ("Onboarding form", "*.pdf *.doc *.docx *.png *.jpg *.jpeg"),
    ("All files", "*.*"),
]


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Trip Support Intake")
        self.geometry("600x680")
        self.minsize(560, 600)

        self.selected_file_path: str | None = None
        self.progress_queue: queue.Queue = queue.Queue()

        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)

        self.new_request_frame = ttk.Frame(notebook)
        self.requests_frame = ttk.Frame(notebook)
        notebook.add(self.new_request_frame, text="Submit New Request")
        notebook.add(self.requests_frame, text="My Requests")

        self._build_new_request_tab()
        self._build_requests_tab()
        self._refresh_requests()

    # -- New Request tab ----------------------------------------------
    def _build_new_request_tab(self):
        frame = self.new_request_frame
        pad = {"padx": 8, "pady": 6}

        ttk.Label(frame, text="Salesperson").grid(row=0, column=0, sticky="w", **pad)
        self.salesperson_var = tk.StringVar()
        self.salesperson_combo = ttk.Combobox(
            frame, textvariable=self.salesperson_var,
            state="readonly", width=30,
        )
        self.salesperson_combo.grid(row=0, column=1, sticky="w", **pad)
        ttk.Button(
            frame, text="Manage...", command=self._open_manage_salespeople_dialog,
        ).grid(row=0, column=2, sticky="w", **pad)
        self._refresh_salespeople_combo()

        ttk.Label(frame, text="Jeppesen Acct Nbr:").grid(row=1, column=0, sticky="w", **pad)
        self.acct_var = tk.StringVar()

        def _limit_len(*_args):
            val = self.acct_var.get()
            if len(val) > MAX_ACCT_NBR_LEN:
                self.acct_var.set(val[:MAX_ACCT_NBR_LEN])

        self.acct_var.trace_add("write", _limit_len)
        ttk.Entry(frame, textvariable=self.acct_var, width=32).grid(row=1, column=1, sticky="w", **pad)

        ttk.Label(frame, text="Payment Method").grid(row=2, column=0, sticky="w", **pad)
        self.payment_var = tk.StringVar(value=PAYMENT_METHODS[0])
        radio_frame = ttk.Frame(frame)
        radio_frame.grid(row=2, column=1, sticky="w", **pad)
        for opt in PAYMENT_METHODS:
            ttk.Radiobutton(
                radio_frame, text=opt, value=opt, variable=self.payment_var,
                command=self._update_conditional,
            ).pack(side="left", padx=4)

        self.conditional_frame = ttk.Frame(frame)
        self.conditional_frame.grid(row=3, column=0, columnspan=2, sticky="we", **pad)

        ttk.Label(self.conditional_frame, text="Amount (required)").grid(row=0, column=0, sticky="w")
        self.amount_var = tk.StringVar()
        ttk.Entry(self.conditional_frame, textvariable=self.amount_var, width=15).grid(
            row=0, column=1, sticky="w", padx=6
        )

        ttk.Label(self.conditional_frame, text="Notes (optional)").grid(row=1, column=0, sticky="nw", pady=(6, 0))
        self.notes_text = tk.Text(self.conditional_frame, width=42, height=3)
        self.notes_text.grid(row=1, column=1, sticky="w", pady=(6, 0))

        ttk.Label(frame, text="Onboarding Form Upload (optional)").grid(row=4, column=0, sticky="w", **pad)
        file_frame = ttk.Frame(frame)
        file_frame.grid(row=5, column=0, columnspan=2, sticky="we", **pad)
        ttk.Button(file_frame, text="Browse...", command=self._browse_file).pack(side="left")
        self.file_label_var = tk.StringVar(value="No file selected")
        ttk.Label(file_frame, textvariable=self.file_label_var).pack(side="left", padx=8)
        ttk.Button(file_frame, text="Clear", command=self._clear_file).pack(side="left")

        self.error_label = ttk.Label(frame, foreground="red", wraplength=480, justify="left")
        self.error_label.grid(row=6, column=0, columnspan=2, sticky="w", **pad)

        self.progress = ttk.Progressbar(frame, length=480, mode="determinate", maximum=100)
        self.progress.grid(row=7, column=0, columnspan=2, sticky="we", **pad)
        self.progress_label_var = tk.StringVar(value="")
        ttk.Label(frame, textvariable=self.progress_label_var).grid(row=8, column=0, columnspan=2, sticky="w", **pad)

        ttk.Button(frame, text="Submit", command=self._on_submit).grid(row=9, column=0, columnspan=2, pady=12)

        self._update_conditional()

    def _update_conditional(self):
        if self.payment_var.get() == "Ok to Invoice":
            self.conditional_frame.grid()
        else:
            self.conditional_frame.grid_remove()

    def _refresh_salespeople_combo(self):
        """Reloads the dropdown from the database. Keeps the current
        selection if it still exists; otherwise falls back to the first
        entry (or blank if the list is now empty)."""
        names = backend.list_salespeople()
        current = self.salesperson_var.get()
        self.salesperson_combo["values"] = names
        if current in names:
            self.salesperson_var.set(current)
        elif names:
            self.salesperson_var.set(names[0])
        else:
            self.salesperson_var.set("")

    def _open_manage_salespeople_dialog(self):
        dialog = tk.Toplevel(self)
        dialog.title("Manage salesperson list")
        dialog.geometry("360x360")
        dialog.transient(self)
        dialog.grab_set()

        ttk.Label(dialog, text="Salespeople").pack(anchor="w", padx=10, pady=(10, 0))
        list_frame = ttk.Frame(dialog)
        list_frame.pack(fill="both", expand=True, padx=10, pady=6)
        listbox = tk.Listbox(list_frame, exportselection=False)
        listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=listbox.yview)
        scrollbar.pack(side="right", fill="y")
        listbox.config(yscrollcommand=scrollbar.set)

        def _reload_listbox(select_name: str | None = None):
            listbox.delete(0, "end")
            names = backend.list_salespeople()
            for name in names:
                listbox.insert("end", name)
            if select_name in names:
                listbox.selection_set(names.index(select_name))

        _reload_listbox()

        entry_frame = ttk.Frame(dialog)
        entry_frame.pack(fill="x", padx=10, pady=(0, 6))
        ttk.Label(entry_frame, text="Name:").pack(side="left")
        name_var = tk.StringVar()
        ttk.Entry(entry_frame, textvariable=name_var).pack(side="left", fill="x", expand=True, padx=6)

        dialog_error = ttk.Label(dialog, foreground="red", wraplength=330, justify="left")
        dialog_error.pack(fill="x", padx=10)

        def _selected_name() -> str | None:
            sel = listbox.curselection()
            if not sel:
                return None
            return listbox.get(sel[0])

        def _on_select(_event=None):
            name = _selected_name()
            if name is not None:
                name_var.set(name)

        listbox.bind("<<ListboxSelect>>", _on_select)

        def _do_add():
            dialog_error.config(text="")
            try:
                backend.add_salesperson(name_var.get())
                added_name = name_var.get().strip()
                name_var.set("")
                _reload_listbox(select_name=added_name)
                self._refresh_salespeople_combo()
            except ValueError as e:
                dialog_error.config(text=str(e))

        def _do_rename():
            dialog_error.config(text="")
            old_name = _selected_name()
            if old_name is None:
                dialog_error.config(text="Select a salesperson to rename first.")
                return
            try:
                backend.update_salesperson(old_name, name_var.get())
                new_name = name_var.get().strip()
                _reload_listbox(select_name=new_name)
                self._refresh_salespeople_combo()
            except ValueError as e:
                dialog_error.config(text=str(e))

        def _do_delete():
            dialog_error.config(text="")
            name = _selected_name()
            if name is None:
                dialog_error.config(text="Select a salesperson to delete first.")
                return
            if not messagebox.askyesno("Delete salesperson", f"Delete '{name}'?", parent=dialog):
                return
            backend.delete_salesperson(name)
            name_var.set("")
            _reload_listbox()
            self._refresh_salespeople_combo()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(button_frame, text="Add", command=_do_add).pack(side="left")
        ttk.Button(button_frame, text="Rename selected", command=_do_rename).pack(side="left", padx=6)
        ttk.Button(button_frame, text="Delete selected", command=_do_delete).pack(side="left")
        ttk.Button(dialog, text="Close", command=dialog.destroy).pack(pady=(0, 10))

    def _browse_file(self):
        path = filedialog.askopenfilename(title="Select onboarding form", filetypes=ALLOWED_FILETYPES)
        if path:
            self.selected_file_path = path
            self.file_label_var.set(Path(path).name)

    def _clear_file(self):
        self.selected_file_path = None
        self.file_label_var.set("No file selected")

    def _on_submit(self):
        self.error_label.config(text="")
        if not self.salesperson_var.get():
            self.error_label.config(text="Add a salesperson via 'Manage...' before submitting.")
            return
        payment_method = self.payment_var.get()
        amount = None
        notes = ""

        if payment_method == "Ok to Invoice":
            raw = self.amount_var.get().strip()
            try:
                amount = float(raw)
            except ValueError:
                amount = None
            if amount is None or amount <= 0:
                self.error_label.config(
                    text="Amount is required and must be greater than 0 when 'Ok to Invoice' is selected."
                )
                return
            notes = self.notes_text.get("1.0", "end").strip()

        onboarding_bytes = None
        onboarding_filename = None
        if self.selected_file_path:
            onboarding_filename = Path(self.selected_file_path).name
            onboarding_bytes = Path(self.selected_file_path).read_bytes()

        request_id = backend.insert_request({
            "salesperson": self.salesperson_var.get(),
            "jeppesen_acct_nbr": self.acct_var.get(),
            "payment_method": payment_method,
            "amount": amount if payment_method == "Ok to Invoice" else None,
            "notes": notes,
            "onboarding_filename": onboarding_filename,
            "onboarding_bytes": onboarding_bytes,
        })

        if onboarding_bytes is not None:
            self._run_parsing_async(request_id, onboarding_bytes, salesperson=self.salesperson_var.get())
        else:
            messagebox.showinfo(
                "Submitted",
                f"Submitted. Request ID: {request_id}\nNo onboarding form attached -- nothing to parse.",
            )
            self._reset_form()
            self._refresh_requests()

    def _run_parsing_async(self, request_id: str, pdf_bytes: bytes, salesperson: str | None = None):
        self.progress["value"] = 0
        self.progress_label_var.set("Starting...")

        def progress_cb(fraction, text):
            # Called from the worker thread -- never touch Tk widgets
            # directly here. Just hand data to the main thread via the
            # queue; _poll_progress (running on the main thread/event
            # loop) does the actual widget updates.
            self.progress_queue.put(("progress", fraction, text))

        def worker():
            backend.run_parsing(request_id, pdf_bytes, progress_callback=progress_cb, salesperson=salesperson)
            self.progress_queue.put(("done", request_id))

        threading.Thread(target=worker, daemon=True).start()
        self.after(100, self._poll_progress)

    def _poll_progress(self):
        done = False
        try:
            while True:
                item = self.progress_queue.get_nowait()
                if item[0] == "done":
                    done = True
                    request_id = item[1]
                    row = backend.fetch_request(request_id)
                    self.progress_label_var.set("")
                    self.progress["value"] = 0
                    if row["status"] == backend.STATUS_COMPLETE:
                        messagebox.showinfo(
                            "Complete",
                            f"Request {request_id} complete. See 'My Requests' to save the result.",
                        )
                    else:
                        messagebox.showerror("Failed", f"Request {request_id} failed:\n{row['error_message']}")
                    self._reset_form()
                    self._refresh_requests()
                    break
                else:
                    _, fraction, text = item
                    self.progress["value"] = fraction * 100
                    self.progress_label_var.set(text)
        except queue.Empty:
            pass
        if not done:
            self.after(100, self._poll_progress)

    def _reset_form(self):
        self.acct_var.set("")
        self.amount_var.set("")
        self.notes_text.delete("1.0", "end")
        self.selected_file_path = None
        self.file_label_var.set("No file selected")
        self.payment_var.set(PAYMENT_METHODS[0])
        self._update_conditional()

    # -- My Requests tab ------------------------------------------------
    def _build_requests_tab(self):
        frame = self.requests_frame
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill="x", pady=6)
        ttk.Button(toolbar, text="Refresh", command=self._refresh_requests).pack(side="left", padx=6)
        ttk.Button(toolbar, text="Save Result As...", command=self._save_selected_result).pack(side="left", padx=6)
        ttk.Button(toolbar, text="Retry", command=self._retry_selected).pack(side="left", padx=6)
        ttk.Button(toolbar, text="View Error", command=self._view_selected_error).pack(side="left", padx=6)

        columns = ("id", "salesperson", "acct_nbr", "status", "created_at")
        self.tree = ttk.Treeview(frame, columns=columns, show="headings", height=16)
        headings = {
            "id": "Request ID", "salesperson": "Salesperson", "acct_nbr": "Acct Nbr",
            "status": "Status", "created_at": "Created At",
        }
        for col in columns:
            self.tree.heading(col, text=headings[col])
            self.tree.column(col, width=160 if col == "created_at" else 110)
        self.tree.pack(fill="both", expand=True, padx=8, pady=6)

    def _refresh_requests(self):
        for row_id in self.tree.get_children():
            self.tree.delete(row_id)
        for r in backend.fetch_all_requests():
            self.tree.insert(
                "", "end", iid=r["id"],
                values=(r["id"], r["salesperson"], r["jeppesen_acct_nbr"] or "", r["status"], r["created_at"]),
            )

    def _selected_request(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("No selection", "Select a request first.")
            return None
        return backend.fetch_request(sel[0])

    def _save_selected_result(self):
        row = self._selected_request()
        if row is None:
            return
        if row["status"] != backend.STATUS_COMPLETE or not row["result_bytes"]:
            messagebox.showinfo("Not ready", "This request doesn't have a completed result yet.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=f"{row['id']}_operational_info.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if path:
            Path(path).write_bytes(row["result_bytes"])
            messagebox.showinfo("Saved", f"Saved to {path}")

    def _retry_selected(self):
        row = self._selected_request()
        if row is None:
            return
        if row["status"] != backend.STATUS_FAILED or not row["onboarding_bytes"]:
            messagebox.showinfo("Can't retry", "Only failed requests with an attached form can be retried.")
            return
        backend.update_status(row["id"], backend.STATUS_PROCESSING)
        self._refresh_requests()
        self._run_parsing_async(row["id"], row["onboarding_bytes"], salesperson=row["salesperson"])

    def _view_selected_error(self):
        row = self._selected_request()
        if row is None:
            return
        if row["status"] != backend.STATUS_FAILED:
            messagebox.showinfo("No error", "This request has no error to show.")
            return
        messagebox.showerror("Error", row["error_message"] or "Unknown error.")


if __name__ == "__main__":
    app = App()
    app.mainloop()
