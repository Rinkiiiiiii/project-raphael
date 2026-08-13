"""
main_script_ui.py
------------------
UI layer for Raphael's "excel_graph_creation" sub-program.

Pure presentation: this module builds the window, collects input,
and hands the assembled GraphRequest to main_script.py for all
validation, workbook/chart building, saving, and file-opening.
No spreadsheet or file-system logic lives here.
"""

from __future__ import annotations

import tkinter as tk
from datetime import date, datetime
from tkinter import messagebox, ttk
from typing import List, Optional

from main_script import (
    DataPoint,
    GraphRequest,
    GraphRequestError,
    GraphType,
    generate_excel_graph,
    open_file,
)

DATE_FORMAT = "%m/%d/%Y"


class RaphaelGraphApp(tk.Tk):
    """Main application window for creating a single graph's Excel file."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Raphael — Excel Graph Creation")
        self.geometry("640x620")
        self.minsize(560, 540)

        self._data_points: List[DataPoint] = []
        self._graph_type_var = tk.StringVar(value=GraphType.BEHAVIOR.value)

        self._build_layout()
        self._on_graph_type_changed()

    # ------------------------------------------------------------------ #
    # Layout
    # ------------------------------------------------------------------ #

    def _build_layout(self) -> None:
        pad = {"padx": 10, "pady": 6}

        # --- Client -------------------------------------------------------
        client_frame = ttk.LabelFrame(self, text="1. Client")
        client_frame.pack(fill="x", **pad)

        ttk.Label(client_frame, text="First Name:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self._client_first_entry = ttk.Entry(client_frame, width=20)
        self._client_first_entry.grid(row=0, column=1, sticky="w", padx=8, pady=6)

        ttk.Label(client_frame, text="Last Name:").grid(row=0, column=2, sticky="w", padx=8, pady=6)
        self._client_last_entry = ttk.Entry(client_frame, width=20)
        self._client_last_entry.grid(row=0, column=3, sticky="w", padx=8, pady=6)

        # --- Graph type selection -----------------------------------------
        type_frame = ttk.LabelFrame(self, text="2. Choose Graph Type")
        type_frame.pack(fill="x", **pad)

        for graph_type in GraphType:
            ttk.Radiobutton(
                type_frame,
                text=graph_type.value,
                value=graph_type.value,
                variable=self._graph_type_var,
                command=self._on_graph_type_changed,
            ).pack(side="left", padx=10, pady=6)

        # --- Details (category name) -----------------------------------------
        details_frame = ttk.LabelFrame(self, text="3. Details")
        details_frame.pack(fill="x", **pad)

        self._category_label_widget = ttk.Label(details_frame, text="")
        self._category_label_widget.grid(row=0, column=0, sticky="w", padx=8, pady=6)
        self._category_entry = ttk.Entry(details_frame, width=35)
        self._category_entry.grid(row=0, column=1, sticky="w", padx=8, pady=6)

        # --- Data entry ------------------------------------------------------
        entry_frame = ttk.LabelFrame(self, text="4. Add Data Points")
        entry_frame.pack(fill="x", **pad)

        ttk.Label(entry_frame, text="Date (MM/DD/YYYY):").grid(
            row=0, column=0, sticky="w", padx=8, pady=6
        )
        self._date_entry = ttk.Entry(entry_frame, width=15)
        self._date_entry.grid(row=0, column=1, sticky="w", padx=8, pady=6)
        self._date_entry.insert(0, date.today().strftime(DATE_FORMAT))

        self._value_label_widget = ttk.Label(entry_frame, text="")
        self._value_label_widget.grid(row=0, column=2, sticky="w", padx=8, pady=6)
        self._value_entry = ttk.Entry(entry_frame, width=10)
        self._value_entry.grid(row=0, column=3, sticky="w", padx=8, pady=6)

        ttk.Button(entry_frame, text="Add Row", command=self._on_add_row).grid(
            row=0, column=4, padx=8, pady=6
        )

        # --- Data table --------------------------------------------------------
        table_frame = ttk.Frame(self)
        table_frame.pack(fill="both", expand=True, **pad)

        self._tree = ttk.Treeview(
            table_frame, columns=("date", "value"), show="headings", height=10
        )
        self._tree.heading("date", text="Date")
        self._tree.heading("value", text="Value")
        self._tree.column("date", width=150, anchor="center")
        self._tree.column("value", width=150, anchor="center")
        self._tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=self._tree.yview)
        scrollbar.pack(side="left", fill="y")
        self._tree.configure(yscrollcommand=scrollbar.set)

        ttk.Button(self, text="Remove Selected Row", command=self._on_remove_row).pack(
            anchor="e", padx=10
        )

        # --- Generate ----------------------------------------------------------
        ttk.Button(
            self, text="Generate Graph", command=self._on_generate
        ).pack(pady=14, ipadx=10, ipady=6)

        self._status_var = tk.StringVar(value="")
        ttk.Label(self, textvariable=self._status_var, foreground="#555").pack(pady=(0, 10))

    # ------------------------------------------------------------------ #
    # Dynamic label updates based on selected graph type
    # ------------------------------------------------------------------ #

    def _current_graph_type(self) -> GraphType:
        return GraphType(self._graph_type_var.get())

    def _on_graph_type_changed(self) -> None:
        graph_type = self._current_graph_type()
        self._category_label_widget.config(text=f"{graph_type.category_label}:")
        unit = "%" if graph_type.is_percentage else "count"
        self._value_label_widget.config(text=f"{graph_type.value_label} ({unit}):")

        # Data points are specific to one category/graph type; changing type
        # invalidates whatever was already entered, so clear the table.
        self._data_points.clear()
        self._refresh_table()

    # ------------------------------------------------------------------ #
    # Data entry handlers
    # ------------------------------------------------------------------ #

    def _on_add_row(self) -> None:
        raw_date = self._date_entry.get().strip()
        raw_value = self._value_entry.get().strip()

        try:
            parsed_date = datetime.strptime(raw_date, DATE_FORMAT).date()
        except ValueError:
            messagebox.showerror("Invalid Date", f"Please enter a date as {DATE_FORMAT}.")
            return

        try:
            parsed_value = float(raw_value)
        except ValueError:
            messagebox.showerror("Invalid Value", "Please enter a numeric value.")
            return

        graph_type = self._current_graph_type()
        if graph_type.is_percentage and not (0 <= parsed_value <= 100):
            messagebox.showerror("Invalid Value", "Percentage values must be between 0 and 100.")
            return
        if not graph_type.is_percentage and parsed_value < 0:
            messagebox.showerror("Invalid Value", "Frequency cannot be negative.")
            return

        self._data_points.append(DataPoint(entry_date=parsed_date, value=parsed_value))
        self._data_points.sort(key=lambda p: p.entry_date)
        self._refresh_table()

        # Convenience: clear only the value field, keep the date field for
        # quick consecutive entry of the next day.
        self._value_entry.delete(0, tk.END)

    def _on_remove_row(self) -> None:
        selected = self._tree.selection()
        if not selected:
            return
        index = self._tree.index(selected[0])
        del self._data_points[index]
        self._refresh_table()

    def _refresh_table(self) -> None:
        self._tree.delete(*self._tree.get_children())
        graph_type = self._current_graph_type()
        for point in self._data_points:
            display_value = f"{point.value:.1f}%" if graph_type.is_percentage else f"{point.value:g}"
            self._tree.insert("", "end", values=(point.entry_date.strftime(DATE_FORMAT), display_value))

    # ------------------------------------------------------------------ #
    # Generate
    # ------------------------------------------------------------------ #

    def _on_generate(self) -> None:
        request = GraphRequest(
            graph_type=self._current_graph_type(),
            category_name=self._category_entry.get(),
            client_first_name=self._client_first_entry.get(),
            client_last_name=self._client_last_entry.get(),
            data_points=list(self._data_points),
        )

        try:
            result = generate_excel_graph(request)
        except GraphRequestError as exc:
            messagebox.showerror("Cannot Generate Graph", str(exc))
            return
        except OSError as exc:
            messagebox.showerror("File Error", f"Could not save the file:\n{exc}")
            return

        action = "Replaced existing" if result.was_replacement else "Added new"
        self._status_var.set(
            f"{action} table on '{result.sheet_name}' tab (row {result.anchor_row}) — "
            f"{result.workbook_path}"
        )
        open_file(result.workbook_path)


def main() -> None:
    app = RaphaelGraphApp()
    app.mainloop()


if __name__ == "__main__":
    main()