"""DFIP Case Sorting System backed by a Microsoft Access database.

The application reads cases from an Access .accdb or .mdb file. It expects a
selected table to contain these columns:
    CaseID, Case_Title, Case_Type, Date_Reported, Severity, Status

The database is read-only from this application. Sorting and removing rows
only affect the current display; they do not change the database.
"""

from __future__ import annotations

import os
import tkinter as tk
from datetime import date, datetime
from tkinter import filedialog, messagebox, ttk
from typing import Any, List, Optional, Sequence, Tuple


# Severity ranks allow sorting by seriousness instead of alphabetically.
# Descending order therefore displays Critical, High, Medium, then Low.
SEVERITY_ORDER = {
    "informational": 0,
    "info": 0,
    "low": 1,
    "minor": 1,
    "medium": 2,
    "moderate": 2,
    "high": 3,
    "major": 3,
    "critical": 4,
    "severe": 4,
}

# The field name in the database is mapped to the attribute used by the class.
DATABASE_FIELDS = (
    "CaseID",
    "Case_Title",
    "Case_Type",
    "Date_Reported",
    "Severity",
    "Status",
)

SORT_OPTIONS = {
    "Case ID": "case_id",
    "Case Title": "case_title",
    "Case Type": "case_type",
    "Date Reported": "date_reported",
    "Severity": "severity",
    "Status": "status",
}


class DFIPCase:
    """Store the six fields belonging to one DFIP case."""

    def __init__(
        self,
        case_id: Any,
        case_title: Any,
        case_type: Any,
        date_reported: Any,
        severity: Any,
        status: Any,
    ) -> None:
        self.case_id = case_id
        self.case_title = _display_text(case_title)
        self.case_type = _display_text(case_type)
        self.date_reported = _convert_date(date_reported)
        self.severity = _display_text(severity)
        self.status = _display_text(status)

    def get_sort_value(self, sort_by: str) -> Any:
        """Return a normalised value for the selected sorting field."""
        if sort_by == "case_id":
            if self.case_id is None or str(self.case_id).strip() == "":
                return None
            return _normalise_case_id(self.case_id)
        if sort_by == "case_title":
            return self.case_title.casefold()
        if sort_by == "case_type":
            return self.case_type.casefold()
        if sort_by == "date_reported":
            return self.date_reported
        if sort_by == "severity":
            severity_key = self.severity.casefold()
            if not severity_key:
                return None
            if severity_key in SEVERITY_ORDER:
                return (0, SEVERITY_ORDER[severity_key])
            return (1, severity_key)
        if sort_by == "status":
            return self.status.casefold()
        raise ValueError(
            "Choose Case ID, Case Title, Case Type, Date Reported, "
            "Severity or Status to sort by."
        )

    def display_details(self) -> Tuple[str, str, str, str, str, str]:
        """Return the case details in the order used by the results table."""
        return (
            _display_text(self.case_id),
            self.case_title,
            self.case_type,
            self.date_reported.strftime("%d/%m/%Y")
            if self.date_reported else "",
            self.severity,
            self.status,
        )


def _display_text(value: Any) -> str:
    """Convert database values to clean display text without showing None."""
    if value is None:
        return ""
    return str(value).strip()


def _convert_date(value: Any) -> Optional[date]:
    """Convert common Access date values into Python dates."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    date_text = str(value).strip()
    for date_format in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%d-%m-%Y",
    ):
        try:
            return datetime.strptime(date_text, date_format).date()
        except ValueError:
            continue

    # A value that cannot be parsed is treated as missing rather than
    # preventing the rest of the database from loading.
    return None


def _normalise_case_id(value: Any) -> Tuple[int, Any]:
    """Sort numeric IDs numerically and other IDs alphabetically."""
    if value is None or str(value).strip() == "":
        return (2, "")
    try:
        return (0, float(value))
    except (TypeError, ValueError):
        return (1, str(value).casefold())


def _sort_comparison_value(case: DFIPCase, sort_by: str) -> Tuple[bool, Any]:
    """Return (is_missing, value), keeping missing data last in either order."""
    value = case.get_sort_value(sort_by)
    missing = value is None or value == ""
    return missing, value


def sort_cases(
    cases: List[DFIPCase],
    sort_by: str = "case_id",
    descending: bool = False,
) -> List[DFIPCase]:
    """Sort cases in place with a stable bubble sort and return the list.

    Missing values are always kept at the bottom. Duplicate values keep their
    original relative order, including when descending order is selected.
    """
    valid_fields = set(SORT_OPTIONS.values())
    if sort_by not in valid_fields:
        raise ValueError(f"sort_by must be one of: {', '.join(sorted(valid_fields))}.")

    for pass_number in range(len(cases) - 1):
        swapped = False
        for position in range(len(cases) - pass_number - 1):
            left_missing, left_value = _sort_comparison_value(
                cases[position], sort_by
            )
            right_missing, right_value = _sort_comparison_value(
                cases[position + 1], sort_by
            )

            # Missing values remain last, regardless of the selected direction.
            if left_missing and not right_missing:
                should_swap = True
            elif not left_missing and right_missing:
                should_swap = False
            elif left_missing and right_missing:
                should_swap = False
            elif descending:
                should_swap = left_value < right_value
            else:
                should_swap = left_value > right_value

            if should_swap:
                cases[position], cases[position + 1] = (
                    cases[position + 1],
                    cases[position],
                )
                swapped = True

        if not swapped:
            break

    return cases


def _get_pyodbc():
    """Import pyodbc only when a database connection is requested."""
    try:
        import pyodbc  # type: ignore
    except ImportError as error:
        raise RuntimeError(
            "The pyodbc package is not installed. Open a terminal and run:\n\n"
            "python -m pip install pyodbc"
        ) from error
    return pyodbc


def _access_connection_string(database_path: str) -> str:
    """Build a connection string using an installed Microsoft Access driver."""
    pyodbc = _get_pyodbc()
    available_drivers = pyodbc.drivers()
    access_drivers = [
        driver for driver in available_drivers if "access" in driver.casefold()
    ]
    if not access_drivers:
        raise RuntimeError(
            "A Microsoft Access ODBC driver could not be found. Install the "
            "Microsoft Access Database Engine that matches your Python "
            "installation (32-bit or 64-bit), then try again."
        )

    driver = access_drivers[0]
    return f"DRIVER={{{driver}}};DBQ={database_path};"


def get_access_tables(database_path: str) -> List[str]:
    """Return user table names from an Access database."""
    pyodbc = _get_pyodbc()
    connection_string = _access_connection_string(database_path)
    connection = None
    try:
        connection = pyodbc.connect(connection_string)
        cursor = connection.cursor()
        tables = [
            row.table_name
            for row in cursor.tables(tableType="TABLE")
            if not str(row.table_name).startswith("MSys")
        ]
        cursor.close()
        return sorted(set(tables), key=str.casefold)
    except Exception as error:
        raise RuntimeError(f"Could not read the Access database: {error}") from error
    finally:
        if connection is not None:
            connection.close()


def load_cases_from_access(database_path: str, table_name: str) -> List[DFIPCase]:
    """Read the six required case fields from an Access table.

    The table name comes from the database's discovered table list and is
    quoted as an Access identifier. Case data is retrieved using parameters
    only where values are needed; no user-provided SQL is executed.
    """
    if not os.path.isfile(database_path):
        raise FileNotFoundError(f"Database file not found: {database_path}")

    pyodbc = _get_pyodbc()
    connection_string = _access_connection_string(database_path)
    safe_table_name = table_name.replace("]", "]]")
    selected_fields = ", ".join(f"[{field}]" for field in DATABASE_FIELDS)
    query = f"SELECT {selected_fields} FROM [{safe_table_name}]"

    connection = None
    try:
        connection = pyodbc.connect(connection_string)
        cursor = connection.cursor()
        cursor.execute(query)
        rows = cursor.fetchall()
        cases = [
            DFIPCase(
                case_id=row[0],
                case_title=row[1],
                case_type=row[2],
                date_reported=row[3],
                severity=row[4],
                status=row[5],
            )
            for row in rows
        ]
        cursor.close()
        return cases
    except Exception as error:
        raise RuntimeError(
            "The table could not be loaded. Check that it contains all six "
            "columns with these exact names: " + ", ".join(DATABASE_FIELDS)
            + f". Database message: {error}"
        ) from error
    finally:
        if connection is not None:
            connection.close()


class DFIPCaseApp:
    """Build the GUI and handle user and database events."""

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.cases: List[DFIPCase] = []
        self.database_path: Optional[str] = None
        self.table_name: Optional[str] = None

        self.root.title("DFIP Case Sorting System")
        self.root.geometry("1120x680")
        self.root.minsize(880, 560)

        self._build_interface()
        self.display_cases()

    def _build_interface(self) -> None:
        """Create the database controls, sorting controls and results table."""
        main_frame = ttk.Frame(self.root, padding=14)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(
            main_frame,
            text="DFIP Case Sorting System",
            font=("Arial", 18, "bold"),
        ).pack(pady=(0, 12))

        database_frame = ttk.LabelFrame(
            main_frame, text="Microsoft Access database", padding=10
        )
        database_frame.pack(fill="x", pady=(0, 10))

        ttk.Button(
            database_frame,
            text="Connect / Load Database",
            command=self.connect_to_database,
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            database_frame,
            text="Refresh Data",
            command=self.refresh_database,
        ).pack(side="left", padx=8)
        ttk.Button(
            database_frame,
            text="Disconnect",
            command=self.disconnect_database,
        ).pack(side="left", padx=8)

        self.database_label = ttk.Label(
            database_frame, text="No database loaded", wraplength=600
        )
        self.database_label.pack(side="left", padx=12, fill="x", expand=True)

        sort_frame = ttk.LabelFrame(main_frame, text="Sort cases", padding=10)
        sort_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(sort_frame, text="Sort by").grid(
            row=0, column=0, padx=5, pady=5
        )
        self.sort_choice = ttk.Combobox(
            sort_frame,
            values=tuple(SORT_OPTIONS.keys()),
            state="readonly",
            width=19,
        )
        self.sort_choice.set("Case ID")
        self.sort_choice.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(sort_frame, text="Order").grid(
            row=0, column=2, padx=(15, 5), pady=5
        )
        self.order_choice = ttk.Combobox(
            sort_frame,
            values=("Ascending", "Descending"),
            state="readonly",
            width=15,
        )
        self.order_choice.set("Ascending")
        self.order_choice.grid(row=0, column=3, padx=5, pady=5)

        ttk.Button(
            sort_frame, text="Sort Cases", command=self.sort_and_display
        ).grid(row=0, column=4, padx=(15, 5), pady=5)
        ttk.Button(
            sort_frame, text="Display All", command=self.display_cases
        ).grid(row=0, column=5, padx=5, pady=5)

        results_frame = ttk.LabelFrame(
            main_frame, text="DFIP cases from database", padding=8
        )
        results_frame.pack(fill="both", expand=True)

        columns = (
            "case_id",
            "case_title",
            "case_type",
            "date_reported",
            "severity",
            "status",
        )
        self.case_table = ttk.Treeview(
            results_frame,
            columns=columns,
            show="headings",
            selectmode="browse",
        )

        headings = {
            "case_id": "CaseID",
            "case_title": "Case_Title",
            "case_type": "Case_Type",
            "date_reported": "Date_Reported",
            "severity": "Severity",
            "status": "Status",
        }
        column_widths = {
            "case_id": 85,
            "case_title": 240,
            "case_type": 150,
            "date_reported": 130,
            "severity": 110,
            "status": 130,
        }

        for column in columns:
            self.case_table.heading(column, text=headings[column])
            self.case_table.column(
                column,
                width=column_widths[column],
                minwidth=70,
                anchor="center" if column in {"case_id", "date_reported"} else "w",
            )

        vertical_scrollbar = ttk.Scrollbar(
            results_frame, orient="vertical", command=self.case_table.yview
        )
        horizontal_scrollbar = ttk.Scrollbar(
            results_frame, orient="horizontal", command=self.case_table.xview
        )
        self.case_table.configure(
            yscrollcommand=vertical_scrollbar.set,
            xscrollcommand=horizontal_scrollbar.set,
        )

        self.case_table.grid(row=0, column=0, sticky="nsew")
        vertical_scrollbar.grid(row=0, column=1, sticky="ns")
        horizontal_scrollbar.grid(row=1, column=0, sticky="ew")
        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill="x", pady=(10, 0))
        ttk.Button(
            action_frame,
            text="Remove Selected From View",
            command=self.remove_selected_from_view,
        ).pack(side="left")
        ttk.Button(
            action_frame, text="Exit", command=self.root.destroy
        ).pack(side="right")

        self.status_label = ttk.Label(main_frame, text="")
        self.status_label.pack(anchor="w", pady=(8, 0))

    def connect_to_database(self) -> None:
        """Prompt for an Access file, select a table and load its cases."""
        database_path = filedialog.askopenfilename(
            parent=self.root,
            title="Select a Microsoft Access database",
            filetypes=(
                ("Microsoft Access databases", "*.accdb *.mdb"),
                ("All files", "*.*"),
            ),
        )
        if not database_path:
            return

        try:
            tables = get_access_tables(database_path)
            if not tables:
                messagebox.showerror(
                    "No tables found",
                    "The selected database does not contain any user tables.",
                    parent=self.root,
                )
                return

            selected_table = self._choose_table(tables)
            if selected_table is None:
                return

            loaded_cases = load_cases_from_access(database_path, selected_table)
            self.database_path = database_path
            self.table_name = selected_table
            self.cases = loaded_cases
            self.database_label.config(
                text=f"{os.path.basename(database_path)}  |  Table: {selected_table}"
            )
            self.display_cases()
            messagebox.showinfo(
                "Database loaded",
                f"Loaded {len(self.cases)} case(s) from '{selected_table}'.",
                parent=self.root,
            )
        except Exception as error:
            messagebox.showerror(
                "Database connection problem", str(error), parent=self.root
            )

    def _choose_table(self, tables: Sequence[str]) -> Optional[str]:
        """Show a small dialog allowing the user to choose a database table."""
        if len(tables) == 1:
            return tables[0]

        dialog = tk.Toplevel(self.root)
        dialog.title("Choose database table")
        dialog.transient(self.root)
        dialog.resizable(False, False)
        dialog.grab_set()

        ttk.Label(
            dialog, text="Select the table containing the DFIP case records:"
        ).pack(padx=14, pady=(14, 6))

        table_choice = ttk.Combobox(
            dialog, values=tuple(tables), state="readonly", width=38
        )
        table_choice.current(0)
        table_choice.pack(padx=14, pady=6)

        selected = {"table": None}

        def confirm() -> None:
            selected["table"] = table_choice.get()
            dialog.destroy()

        def cancel() -> None:
            dialog.destroy()

        button_frame = ttk.Frame(dialog)
        button_frame.pack(pady=(6, 14))
        ttk.Button(button_frame, text="Load", command=confirm).pack(
            side="left", padx=5
        )
        ttk.Button(button_frame, text="Cancel", command=cancel).pack(
            side="left", padx=5
        )

        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.root.wait_window(dialog)
        return selected["table"]

    def refresh_database(self) -> None:
        """Reload the currently selected table from the Access database."""
        if not self.database_path or not self.table_name:
            messagebox.showinfo(
                "No database loaded",
                "Connect to an Access database before refreshing.",
                parent=self.root,
            )
            return
        try:
            self.cases = load_cases_from_access(
                self.database_path, self.table_name
            )
            self.display_cases()
            self.status_label.config(
                text=f"Refreshed from database: {len(self.cases)} case(s)."
            )
        except Exception as error:
            messagebox.showerror(
                "Refresh failed", str(error), parent=self.root
            )

    def disconnect_database(self) -> None:
        """Clear the current database selection and displayed cases."""
        self.database_path = None
        self.table_name = None
        self.cases.clear()
        self.database_label.config(text="No database loaded")
        self.display_cases()

    def sort_and_display(self) -> None:
        """Sort the loaded cases using the selected field and direction."""
        if not self.cases:
            self.display_cases()
            messagebox.showinfo(
                "No cases",
                "Load a database containing DFIP cases before sorting.",
                parent=self.root,
            )
            return

        sort_by = SORT_OPTIONS.get(self.sort_choice.get())
        if sort_by is None:
            messagebox.showerror(
                "Invalid sort option",
                "Please choose one of the available fields.",
                parent=self.root,
            )
            return

        descending = self.order_choice.get() == "Descending"
        sort_cases(self.cases, sort_by, descending)
        self.display_cases()

        direction = "descending" if descending else "ascending"
        self.status_label.config(
            text=(
                f"Sorted {len(self.cases)} case(s) by "
                f"{self.sort_choice.get()} ({direction})."
            )
        )

    def display_cases(self) -> None:
        """Refresh the results table from the cases currently in memory."""
        for item in self.case_table.get_children():
            self.case_table.delete(item)

        # Treeview item IDs must be unique strings, so use row numbers rather
        # than CaseID (a database could contain duplicate or blank IDs).
        for row_number, case in enumerate(self.cases):
            self.case_table.insert(
                "", "end", iid=f"case_row_{row_number}",
                values=case.display_details(),
            )

        self.status_label.config(text=f"{len(self.cases)} case(s) displayed.")

    def remove_selected_from_view(self) -> None:
        """Remove one displayed record from memory without editing the database."""
        selected_items = self.case_table.selection()
        if not selected_items:
            messagebox.showinfo(
                "Select a case",
                "Select a row before choosing Remove Selected From View.",
                parent=self.root,
            )
            return

        selected_index = int(selected_items[0].replace("case_row_", ""))
        if 0 <= selected_index < len(self.cases):
            self.cases.pop(selected_index)
            self.display_cases()
            self.status_label.config(
                text="Selected case removed from this view only. "
                "The database was not changed."
            )


def main() -> None:
    """Create the application window and start the Tkinter event loop."""
    root = tk.Tk()
    DFIPCaseApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
