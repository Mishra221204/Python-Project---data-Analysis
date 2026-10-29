"""
Corporate Data Analysis & Report Builder
----------------------------------------
Features:
1. Read CSV, XLS, and XLSX files.
2. Display dataset dimensions and column names.
3. Automatically detect text and numeric columns.
4. Generate interactive GroupBy reports.
5. Sort reports in descending order.
6. Display reports inside a Tkinter Treeview.
7. Export reports to Excel or CSV.
8. Display bar, column, line, and pie charts inside the GUI.
9. Export charts as PNG images.
10. Handle errors and reset outdated outputs.

Install dependencies:
    pip install pandas matplotlib openpyxl xlrd

Run:
    python data_analysis_app.py

Create Windows executable:
    pyinstaller --onefile --windowed --name CorporateDataAnalyzer data_analysis_app.py
"""

import os
import sys
import logging
import tkinter as tk
import openpyxl
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import matplotlib

# Use Tkinter-compatible Matplotlib backend.
matplotlib.use("TkAgg")

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


# ---------------------------------------------------------
# APPLICATION CONFIGURATION
# ---------------------------------------------------------

APP_TITLE = "Corporate Data Analysis & Report Builder"
APP_VERSION = "1.0.0"

AGGREGATION_METHODS = [
    "sum",
    "mean",
    "average",
    "max",
    "min",
    "count",
    "median",
]

CHART_TYPES = [
    "Bar chart",
    "Column chart",
    "Line chart",
    "Pie chart",
]

EXPORT_FORMATS = [
    "Excel (.xlsx)",
    "CSV (.csv)",
]

MAX_PIE_SLICES = 15
TABLE_PREVIEW_ROWS = 1000

logging.basicConfig(
    filename="data_analysis_app.log",
    level=logging.ERROR,
    format="%(asctime)s - %(levelname)s - %(message)s",
)


# ---------------------------------------------------------
# MAIN APPLICATION
# ---------------------------------------------------------

class CorporateDataAnalyzer:

    def __init__(self, root):
        self.root = root

        self.root.title(f"{APP_TITLE} - v{APP_VERSION}")
        self.root.geometry("1250x820")
        self.root.minsize(950, 650)

        # Application data.
        self.file_path = None
        self.df = None
        self.report_df = None

        self.text_columns = []
        self.numeric_columns = []
        self.numeric_data = {}

        # Tkinter variables.
        self.file_var = tk.StringVar()
        self.group_var = tk.StringVar()
        self.aggregation_var = tk.StringVar(value="sum")
        self.value_var = tk.StringVar()
        self.export_format_var = tk.StringVar(
            value="Excel (.xlsx)"
        )
        self.chart_type_var = tk.StringVar(
            value="Column chart"
        )
        self.status_var = tk.StringVar(
            value="Select a CSV or Excel file to begin."
        )

        self._configure_style()
        self._build_interface()

        self.root.protocol(
            "WM_DELETE_WINDOW",
            self._on_close,
        )

    # -----------------------------------------------------
    # STYLE
    # -----------------------------------------------------

    def _configure_style(self):
        style = ttk.Style(self.root)

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg="#f9f5f3")

        style.configure(
            "TFrame",
            background="#f9f6f3",
        )

        style.configure(
            "Card.TFrame",
            background="#ffffff",
            relief="flat",
        )

        style.configure(
            "TLabel",
            background="#f3f5f9",
            foreground="#263248",
            font=("Segoe UI", 10),
        )

        style.configure(
            "Title.TLabel",
            font=("Segoe UI", 20, "bold"),
            foreground="#17365d",
            background="#f3f5f9",
        )

        style.configure(
            "Subtitle.TLabel",
            font=("Segoe UI", 10),
            foreground="#5e6b7d",
            background="#f3f5f9",
        )

        style.configure(
            "CardTitle.TLabel",
            font=("Segoe UI", 11, "bold"),
            foreground="#17365d",
            background="#ffffff",
        )

        style.configure(
            "TButton",
            font=("Segoe UI", 10),
            padding=(10, 7),
        )

        style.configure(
            "Primary.TButton",
            font=("Segoe UI", 10, "bold"),
            padding=(12, 8),
        )

        style.configure(
            "Treeview",
            font=("Segoe UI", 9),
            rowheight=27,
            background="#ffffff",
            fieldbackground="#ffffff",
        )

        style.configure(
            "Treeview.Heading",
            font=("Segoe UI", 9, "bold"),
            padding=7,
        )

        style.map(
            "Treeview",
            background=[
                ("selected", "#cfe2ff")
            ],
            foreground=[
                ("selected", "#17365d")
            ],
        )

    # -----------------------------------------------------
    # BUILD GUI
    # -----------------------------------------------------

    def _build_interface(self):
        main = ttk.Frame(self.root, padding=18)
        main.pack(fill="both", expand=True)

        # Header.
        header = ttk.Frame(main)
        header.pack(fill="x", pady=(0, 14))

        ttk.Label(
            header,
            text="Corporate Data Analyzer",
            style="Title.TLabel",
        ).pack(anchor="w")

        ttk.Label(
            header,
            text=(
                "Import data, build reports, visualize results "
                "and export your analysis."
            ),
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        # File selection section.
        file_card = ttk.LabelFrame(
            main,
            text="  1. Data Source  ",
            padding=12,
        )
        file_card.pack(fill="x", pady=(0, 10))

        file_row = ttk.Frame(file_card)
        file_row.pack(fill="x")

        ttk.Entry(
            file_row,
            textvariable=self.file_var,
            state="readonly",
        ).pack(side="left", fill="x", expand=True, padx=(0, 8))

        ttk.Button(
            file_row,
            text="Browse",
            command=self.select_file,
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            file_row,
            text="Read",
            style="Primary.TButton",
            command=self.read_file,
        ).pack(side="left")

        # Dataset information.
        self.dataset_info = ttk.Label(
            file_card,
            text="No dataset loaded.",
        )
        self.dataset_info.pack(anchor="w", pady=(9, 0))

        self.columns_info = ttk.Label(
            file_card,
            text="",
            wraplength=1100,
            justify="left",
        )
        self.columns_info.pack(anchor="w", pady=(4, 0))

        # Report builder.
        report_card = ttk.LabelFrame(
            main,
            text="  2. Report Builder  ",
            padding=12,
        )
        report_card.pack(fill="x", pady=(0, 10))

        controls = ttk.Frame(report_card)
        controls.pack(fill="x")

        # Group By.
        ttk.Label(
            controls,
            text="Group By Column",
        ).grid(row=0, column=0, sticky="w", padx=(0, 8))

        self.group_combo = ttk.Combobox(
            controls,
            textvariable=self.group_var,
            state="readonly",
            width=23,
        )
        self.group_combo.grid(
            row=1,
            column=0,
            sticky="ew",
            padx=(0, 12),
            pady=(5, 0),
        )

        # Aggregation.
        ttk.Label(
            controls,
            text="Aggregation Method",
        ).grid(row=0, column=1, sticky="w", padx=(0, 8))

        self.aggregation_combo = ttk.Combobox(
            controls,
            textvariable=self.aggregation_var,
            state="readonly",
            values=AGGREGATION_METHODS,
            width=18,
        )
        self.aggregation_combo.grid(
            row=1,
            column=1,
            sticky="ew",
            padx=(0, 12),
            pady=(5, 0),
        )

        # Value column.
        ttk.Label(
            controls,
            text="Value Column",
        ).grid(row=0, column=2, sticky="w", padx=(0, 8))

        self.value_combo = ttk.Combobox(
            controls,
            textvariable=self.value_var,
            state="readonly",
            width=23,
        )
        self.value_combo.grid(
            row=1,
            column=2,
            sticky="ew",
            padx=(0, 12),
            pady=(5, 0),
        )

        ttk.Button(
            controls,
            text="Preview Report",
            style="Primary.TButton",
            command=self.preview_report,
        ).grid(
            row=1,
            column=3,
            sticky="ew",
            pady=(5, 0),
        )

        controls.columnconfigure(0, weight=1)
        controls.columnconfigure(2, weight=1)

        # Export report controls.
        export_row = ttk.Frame(report_card)
        export_row.pack(fill="x", pady=(12, 0))

        ttk.Label(
            export_row,
            text="Export format:",
        ).pack(side="left", padx=(0, 8))

        ttk.Combobox(
            export_row,
            textvariable=self.export_format_var,
            state="readonly",
            values=EXPORT_FORMATS,
            width=18,
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            export_row,
            text="Export Report",
            command=self.export_report,
        ).pack(side="left")

        # Chart builder.
        chart_card = ttk.LabelFrame(
            main,
            text="  3. Chart Builder  ",
            padding=12,
        )
        chart_card.pack(fill="x", pady=(0, 10))

        chart_row = ttk.Frame(chart_card)
        chart_row.pack(fill="x")

        ttk.Label(
            chart_row,
            text="Chart Type:",
        ).pack(side="left", padx=(0, 8))

        ttk.Combobox(
            chart_row,
            textvariable=self.chart_type_var,
            state="readonly",
            values=CHART_TYPES,
            width=18,
        ).pack(side="left", padx=(0, 10))

        ttk.Button(
            chart_row,
            text="Preview Chart",
            style="Primary.TButton",
            command=self.preview_chart,
        ).pack(side="left", padx=(0, 8))

        ttk.Button(
            chart_row,
            text="Export Chart (PNG)",
            command=self.export_chart,
        ).pack(side="left")

        # Results area with report and chart tabs.
        results_frame = ttk.Frame(main)
        results_frame.pack(
            fill="both",
            expand=True,
            pady=(0, 8),
        )

        self.notebook = ttk.Notebook(results_frame)
        self.notebook.pack(fill="both", expand=True)

        self.table_tab = ttk.Frame(self.notebook)
        self.chart_tab = ttk.Frame(self.notebook)

        self.notebook.add(
            self.table_tab,
            text="Report Results",
        )

        self.notebook.add(
            self.chart_tab,
            text="Chart Visualization",
        )

        self._build_table()
        self._build_chart_area()

        # Status bar.
        status_bar = ttk.Label(
            main,
            textvariable=self.status_var,
            anchor="w",
            relief="sunken",
            padding=(8, 5),
        )
        status_bar.pack(fill="x")

    # -----------------------------------------------------
    # REPORT TABLE
    # -----------------------------------------------------

    def _build_table(self):
        self.table_container = ttk.Frame(self.table_tab)
        self.table_container.pack(
            fill="both",
            expand=True,
            padx=6,
            pady=6,
        )

        self.table_container.columnconfigure(0, weight=1)
        self.table_container.rowconfigure(0, weight=1)

        self.tree = ttk.Treeview(
            self.table_container,
            show="headings",
        )

        vertical_scroll = ttk.Scrollbar(
            self.table_container,
            orient="vertical",
            command=self.tree.yview,
        )

        horizontal_scroll = ttk.Scrollbar(
            self.table_container,
            orient="horizontal",
            command=self.tree.xview,
        )

        self.tree.configure(
            yscrollcommand=vertical_scroll.set,
            xscrollcommand=horizontal_scroll.set,
        )

        self.tree.grid(
            row=0,
            column=0,
            sticky="nsew",
        )

        vertical_scroll.grid(
            row=0,
            column=1,
            sticky="ns",
        )

        horizontal_scroll.grid(
            row=1,
            column=0,
            sticky="ew",
        )

        self.table_hint = ttk.Label(
            self.table_tab,
            text="Generate a report to display the results here.",
        )
        self.table_hint.pack(anchor="w", padx=8, pady=(0, 5))

    def _clear_table(self):
        self.tree.delete(*self.tree.get_children())

        self.tree["columns"] = ()

        for item in self.tree.get_children():
            self.tree.delete(item)

    def _display_report(self):
        self._clear_table()

        if self.report_df is None:
            return

        columns = list(self.report_df.columns)
        self.tree["columns"] = columns

        for column in columns:
            self.tree.heading(
                column,
                text=str(column),
                anchor="w",
            )

            self.tree.column(
                column,
                width=180,
                minwidth=100,
                anchor="w",
                stretch=True,
            )

        # Limit displayed rows for GUI responsiveness.
        display_df = self.report_df.head(TABLE_PREVIEW_ROWS)

        for row in display_df.itertuples(
            index=False,
            name=None,
        ):
            values = [
                self._format_cell(value)
                for value in row
            ]

            self.tree.insert(
                "",
                "end",
                values=values,
            )

        total = len(self.report_df)

        if total > TABLE_PREVIEW_ROWS:
            self.table_hint.configure(
                text=(
                    f"Showing the first {TABLE_PREVIEW_ROWS:,} "
                    f"of {total:,} report rows. "
                    "Export the report to access every row."
                )
            )
        else:
            self.table_hint.configure(
                text=f"Report contains {total:,} rows."
            )

    @staticmethod
    def _format_cell(value):
        if pd.isna(value):
            return ""

        if isinstance(value, float):
            return f"{value:,.4f}".rstrip("0").rstrip(".")

        return str(value)

    # -----------------------------------------------------
    # CHART AREA
    # -----------------------------------------------------

    def _build_chart_area(self):
        self.chart_container = ttk.Frame(self.chart_tab)
        self.chart_container.pack(
            fill="both",
            expand=True,
            padx=6,
            pady=6,
        )

        self.figure = Figure(
            figsize=(9, 5),
            dpi=100,
            constrained_layout=True,
        )

        self.ax = self.figure.add_subplot(111)

        self.ax.set_title(
            "Preview Chart",
            fontsize=14,
            fontweight="bold",
        )

        self.ax.text(
            0.5,
            0.5,
            "Generate a report, then click Preview Chart.",
            ha="center",
            va="center",
            transform=self.ax.transAxes,
        )

        self.canvas = FigureCanvasTkAgg(
            self.figure,
            master=self.chart_container,
        )

        self.canvas.get_tk_widget().pack(
            fill="both",
            expand=True,
        )

        self.chart_created = False

    def _clear_chart(self):
        self.figure.clear()
        self.ax = self.figure.add_subplot(111)

        self.ax.set_title("Preview Chart")

        self.canvas.draw_idle()

        self.chart_created = False

    # -----------------------------------------------------
    # FILE SELECTION
    # -----------------------------------------------------

    def select_file(self):
        selected = filedialog.askopenfilename(
            title="Select CSV or Excel File",
            filetypes=[
                (
                    "Supported Data Files",
                    "*.csv *.xlsx *.xls",
                ),
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls"),
                ("All Files", "*.*"),
            ],
        )

        if not selected:
            return

        self.file_path = Path(selected)
        self.file_var.set(str(self.file_path))

        # Invalidate any previously loaded dataset.
        self.df = None
        self.report_df = None

        self.text_columns = []
        self.numeric_columns = []
        self.numeric_data = {}

        self.group_combo["values"] = ()
        self.value_combo["values"] = ()

        self.group_var.set("")
        self.value_var.set("")

        self.dataset_info.configure(
            text="File selected. Click Read to load the dataset."
        )

        self.columns_info.configure(text="")

        self._clear_table()
        self._clear_chart()

        self.status_var.set(
            f"Selected: {self.file_path.name}"
        )

    # -----------------------------------------------------
    # READ DATASET
    # -----------------------------------------------------

    def read_file(self):
        if not self.file_path:
            messagebox.showerror(
                "File Required",
                "Please select a CSV or Excel file first.",
            )
            return

        if not self.file_path.exists():
            messagebox.showerror(
                "File Not Found",
                "The selected file no longer exists.",
            )
            return

        try:
            self.status_var.set("Reading dataset...")
            self.root.update_idletasks()

            suffix = self.file_path.suffix.lower()

            if suffix == ".csv":
                # utf-8-sig also handles a UTF-8 BOM.
                try:
                    loaded_df = pd.read_csv(
                        self.file_path,
                        encoding="utf-8-sig",
                        low_memory=False,
                    )
                except UnicodeDecodeError:
                    loaded_df = pd.read_csv(
                        self.file_path,
                        encoding="cp1252",
                        low_memory=False,
                    )

            elif suffix in (".xlsx", ".xls"):
                loaded_df = pd.read_excel(
                    self.file_path,
                    sheet_name=0,
                )

            else:
                raise ValueError(
                    "Unsupported file format. "
                    "Choose a CSV, XLS, or XLSX file."
                )

            if len(loaded_df.columns) == 0:
                raise ValueError(
                    "The selected file contains no columns."
                )

            # Remove leading/trailing spaces from headings.
            # Duplicate headings are handled by Pandas.
            loaded_df.columns = [
                str(column).strip()
                for column in loaded_df.columns
            ]

            self.df = loaded_df

            self._detect_column_types()
            self._update_dataset_information()
            self._reset_report_outputs()

            self.status_var.set(
                f"Dataset loaded successfully: "
                f"{len(self.df):,} rows, "
                f"{len(self.df.columns):,} columns."
            )

        except Exception as exc:
            logging.exception("Failed to read input file.")

            self.df = None
            self.report_df = None

            messagebox.showerror(
                "Read Error",
                f"Unable to read the selected file.\n\n{exc}",
            )

            self.status_var.set("Failed to read dataset.")

    # -----------------------------------------------------
    # AUTOMATIC COLUMN DETECTION
    # -----------------------------------------------------

    def _detect_column_types(self):
        self.text_columns = []
        self.numeric_columns = []
        self.numeric_data = {}

        for column in self.df.columns:
            series = self.df[column]

            # Native numeric columns.
            if pd.api.types.is_numeric_dtype(series.dtype):
                if pd.api.types.is_bool_dtype(series.dtype):
                    self.text_columns.append(column)
                    continue

                self.numeric_columns.append(column)
                self.numeric_data[column] = series
                continue

            # Dates and times should not be treated as numbers.
            if (
                pd.api.types.is_datetime64_any_dtype(series.dtype)
                or pd.api.types.is_timedelta64_dtype(series.dtype)
            ):
                continue

            # Attempt to recognize numeric-looking text columns.
            if (
                pd.api.types.is_object_dtype(series.dtype)
                or isinstance(series.dtype, pd.StringDtype)
            ):
                converted = self._convert_numeric_text(series)

                nonempty = series.notna() & (
                    series.astype("string").str.strip() != ""
                )

                nonempty_count = int(nonempty.sum())

                # A column is considered numeric if every
                # nonempty value can be converted to a number.
                if (
                    nonempty_count > 0
                    and converted[nonempty].notna().all()
                ):
                    self.numeric_columns.append(column)
                    self.numeric_data[column] = converted
                else:
                    self.text_columns.append(column)

            else:
                # Other categorical types can be grouped by.
                self.text_columns.append(column)

        # GroupBy needs a grouping column, and numeric
        # aggregations need a numeric value column.
        self.group_combo["values"] = self.text_columns
        self.value_combo["values"] = self.numeric_columns

        self.group_var.set(
            self.text_columns[0]
            if self.text_columns
            else ""
        )

        self.value_var.set(
            self.numeric_columns[0]
            if self.numeric_columns
            else ""
        )

    @staticmethod
    def _convert_numeric_text(series):
        """
        Convert common numeric-looking strings into numbers.

        Supports examples such as:
            "1,200"
            "$1,200.50"
            " 450 "
            "(250)"   -> -250

        Currency symbols and thousands separators are removed.
        If a column contains nonnumeric text, it will not
        qualify as a fully numeric text column.
        """
        values = series.astype("string").str.strip()

        values = values.replace("", pd.NA)

        # Accounting-style negative numbers: (250) -> -250.
        accounting_negative = values.str.match(
            r"^\(.*\)$",
            na=False,
        )

        values = values.str.replace(
            r"^\((.*)\)$",
            r"-\1",
            regex=True,
        )

        # Remove common currency symbols and separators.
        values = values.str.replace(
            r"[$€£₹,\s]",
            "",
            regex=True,
        )

        # Allow common percent notation, treating 12% as 12.
        values = values.str.replace(
            "%",
            "",
            regex=False,
        )

        converted = pd.to_numeric(
            values,
            errors="coerce",
        )

        return converted

    def _update_dataset_information(self):
        row_count = len(self.df)
        column_count = len(self.df.columns)

        self.dataset_info.configure(
            text=(
                f"Rows: {row_count:,}  |  "
                f"Columns: {column_count:,}  |  "
                f"Text columns: {len(self.text_columns)}  |  "
                f"Numeric columns: {len(self.numeric_columns)}"
            )
        )

        headings = ", ".join(
            str(column)
            for column in self.df.columns
        )

        self.columns_info.configure(
            text=f"Column headings: {headings}"
        )

    # -----------------------------------------------------
    # RESET PREVIOUS OUTPUTS
    # -----------------------------------------------------

    def _reset_report_outputs(self):
        self.report_df = None

        self._clear_table()
        self._clear_chart()

        self.status_var.set(
            "Dataset loaded. Configure your report."
        )

    # -----------------------------------------------------
    # VALIDATION
    # -----------------------------------------------------

    def _validate_dataset(self):
        if not self.file_path:
            messagebox.showerror(
                "File Required",
                "Please select a file first.",
            )
            return False

        if self.df is None:
            messagebox.showerror(
                "Dataset Not Loaded",
                "Click Read before generating a report.",
            )
            return False

        return True

    # -----------------------------------------------------
    # REPORT GENERATION
    # -----------------------------------------------------

    def preview_report(self):
        if not self._validate_dataset():
            return

        group_column = self.group_var.get()
        aggregation = self.aggregation_var.get()
        value_column = self.value_var.get()

        if not group_column:
            messagebox.showerror(
                "Group By Required",
                "Please select a Group By Column.",
            )
            return

        if not aggregation:
            messagebox.showerror(
                "Aggregation Required",
                "Please select an aggregation method.",
            )
            return

        if not value_column:
            messagebox.showerror(
                "Value Column Required",
                "Please select a numeric Value Column.",
            )
            return

        if group_column not in self.text_columns:
            messagebox.showerror(
                "Invalid Group By Column",
                "Select a valid text or categorical column.",
            )
            return

        if value_column not in self.numeric_columns:
            messagebox.showerror(
                "Invalid Value Column",
                "Select a valid numeric column.",
            )
            return

        # Invalidate old outputs before attempting a new report.
        self.report_df = None
        self._clear_table()
        self._clear_chart()

        try:
            self.status_var.set("Generating report...")
            self.root.update_idletasks()

            numeric_series = self.numeric_data[value_column]

            working_df = pd.DataFrame({
                group_column: self.df[group_column],
                value_column: numeric_series,
            })

            # Treat average as an alias for mean.
            normalized_aggregation = (
                "mean"
                if aggregation == "average"
                else aggregation
            )

            grouped = working_df.groupby(
                group_column,
                dropna=False,
                sort=False,
                observed=True,
            )[value_column]

            result = grouped.agg(normalized_aggregation)

            # Always produce a DataFrame with predictable headings.
            result_df = result.reset_index()

            result_df.columns = [
                group_column,
                f"{aggregation.title()} of {value_column}",
            ]

            result_value_column = result_df.columns[1]

            # Descending sort by aggregated result.
            result_df = result_df.sort_values(
                by=result_value_column,
                ascending=False,
                na_position="last",
                kind="stable",
            ).reset_index(drop=True)

            self.report_df = result_df

            self._display_report()

            self.notebook.select(self.table_tab)

            self.status_var.set(
                f"Report generated successfully: "
                f"{len(result_df):,} groups."
            )

        except Exception as exc:
            logging.exception("Failed to generate report.")

            self.report_df = None
            self._clear_table()
            self._clear_chart()

            messagebox.showerror(
                "Report Error",
                f"Unable to generate the report.\n\n{exc}",
            )

            self.status_var.set("Report generation failed.")

    # -----------------------------------------------------
    # EXPORT REPORT
    # -----------------------------------------------------

    def export_report(self):
        if not self._validate_dataset():
            return

        if self.report_df is None:
            messagebox.showerror(
                "No Report",
                "Please click Preview Report before exporting.",
            )
            return

        try:
            export_format = self.export_format_var.get()

            if export_format == "Excel (.xlsx)":
                output_path = self.file_path.with_name(
                    f"{self.file_path.stem}_report.xlsx"
                )

                self.report_df.to_excel(
                    output_path,
                    index=False,
                    engine="openpyxl",
                )

            elif export_format == "CSV (.csv)":
                output_path = self.file_path.with_name(
                    f"{self.file_path.stem}_report.csv"
                )

                self.report_df.to_csv(
                    output_path,
                    index=False,
                    encoding="utf-8-sig",
                )

            else:
                messagebox.showerror(
                    "Export Format",
                    "Please select a supported export format.",
                )
                return

            self.status_var.set(
                f"Report exported: {output_path.name}"
            )

            messagebox.showinfo(
                "Export Successful",
                f"Report saved successfully.\n\n{output_path}",
            )

        except PermissionError:
            messagebox.showerror(
                "Permission Error",
                "Unable to save the report. "
                "Close the output file if it is already open, "
                "then try again.",
            )

        except Exception as exc:
            logging.exception("Report export failed.")

            messagebox.showerror(
                "Export Error",
                f"Unable to export the report.\n\n{exc}",
            )

    # -----------------------------------------------------
    # CHART GENERATION
    # -----------------------------------------------------

    def preview_chart(self):
        if not self._validate_dataset():
            return

        if self.report_df is None:
            messagebox.showerror(
                "No Report",
                "Please generate a report before creating a chart.",
            )
            return

        try:
            self.figure.clear()
            self.ax = self.figure.add_subplot(111)

            group_column = self.report_df.columns[0]
            result_column = self.report_df.columns[1]

            chart_df = self.report_df.copy()

            if chart_df.empty:
                raise ValueError(
                    "The report contains no data to visualize."
                )

            chart_type = self.chart_type_var.get()

            labels = (
                chart_df[group_column]
                .astype("string")
                .fillna("(Missing)")
                .astype(str)
            )

            values = pd.to_numeric(
                chart_df[result_column],
                errors="coerce",
            )

            # Replace missing aggregate values for plotting.
            valid = values.notna()
            labels = labels[valid]
            values = values[valid]

            if len(values) == 0:
                raise ValueError(
                    "No valid numeric report values are available."
                )

            if chart_type == "Bar chart":
                # Horizontal bar chart.
                self.ax.barh(
                    labels.iloc[::-1],
                    values.iloc[::-1],
                )
                self.ax.set_xlabel(result_column)
                self.ax.set_ylabel(group_column)

            elif chart_type == "Column chart":
                self.ax.bar(
                    labels,
                    values,
                )
                self.ax.set_xlabel(group_column)
                self.ax.set_ylabel(result_column)
                self.ax.tick_params(
                    axis="x",
                    labelrotation=45,
                )

            elif chart_type == "Line chart":
                self.ax.plot(
                    labels,
                    values,
                    marker="o",
                    linewidth=2,
                )
                self.ax.set_xlabel(group_column)
                self.ax.set_ylabel(result_column)
                self.ax.tick_params(
                    axis="x",
                    labelrotation=45,
                )
                self.ax.grid(
                    True,
                    axis="y",
                    linestyle="--",
                    alpha=0.4,
                )

            elif chart_type == "Pie chart":
                # Pie charts require nonnegative values and
                # a positive total.
                if (values < 0).any():
                    raise ValueError(
                        "Pie charts cannot represent negative "
                        "values. Choose a bar, column, or line chart."
                    )

                if values.sum() <= 0:
                    raise ValueError(
                        "Pie charts require a positive total."
                    )

                # Keep the largest categories visible.
                # Smaller categories are combined into Other.
                if len(values) > MAX_PIE_SLICES:
                    pie_values = values.iloc[:MAX_PIE_SLICES].copy()
                    pie_labels = labels.iloc[:MAX_PIE_SLICES].tolist()

                    other_value = values.iloc[MAX_PIE_SLICES:].sum()

                    if other_value > 0:
                        pie_values = pd.concat([
                            pie_values,
                            pd.Series([other_value]),
                        ])
                        pie_labels.append("Other")

                    values = pie_values
                    labels = pd.Series(pie_labels)

                self.ax.pie(
                    values,
                    labels=labels,
                    autopct="%1.1f%%",
                    startangle=90,
                    pctdistance=0.7,
                )
                self.ax.axis("equal")

            else:
                raise ValueError("Unsupported chart type selected.")

            self.ax.set_title(
                f"{chart_type}: {result_column}",
                fontsize=13,
                fontweight="bold",
            )

            if chart_type != "Pie chart":
                self.figure.tight_layout()

            self.canvas.draw()

            self.chart_created = True
            self.notebook.select(self.chart_tab)

            self.status_var.set(
                f"{chart_type} generated successfully."
            )

        except Exception as exc:
            logging.exception("Chart generation failed.")

            self._clear_chart()

            messagebox.showerror(
                "Chart Error",
                f"Unable to generate the chart.\n\n{exc}",
            )

            self.status_var.set("Chart generation failed.")

    # -----------------------------------------------------
    # EXPORT CHART
    # -----------------------------------------------------

    def export_chart(self):
        if not self._validate_dataset():
            return

        if self.report_df is None:
            messagebox.showerror(
                "No Report",
                "Please generate a report before exporting a chart.",
            )
            return

        if not self.chart_created:
            messagebox.showerror(
                "No Chart",
                "Please click Preview Chart before exporting.",
            )
            return

        try:
            chart_name = (
                self.chart_type_var.get()
                .lower()
                .replace(" ", "_")
            )

            output_path = self.file_path.with_name(
                f"{self.file_path.stem}_{chart_name}.png"
            )

            self.figure.savefig(
                output_path,
                dpi=300,
                bbox_inches="tight",
                facecolor="white",
            )

            self.status_var.set(
                f"Chart exported: {output_path.name}"
            )

            messagebox.showinfo(
                "Chart Export Successful",
                f"Chart saved successfully.\n\n{output_path}",
            )

        except PermissionError:
            messagebox.showerror(
                "Permission Error",
                "Unable to save the chart. "
                "Check whether the output file is open "
                "or the folder is writable.",
            )

        except Exception as exc:
            logging.exception("Chart export failed.")

            messagebox.showerror(
                "Chart Export Error",
                f"Unable to export the chart.\n\n{exc}",
            )

    # -----------------------------------------------------
    # CLEAN SHUTDOWN
    # -----------------------------------------------------

    def _on_close(self):
        try:
            self.figure.clear()
        except Exception:
            pass

        self.root.destroy()


# ---------------------------------------------------------
# APPLICATION ENTRY POINT
# ---------------------------------------------------------

def main():
    root = tk.Tk()

    app = CorporateDataAnalyzer(root)

    root.mainloop()


if __name__ == "__main__":
    main()