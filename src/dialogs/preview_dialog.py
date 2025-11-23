"""
Data Preview Dialog - Preview and edit file data before import
"""

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
from src.file_processor import get_dataframes, infer_column_type


class DataPreviewDialog:
    """Dialog for previewing file data and editing column names/types"""
    def __init__(self, parent, main_app, file_path):
        self.main_app = main_app
        self.file_path = file_path
        self.filename = os.path.basename(file_path)

        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Data Preview - {self.filename}")
        self.dialog.geometry("1200x800")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Load file data
        try:
            self.dataframes = get_dataframes(file_path)
            self.main_app.log_message(f"Loaded {len(self.dataframes)} sheet(s) from {self.filename}", "INFO")
        except Exception as e:
            self.main_app.log_message(f"Failed to load file: {e}", "ERROR")
            messagebox.showerror("Error", f"Failed to load file:\n{e}")
            self.dialog.destroy()
            return

        # Initialize overrides for this file if not exist
        if file_path not in self.main_app.column_overrides:
            self.main_app.column_overrides[file_path] = {}

        # Main frame
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Sheet selector (if multiple sheets)
        if len(self.dataframes) > 1:
            sheet_frame = ttk.Frame(main_frame)
            sheet_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

            ttk.Label(sheet_frame, text="Sheet:", font=("Helvetica", 10, "bold")).pack(side=tk.LEFT, padx=(0, 10))

            self.sheet_var = tk.StringVar(value=list(self.dataframes.keys())[0])
            sheet_combo = ttk.Combobox(
                sheet_frame,
                textvariable=self.sheet_var,
                values=list(self.dataframes.keys()),
                state="readonly",
                width=30
            )
            sheet_combo.pack(side=tk.LEFT)
            sheet_combo.bind('<<ComboboxSelected>>', lambda e: self.load_sheet())
        else:
            self.sheet_var = tk.StringVar(value=list(self.dataframes.keys())[0])

        # Content frame (will hold stats and data grid)
        self.content_frame = ttk.Frame(main_frame)
        self.content_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(1, weight=1)

        # Bottom buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        ttk.Button(button_frame, text="Apply Changes", command=self.apply_changes).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT)
        ttk.Button(button_frame, text="Reset to Defaults", command=self.reset_defaults).pack(side=tk.LEFT)

        # Load first sheet
        self.load_sheet()

    def load_sheet(self):
        """Load and display the selected sheet"""
        # Clear content frame
        for widget in self.content_frame.winfo_children():
            widget.destroy()

        sheet_name = self.sheet_var.get()
        df = self.dataframes[sheet_name]

        # Reset column edit widgets for new sheet
        self.column_name_vars = {}
        self.column_type_vars = {}

        # Statistics frame
        stats_frame = ttk.LabelFrame(self.content_frame, text="Statistics", padding="10")
        stats_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        stats_frame.columnconfigure(1, weight=1)

        # Display statistics
        row_count = len(df)
        col_count = len(df.columns)
        null_counts = df.isna().sum()
        total_nulls = null_counts.sum()

        ttk.Label(stats_frame, text=f"Rows: {row_count:,}", font=("Courier", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        ttk.Label(stats_frame, text=f"Columns: {col_count}", font=("Courier", 10, "bold")).grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        ttk.Label(stats_frame, text=f"Total NULL values: {total_nulls:,}", font=("Courier", 10, "bold")).grid(row=0, column=2, sticky=tk.W)

        # Preview frame with scrollable area
        preview_frame = ttk.LabelFrame(self.content_frame, text=f"Data Preview (First 20 rows) - Sheet: {sheet_name}", padding="10")
        preview_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)

        # Create canvas and scrollbars for scrollable content
        canvas = tk.Canvas(preview_frame, bg='white')
        canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        h_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=canvas.xview)
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))

        v_scrollbar = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=canvas.yview)
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))

        canvas.config(xscrollcommand=h_scrollbar.set, yscrollcommand=v_scrollbar.set)

        # Inner frame for content
        inner_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=inner_frame, anchor=tk.NW)

        # Get existing overrides for this sheet
        sheet_overrides = self.main_app.column_overrides.get(self.file_path, {}).get(sheet_name, {})
        column_name_overrides = sheet_overrides.get('columns', {})
        column_type_overrides = sheet_overrides.get('types', {})

        # Available SQL types
        sql_types = ["NVARCHAR(MAX)", "BIGINT", "FLOAT", "INT", "DECIMAL(18,2)", "DATE", "DATETIME", "BIT"]

        # Create unified grid with column headers and data
        # Display first 20 rows
        preview_df = df.head(20)

        # Fixed column width for alignment
        col_width = 150

        for col_idx, col_name in enumerate(df.columns):
            # Create a container frame for each column (header + data)
            column_container = ttk.Frame(inner_frame)
            column_container.grid(row=0, column=col_idx, sticky=(tk.N, tk.S), padx=2)

            # Header section with edit controls
            col_frame = ttk.Frame(column_container, relief=tk.RIDGE, borderwidth=1)
            col_frame.pack(fill=tk.X, pady=(0, 2))

            # Column name editor
            ttk.Label(col_frame, text="Column Name:", font=("Arial", 8)).pack(anchor=tk.W, padx=2, pady=(2, 0))
            name_var = tk.StringVar(value=column_name_overrides.get(col_name, col_name))
            name_entry = ttk.Entry(col_frame, textvariable=name_var, width=20, font=("Arial", 9))
            name_entry.pack(fill=tk.X, padx=2, pady=(0, 5))
            self.column_name_vars[col_name] = name_var

            # Detected type display
            detected_type = infer_column_type(df[col_name], col_name)
            ttk.Label(col_frame, text=f"Detected: {detected_type}", font=("Arial", 8), foreground="gray").pack(anchor=tk.W, padx=2)

            # Type selector
            ttk.Label(col_frame, text="SQL Type:", font=("Arial", 8)).pack(anchor=tk.W, padx=2, pady=(5, 0))
            type_var = tk.StringVar(value=column_type_overrides.get(col_name, detected_type))
            type_combo = ttk.Combobox(col_frame, textvariable=type_var, values=sql_types, state="readonly", width=18, font=("Arial", 8))
            type_combo.pack(fill=tk.X, padx=2, pady=(0, 5))
            self.column_type_vars[col_name] = type_var

            # NULL count for this column
            null_count = null_counts[col_name]
            null_pct = (null_count / row_count * 100) if row_count > 0 else 0
            ttk.Label(col_frame, text=f"NULLs: {null_count} ({null_pct:.1f}%)", font=("Arial", 7), foreground="red" if null_count > 0 else "green").pack(anchor=tk.W, padx=2, pady=(0, 2))

            # Data cells for this column
            for row_idx, value in enumerate(preview_df[col_name]):
                cell_frame = ttk.Frame(column_container, relief=tk.SOLID, borderwidth=1)
                cell_frame.pack(fill=tk.X, pady=1)

                # Format value
                if pd.isna(value):
                    display_value = "NULL"
                    fg_color = "gray"
                else:
                    display_value = str(value)
                    if len(display_value) > 30:
                        display_value = display_value[:27] + "..."
                    fg_color = "black"

                label = tk.Label(cell_frame, text=display_value, font=("Courier", 8), anchor=tk.W, fg=fg_color, bg="white", padx=5, pady=2, width=20)
                label.pack(fill=tk.BOTH, expand=True)

        # Update scroll region
        inner_frame.update_idletasks()
        canvas.config(scrollregion=canvas.bbox("all"))

        # Configure canvas resize
        def configure_canvas(event):
            canvas.config(scrollregion=canvas.bbox("all"))

        inner_frame.bind("<Configure>", configure_canvas)

    def apply_changes(self):
        """Apply column name and type overrides"""
        sheet_name = self.sheet_var.get()

        # Collect overrides
        column_name_map = {}
        column_type_map = {}

        df = self.dataframes[sheet_name]
        for original_col in df.columns:
            new_name = self.column_name_vars[original_col].get().strip()
            new_type = self.column_type_vars[original_col].get()

            if new_name and new_name != original_col:
                column_name_map[original_col] = new_name

            # Always store type override
            column_type_map[original_col] = new_type

        # Store in main app
        if self.file_path not in self.main_app.column_overrides:
            self.main_app.column_overrides[self.file_path] = {}

        self.main_app.column_overrides[self.file_path][sheet_name] = {
            'columns': column_name_map,
            'types': column_type_map
        }

        self.main_app.log_message(f"Applied overrides for sheet '{sheet_name}': {len(column_name_map)} column renames, {len(column_type_map)} type overrides", "SUCCESS")

        # If multiple sheets, stay open; otherwise close
        if len(self.dataframes) > 1:
            messagebox.showinfo("Success", f"Changes applied for sheet '{sheet_name}'.\n\nYou can now select another sheet or close this dialog.")
        else:
            messagebox.showinfo("Success", "Changes applied successfully!")
            self.dialog.destroy()

    def reset_defaults(self):
        """Reset all overrides for current sheet to defaults"""
        sheet_name = self.sheet_var.get()
        df = self.dataframes[sheet_name]

        if messagebox.askyesno("Reset to Defaults", f"Reset all column names and types to detected defaults for sheet '{sheet_name}'?"):
            # Clear all overrides for this sheet
            for col_name in df.columns:
                self.column_name_vars[col_name].set(col_name)
                detected_type = infer_column_type(df[col_name], col_name)
                self.column_type_vars[col_name].set(detected_type)

            # Remove from stored overrides
            if self.file_path in self.main_app.column_overrides and sheet_name in self.main_app.column_overrides[self.file_path]:
                del self.main_app.column_overrides[self.file_path][sheet_name]

            self.main_app.log_message(f"Reset sheet '{sheet_name}' to defaults", "INFO")
            messagebox.showinfo("Reset Complete", f"Sheet '{sheet_name}' has been reset to defaults.")

    def cancel(self):
        """Close dialog without applying changes"""
        self.dialog.destroy()
