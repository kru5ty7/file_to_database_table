"""
Data Preview Dialog - Preview and edit file data before import
"""

import tkinter as tk
from tkinter import ttk, messagebox
import customtkinter as ctk
import pandas as pd
import os
from src.file_processor import get_dataframes, infer_column_type


class DataPreviewDialog:
    """Dialog for previewing file data and editing column names/types"""
    def __init__(self, parent, main_app, file_path):
        self.main_app = main_app
        self.file_path = file_path
        self.filename = os.path.basename(file_path)
        self.is_csv = file_path.lower().endswith('.csv')

        self.dialog = ctk.CTkToplevel(parent)
        self.dialog.title(f"Data Preview - {self.filename}")
        self.dialog.geometry("1200x800")

        # Enable minimize and maximize buttons (remove transient to allow window controls)
        # self.dialog.transient(parent)  # Commented out to enable min/max buttons
        self.dialog.resizable(True, True)  # Allow window resizing
        self.dialog.grab_set()

        # Get delimiter preference for CSV files
        self.current_delimiter = self.main_app.csv_delimiters.get(file_path, ',')

        # Load file data
        try:
            self.dataframes = get_dataframes(file_path, delimiter=self.current_delimiter)
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
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Delimiter selector for CSV files (row 0)
        current_row = 0
        if self.is_csv:
            delimiter_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            delimiter_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
            current_row += 1

            ctk.CTkLabel(delimiter_frame, text="CSV Delimiter:", font=ctk.CTkFont(size=12, weight="bold")).pack(side=tk.LEFT, padx=(0, 10))

            self.delimiter_var = tk.StringVar(value=self.current_delimiter)
            delimiter_options = [
                (",", "Comma (,)"),
                (";", "Semicolon (;)"),
                ("\t", "Tab"),
                ("|", "Pipe (|)"),
                (" ", "Space")
            ]

            for delim_char, delim_label in delimiter_options:
                ctk.CTkRadioButton(
                    delimiter_frame,
                    text=delim_label,
                    variable=self.delimiter_var,
                    value=delim_char,
                    command=self.reload_with_delimiter
                ).pack(side=tk.LEFT, padx=5)

        # Sheet selector (if multiple sheets)
        if len(self.dataframes) > 1:
            sheet_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            sheet_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
            current_row += 1

            ctk.CTkLabel(sheet_frame, text="Sheet:", font=ctk.CTkFont(size=12, weight="bold")).pack(side=tk.LEFT, padx=(0, 10))

            self.sheet_var = tk.StringVar(value=list(self.dataframes.keys())[0])
            sheet_combo = ctk.CTkComboBox(
                sheet_frame,
                variable=self.sheet_var,
                values=list(self.dataframes.keys()),
                state="readonly",
                width=250,
                command=lambda choice: self.load_sheet()
            )
            sheet_combo.pack(side=tk.LEFT)
        else:
            self.sheet_var = tk.StringVar(value=list(self.dataframes.keys())[0])

        # Content frame (will hold stats and data grid)
        self.content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        self.content_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        current_row += 1
        self.content_frame.columnconfigure(0, weight=1)
        self.content_frame.rowconfigure(1, weight=1)

        # Bottom buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.grid(row=current_row, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

        ctk.CTkButton(button_frame, text="✓ Apply Changes", command=self.apply_changes, fg_color="#2fa572", hover_color="#26734f", width=140).pack(side=tk.RIGHT, padx=(8, 0))
        ctk.CTkButton(button_frame, text="✗ Cancel", command=self.cancel, fg_color="#666666", hover_color="#555555", width=100).pack(side=tk.RIGHT)
        ctk.CTkButton(button_frame, text="🔄 Reset to Defaults", command=self.reset_defaults, width=160).pack(side=tk.LEFT)

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
        stats_frame = ctk.CTkFrame(self.content_frame)
        stats_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        stats_frame.columnconfigure(1, weight=1)

        # Frame label
        ctk.CTkLabel(stats_frame, text="Statistics", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, columnspan=3, sticky=tk.W, padx=10, pady=(10, 5))

        # Display statistics
        row_count = len(df)
        col_count = len(df.columns)
        null_counts = df.isna().sum()
        total_nulls = null_counts.sum()

        ctk.CTkLabel(stats_frame, text=f"Rows: {row_count:,}", font=ctk.CTkFont(family="Courier", size=11, weight="bold")).grid(row=1, column=0, sticky=tk.W, padx=10, pady=(0, 10))
        ctk.CTkLabel(stats_frame, text=f"Columns: {col_count}", font=ctk.CTkFont(family="Courier", size=11, weight="bold")).grid(row=1, column=1, sticky=tk.W, padx=20, pady=(0, 10))
        ctk.CTkLabel(stats_frame, text=f"NULL values: {total_nulls:,}", font=ctk.CTkFont(family="Courier", size=11, weight="bold"), text_color="orange" if total_nulls > 0 else "green").grid(row=1, column=2, sticky=tk.W, pady=(0, 10))

        # Preview frame with scrollable area
        preview_container = ctk.CTkFrame(self.content_frame)
        preview_container.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        preview_container.columnconfigure(0, weight=1)
        preview_container.rowconfigure(1, weight=1)

        # Frame label
        ctk.CTkLabel(preview_container, text=f"Data Preview (First 20 rows) - Sheet: {sheet_name}", font=ctk.CTkFont(size=12, weight="bold")).grid(row=0, column=0, sticky=tk.W, padx=10, pady=(10, 5))

        # Create scrollable frame for data
        scrollable_frame = ctk.CTkScrollableFrame(
            preview_container,
            orientation="horizontal",
            label_text="",
            height=500  # Fixed height to ensure visibility
        )
        scrollable_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))

        # Inner container for all columns
        inner_container = ctk.CTkFrame(scrollable_frame, fg_color="transparent")
        inner_container.pack(fill=tk.BOTH, expand=True)

        # Get existing overrides for this sheet
        sheet_overrides = self.main_app.column_overrides.get(self.file_path, {}).get(sheet_name, {})
        column_name_overrides = sheet_overrides.get('columns', {})
        column_type_overrides = sheet_overrides.get('types', {})

        # Available SQL types
        sql_types = ["NVARCHAR(MAX)", "BIGINT", "FLOAT", "INT", "DECIMAL(18,2)", "DATE", "DATETIME", "BIT"]

        # Create unified grid with column headers and data
        # Display first 20 rows
        preview_df = df.head(20)

        # Handle empty dataframe
        if len(preview_df) == 0:
            ctk.CTkLabel(
                inner_container,
                text="No data rows in this sheet",
                font=ctk.CTkFont(size=12),
                text_color="orange"
            ).pack(pady=20)
            return

        # Create grid layout for columns
        for col_idx, col_name in enumerate(df.columns):
            # Create a container frame for each column (header + data)
            column_container = ctk.CTkFrame(inner_container, corner_radius=5, width=180)
            column_container.grid(row=0, column=col_idx, sticky=(tk.N, tk.S), padx=3, pady=3)
            # Allow container to expand vertically for data cells
            column_container.grid_rowconfigure(list(range(len(preview_df) + 1)), weight=0)
            column_container.grid_columnconfigure(0, weight=1)

            # Header section with edit controls
            col_frame = ctk.CTkFrame(column_container, fg_color=("#E8E8E8", "#2B2B2B"))
            col_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=2, pady=2)

            # Column name editor
            ctk.CTkLabel(col_frame, text="Column Name:", font=ctk.CTkFont(size=9)).pack(anchor=tk.W, padx=5, pady=(5, 0))
            name_var = tk.StringVar(value=column_name_overrides.get(col_name, col_name))
            name_entry = ctk.CTkEntry(col_frame, textvariable=name_var, width=160, height=28, font=ctk.CTkFont(size=10))
            name_entry.pack(fill=tk.X, padx=5, pady=(2, 5))
            self.column_name_vars[col_name] = name_var

            # Detected type display
            detected_type = infer_column_type(df[col_name], col_name)
            ctk.CTkLabel(col_frame, text=f"Detected: {detected_type}", font=ctk.CTkFont(size=9), text_color="gray").pack(anchor=tk.W, padx=5)

            # Type selector
            ctk.CTkLabel(col_frame, text="SQL Type:", font=ctk.CTkFont(size=9)).pack(anchor=tk.W, padx=5, pady=(3, 0))
            type_var = tk.StringVar(value=column_type_overrides.get(col_name, detected_type))
            type_menu = ctk.CTkOptionMenu(
                col_frame,
                variable=type_var,
                values=sql_types,
                width=160,
                height=28,
                font=ctk.CTkFont(size=9),
                dropdown_font=ctk.CTkFont(size=9)
            )
            type_menu.pack(fill=tk.X, padx=5, pady=(2, 5))
            self.column_type_vars[col_name] = type_var

            # NULL count for this column
            null_count = null_counts[col_name]
            null_pct = (null_count / row_count * 100) if row_count > 0 else 0
            null_color = "#c62828" if null_count > 0 else "#2e7d32"
            ctk.CTkLabel(col_frame, text=f"NULLs: {null_count} ({null_pct:.1f}%)", font=ctk.CTkFont(size=9), text_color=null_color).pack(anchor=tk.W, padx=5, pady=(0, 5))

            # Data cells for this column
            for row_idx, value in enumerate(preview_df[col_name]):
                # Format value
                if pd.isna(value):
                    display_value = "NULL"
                    text_color = "gray"
                else:
                    display_value = str(value)
                    if len(display_value) > 25:
                        display_value = display_value[:22] + "..."
                    text_color = None  # Use default color

                # Create label directly (simpler approach)
                cell_label = ctk.CTkLabel(
                    column_container,
                    text=display_value,
                    font=ctk.CTkFont(family="Courier", size=9),
                    anchor="w",
                    text_color=text_color,
                    fg_color=("#F5F5F5", "#3B3B3B"),
                    corner_radius=3,
                    height=24,
                    padx=5
                )
                cell_label.grid(row=row_idx+1, column=0, sticky=(tk.W, tk.E), padx=2, pady=1)

    def reload_with_delimiter(self):
        """Reload CSV file with new delimiter"""
        if not self.is_csv:
            return

        new_delimiter = self.delimiter_var.get()
        if new_delimiter == self.current_delimiter:
            return

        self.current_delimiter = new_delimiter
        self.main_app.log_message(f"Reloading {self.filename} with delimiter: '{new_delimiter}'", "INFO")

        try:
            # Reload dataframes with new delimiter
            self.dataframes = get_dataframes(self.file_path, delimiter=new_delimiter)
            self.main_app.log_message(f"Reloaded with new delimiter successfully", "SUCCESS")

            # Reload the current sheet display
            self.load_sheet()
        except Exception as e:
            self.main_app.log_message(f"Failed to reload with new delimiter: {e}", "ERROR")
            messagebox.showerror("Error", f"Failed to reload file with new delimiter:\n{e}")
            # Revert to previous delimiter
            self.delimiter_var.set(self.current_delimiter)

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

        # Save CSV delimiter preference
        if self.is_csv:
            self.main_app.csv_delimiters[self.file_path] = self.current_delimiter
            self.main_app.log_message(f"Saved delimiter preference: '{self.current_delimiter}'", "INFO")

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
