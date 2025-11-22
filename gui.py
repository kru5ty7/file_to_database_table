import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
from datetime import datetime
import os
import json

# Import main application logic
from main import (
    get_dataframes,
    get_db_connection,
    create_table_from_dataframe,
    sanitize_name,
    setup_logging,
    logger,
    get_available_connections,
    encrypt_password,
    decrypt_password,
    infer_column_type
)
import pandas as pd

class FileToDBGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("File to Database Table Converter")
        self.root.geometry("900x700")
        self.root.minsize(600, 500)  # Set minimum window size
        self.root.resizable(True, True)

        # Message queue for thread-safe GUI updates
        self.message_queue = queue.Queue()

        # Store column overrides: {file_path: {sheet_name: {'columns': {old_name: new_name}, 'types': {col_name: type}}}}
        self.column_overrides = {}

        # Configure style
        style = ttk.Style()
        style.theme_use('clam')

        # Main container with scrollbar support
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights for responsive layout
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="File to Database Table Converter",
            font=("Helvetica", 14, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10), sticky=tk.W)

        # File Selection Section
        file_frame = ttk.LabelFrame(main_frame, text="File Queue", padding="5")
        file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=2)  # Give more weight to file queue

        # File queue listbox
        queue_container = ttk.Frame(file_frame)
        queue_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))
        queue_container.columnconfigure(0, weight=1)
        queue_container.rowconfigure(0, weight=1)

        self.file_queue_listbox = tk.Listbox(
            queue_container,
            height=4,  # Reduced minimum height
            selectmode=tk.EXTENDED,
            relief=tk.SUNKEN,
            borderwidth=2,
            bg='white',
            fg='black',
            font=('Arial', 10)
        )
        self.file_queue_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=2, pady=2)

        queue_scrollbar = ttk.Scrollbar(queue_container, orient=tk.VERTICAL, command=self.file_queue_listbox.yview)
        queue_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.file_queue_listbox.config(yscrollcommand=queue_scrollbar.set)

        # File queue management buttons
        queue_btn_frame = ttk.Frame(file_frame)
        queue_btn_frame.grid(row=1, column=0, sticky=(tk.W, tk.E))

        self.add_files_button = ttk.Button(queue_btn_frame, text="Add Files", command=self.add_files)
        self.add_files_button.pack(side=tk.LEFT, padx=(0, 5))

        self.remove_files_button = ttk.Button(queue_btn_frame, text="Remove Selected", command=self.remove_selected_files)
        self.remove_files_button.pack(side=tk.LEFT, padx=(0, 5))

        self.clear_queue_button = ttk.Button(queue_btn_frame, text="Clear All", command=self.clear_file_queue)
        self.clear_queue_button.pack(side=tk.LEFT, padx=(0, 5))

        self.preview_button = ttk.Button(queue_btn_frame, text="Preview File", command=self.preview_selected_file)
        self.preview_button.pack(side=tk.LEFT)

        # File queue
        self.file_queue = []

        # Database Info Display
        db_frame = ttk.LabelFrame(main_frame, text="Database Connection", padding="5")
        db_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        db_frame.columnconfigure(1, weight=1)

        # Connection selector
        connection_selector_frame = ttk.Frame(db_frame)
        connection_selector_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 5))

        ttk.Label(connection_selector_frame, text="Connection:").pack(side=tk.LEFT, padx=(0, 10))

        self.connection_var = tk.StringVar()
        self.connection_combo = ttk.Combobox(
            connection_selector_frame,
            textvariable=self.connection_var,
            state="readonly",
            width=30
        )
        self.connection_combo.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(connection_selector_frame, text="Refresh", command=self.refresh_connections).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(connection_selector_frame, text="Manage...", command=self.manage_connections).pack(side=tk.LEFT)

        # Status label
        self.db_status_label = ttk.Label(db_frame, text="Status: Not connected", foreground="gray")
        self.db_status_label.grid(row=1, column=0, sticky=tk.W)

        ttk.Button(db_frame, text="Test Connection", command=self.test_connection).grid(
            row=1, column=1, sticky=tk.E
        )

        # Options Section
        options_frame = ttk.LabelFrame(main_frame, text="Options", padding="5")
        options_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        self.drop_existing_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            options_frame,
            text="Drop existing tables",
            variable=self.drop_existing_var
        ).grid(row=0, column=0, sticky=tk.W)

        # Progress Section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="5")
        progress_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        progress_frame.columnconfigure(0, weight=1)

        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate'
        )
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))

        self.status_label = ttk.Label(progress_frame, text="Ready", foreground="green")
        self.status_label.grid(row=1, column=0, sticky=tk.W)

        # Log Output Section
        log_frame = ttk.LabelFrame(main_frame, text="Log Output", padding="5")
        log_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 5))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(5, weight=3)  # Give more weight to log section

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=8,  # Reduced minimum height
            wrap=tk.WORD,
            font=("Courier", 9),
            state='disabled'
        )
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Action Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        button_frame.columnconfigure(0, weight=1)

        self.convert_button = ttk.Button(
            button_frame,
            text="Convert to Database",
            command=self.start_conversion,
            style="Accent.TButton"
        )
        self.convert_button.grid(row=0, column=0, sticky=tk.E, padx=(0, 10))

        self.clear_log_button = ttk.Button(
            button_frame,
            text="Clear Log",
            command=self.clear_log
        )
        self.clear_log_button.grid(row=0, column=1, sticky=tk.E)

        # Configure button style
        style.configure("Accent.TButton", font=("Helvetica", 10, "bold"))

        # Start message queue processor
        self.process_queue()

        # Load available connections after GUI is fully initialized
        self.refresh_connections()

        # Add initial log message
        self.log_message("Application started. Please select a file to begin.")

    def refresh_connections(self):
        """Refresh the list of available connections"""
        connections = get_available_connections()
        self.connection_combo['values'] = connections

        if connections:
            if not self.connection_var.get() or self.connection_var.get() not in connections:
                self.connection_var.set(connections[0])
            self.log_message(f"Loaded {len(connections)} connection(s): {', '.join(connections)}")
        else:
            self.log_message("No connections found. Please check config.json", "ERROR")

    def manage_connections(self):
        """Open connection management dialog"""
        self.log_message("Opening connection management dialog", "INFO")
        ConnectionManagerDialog(self.root, self)

    def add_files(self):
        """Open file browser dialog to add multiple files"""
        filenames = filedialog.askopenfilenames(
            title="Select files",
            filetypes=[
                ("Supported Files", "*.csv *.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("Excel files", "*.xlsx *.xls"),
                ("All files", "*.*")
            ]
        )

        # Debug: Log the filenames returned
        self.log_message(f"Selected {len(filenames)} file(s) from dialog", "INFO")

        if filenames:
            added_count = 0
            for filename in filenames:
                if filename not in self.file_queue:
                    self.file_queue.append(filename)
                    basename = os.path.basename(filename)
                    self.file_queue_listbox.insert(tk.END, basename)
                    self.log_message(f"Added to listbox: {basename}", "INFO")
                    added_count += 1
                else:
                    self.log_message(f"Skipped duplicate: {os.path.basename(filename)}", "INFO")

            if added_count > 0:
                self.log_message(f"Added {added_count} file(s) to queue. Total: {len(self.file_queue)}")
                # Force listbox update
                self.file_queue_listbox.update()
            else:
                self.log_message("No new files added (duplicates skipped)", "INFO")
        else:
            self.log_message("No files selected", "INFO")

    def remove_selected_files(self):
        """Remove selected files from queue"""
        selected_indices = self.file_queue_listbox.curselection()
        if not selected_indices:
            self.log_message("No files selected for removal", "INFO")
            return

        # Remove in reverse order to maintain correct indices
        for index in reversed(selected_indices):
            filename = self.file_queue[index]
            self.file_queue.pop(index)
            self.file_queue_listbox.delete(index)
            self.log_message(f"Removed: {os.path.basename(filename)}")

        self.log_message(f"Files remaining in queue: {len(self.file_queue)}")

    def clear_file_queue(self):
        """Clear all files from queue"""
        if len(self.file_queue) == 0:
            self.log_message("Queue is already empty", "INFO")
            return

        count = len(self.file_queue)
        self.file_queue.clear()
        self.file_queue_listbox.delete(0, tk.END)
        self.column_overrides.clear()
        self.log_message(f"Cleared {count} file(s) from queue")

    def preview_selected_file(self):
        """Open preview dialog for selected file"""
        selected_indices = self.file_queue_listbox.curselection()
        if not selected_indices:
            self.log_message("No file selected for preview", "INFO")
            messagebox.showinfo("No Selection", "Please select a file to preview")
            return

        file_index = selected_indices[0]
        file_path = self.file_queue[file_index]

        if not os.path.exists(file_path):
            self.log_message(f"File not found: {file_path}", "ERROR")
            messagebox.showerror("File Not Found", f"File not found: {os.path.basename(file_path)}")
            return

        self.log_message(f"Opening preview for: {os.path.basename(file_path)}")
        DataPreviewDialog(self.root, self, file_path)

    def log_message(self, message, level="INFO"):
        """Add message to log text widget"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {level}: {message}\n"

        # Temporarily enable editing to insert text
        self.log_text.config(state='normal')

        # Color coding based on level
        self.log_text.insert(tk.END, formatted_message)

        if level == "ERROR":
            # Find the line we just inserted and tag it
            line_start = self.log_text.index("end-2c linestart")
            line_end = self.log_text.index("end-1c")
            self.log_text.tag_add("error", line_start, line_end)
            self.log_text.tag_config("error", foreground="red")
        elif level == "SUCCESS":
            line_start = self.log_text.index("end-2c linestart")
            line_end = self.log_text.index("end-1c")
            self.log_text.tag_add("success", line_start, line_end)
            self.log_text.tag_config("success", foreground="green", font=("Courier", 9, "bold"))

        self.log_text.see(tk.END)

        # Disable editing again
        self.log_text.config(state='disabled')

    def clear_log(self):
        """Clear the log text widget"""
        self.log_text.config(state='normal')
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state='disabled')
        self.log_message("Log cleared")

    def update_status(self, message, color="black"):
        """Update status label"""
        self.status_label.config(text=message, foreground=color)

    def update_progress(self, value):
        """Update progress bar"""
        self.progress_var.set(value)

    def test_connection(self):
        """Test database connection"""
        connection_name = self.connection_var.get()
        if not connection_name:
            self.log_message("Please select a connection", "ERROR")
            return

        self.log_message(f"Testing connection '{connection_name}'...")
        self.update_status("Testing connection...", "orange")

        def test():
            try:
                conn = get_db_connection(connection_name)
                conn.close()
                self.message_queue.put(("log", f"Database connection '{connection_name}' successful!", "SUCCESS"))
                self.message_queue.put(("status", "Connected", "green"))
                self.message_queue.put(("db_status", "Status: Connected ✓", "green"))
            except Exception as e:
                self.message_queue.put(("log", f"Connection failed: {e}", "ERROR"))
                self.message_queue.put(("status", "Connection failed", "red"))
                self.message_queue.put(("db_status", "Status: Connection failed ✗", "red"))

        threading.Thread(target=test, daemon=True).start()

    def start_conversion(self):
        """Start the batch conversion process in a separate thread"""
        connection_name = self.connection_var.get()

        if len(self.file_queue) == 0:
            messagebox.showwarning("No Files in Queue", "Please add files to the queue before converting.")
            return

        if not connection_name:
            messagebox.showwarning("No Connection Selected", "Please select a database connection.")
            return

        # Validate all files exist
        invalid_files = [f for f in self.file_queue if not os.path.exists(f)]
        if invalid_files:
            messagebox.showerror(
                "Files Not Found",
                f"The following files were not found:\n" + "\n".join([os.path.basename(f) for f in invalid_files])
            )
            return

        # Disable buttons during conversion
        self.convert_button.config(state="disabled")
        self.add_files_button.config(state="disabled")
        self.remove_files_button.config(state="disabled")
        self.clear_queue_button.config(state="disabled")

        # Reset progress
        self.update_progress(0)
        self.update_status("Converting...", "blue")
        self.log_message(f"Starting batch conversion of {len(self.file_queue)} file(s) using connection '{connection_name}'...")

        # Start batch conversion in background thread
        threading.Thread(
            target=self.convert_batch,
            args=(self.file_queue.copy(), connection_name),
            daemon=True
        ).start()

    def convert_batch(self, file_list, connection_name):
        """Convert multiple files to database tables (runs in background thread)"""
        total_files = len(file_list)
        successful_files = 0
        failed_files = []

        try:
            # Connect to database once for all files
            self.message_queue.put(("log", f"Connecting to database using '{connection_name}'...", "INFO"))
            conn = get_db_connection(connection_name)
            cursor = conn.cursor()

            for file_index, file_path in enumerate(file_list, 1):
                try:
                    filename = os.path.basename(file_path)
                    self.message_queue.put(("log", f"\n[{file_index}/{total_files}] Processing: {filename}", "INFO"))

                    # Calculate progress for this file (each file gets equal portion)
                    file_progress_start = int(((file_index - 1) / total_files) * 100)
                    file_progress_range = int(100 / total_files)

                    # Read file
                    self.message_queue.put(("progress", file_progress_start + int(file_progress_range * 0.1)))
                    dataframes = get_dataframes(file_path)
                    self.message_queue.put(("log", f"  Found {len(dataframes)} sheet(s)", "INFO"))

                    # Process each sheet
                    base_table_name = sanitize_name(os.path.splitext(filename)[0])
                    total_sheets = len(dataframes)

                    for idx, (sheet_name, df) in enumerate(dataframes.items()):
                        if len(dataframes) == 1:
                            table_name = base_table_name
                        else:
                            table_name = f"{base_table_name}_{sheet_name}"

                        # Get column overrides for this file and sheet
                        sheet_overrides = self.column_overrides.get(file_path, {}).get(sheet_name, {})
                        column_name_map = sheet_overrides.get('columns', {})
                        column_type_map = sheet_overrides.get('types', {})

                        if column_name_map:
                            self.message_queue.put(("log", f"  Applying {len(column_name_map)} column name override(s)", "INFO"))
                        if column_type_map:
                            self.message_queue.put(("log", f"  Applying {len(column_type_map)} column type override(s)", "INFO"))

                        self.message_queue.put(("log", f"  Creating table: {table_name}", "INFO"))
                        create_table_from_dataframe(df, table_name, cursor, column_name_map, column_type_map)

                        # Update progress within this file
                        sheet_progress = int(file_progress_range * (0.2 + 0.7 * (idx + 1) / total_sheets))
                        self.message_queue.put(("progress", file_progress_start + sheet_progress))

                    self.message_queue.put(("log", f"  ✓ {filename} completed successfully", "SUCCESS"))
                    successful_files += 1

                except Exception as e:
                    self.message_queue.put(("log", f"  ✗ Failed to process {filename}: {e}", "ERROR"))
                    failed_files.append((filename, str(e)))
                    # Continue with next file

            cursor.close()
            conn.close()

            # Final summary
            self.message_queue.put(("progress", 100))
            self.message_queue.put(("log", f"\n{'='*60}", "INFO"))
            self.message_queue.put(("log", f"Batch conversion completed!", "SUCCESS"))
            self.message_queue.put(("log", f"  Total files: {total_files}", "INFO"))
            self.message_queue.put(("log", f"  Successful: {successful_files}", "SUCCESS"))
            if failed_files:
                self.message_queue.put(("log", f"  Failed: {len(failed_files)}", "ERROR"))
                for filename, error in failed_files:
                    self.message_queue.put(("log", f"    - {filename}: {error}", "ERROR"))
            self.message_queue.put(("log", f"{'='*60}", "INFO"))

            self.message_queue.put(("status", f"Completed: {successful_files}/{total_files} files", "green"))
            self.message_queue.put(("enable_buttons", None))

            if failed_files:
                error_summary = f"Completed with {len(failed_files)} error(s).\n\n" + \
                               "\n".join([f"- {f[0]}" for f in failed_files[:5]])
                if len(failed_files) > 5:
                    error_summary += f"\n... and {len(failed_files) - 5} more"
                self.message_queue.put(("show_error", error_summary))
            else:
                self.message_queue.put(("show_success", f"Successfully converted all {successful_files} file(s)!"))

        except Exception as e:
            self.message_queue.put(("log", f"Batch conversion error: {e}", "ERROR"))
            self.message_queue.put(("status", "Batch conversion failed", "red"))
            self.message_queue.put(("progress", 0))
            self.message_queue.put(("enable_buttons", None))
            self.message_queue.put(("show_error", f"Batch conversion failed: {str(e)}"))

    def convert_file(self, file_path, connection_name):
        """Convert file to database tables (runs in background thread) - Legacy single file method"""
        try:
            # Read file
            self.message_queue.put(("log", f"Reading file: {file_path}", "INFO"))
            self.message_queue.put(("progress", 10))

            dataframes = get_dataframes(file_path)
            self.message_queue.put(("log", f"Found {len(dataframes)} sheet(s)", "INFO"))
            self.message_queue.put(("progress", 20))

            # Connect to database
            self.message_queue.put(("log", f"Connecting to database using '{connection_name}'...", "INFO"))
            conn = get_db_connection(connection_name)
            cursor = conn.cursor()
            self.message_queue.put(("progress", 30))

            # Process each sheet
            base_table_name = sanitize_name(os.path.splitext(os.path.basename(file_path))[0])
            total_sheets = len(dataframes)
            progress_per_sheet = 60 / total_sheets

            for idx, (sheet_name, df) in enumerate(dataframes.items()):
                if len(dataframes) == 1:
                    table_name = base_table_name
                else:
                    table_name = f"{base_table_name}_{sheet_name}"

                # Get column overrides for this file and sheet
                sheet_overrides = self.column_overrides.get(file_path, {}).get(sheet_name, {})
                column_name_map = sheet_overrides.get('columns', {})
                column_type_map = sheet_overrides.get('types', {})

                self.message_queue.put(("log", f"Processing sheet: {sheet_name} → table: {table_name}", "INFO"))

                create_table_from_dataframe(df, table_name, cursor, column_name_map, column_type_map)

                self.message_queue.put(("log", f"✓ Table '{table_name}' created successfully", "SUCCESS"))
                current_progress = int(30 + ((idx + 1) * progress_per_sheet))
                self.message_queue.put(("progress", current_progress))

            cursor.close()
            conn.close()

            self.message_queue.put(("progress", 100))
            self.message_queue.put(("log", f"✓ All {total_sheets} table(s) created successfully!", "SUCCESS"))
            self.message_queue.put(("status", "Conversion completed!", "green"))
            self.message_queue.put(("enable_buttons", None))
            self.message_queue.put(("show_success", f"Successfully created {total_sheets} table(s)!"))

        except Exception as e:
            self.message_queue.put(("log", f"Error: {e}", "ERROR"))
            self.message_queue.put(("status", "Conversion failed", "red"))
            self.message_queue.put(("progress", 0))
            self.message_queue.put(("enable_buttons", None))
            self.message_queue.put(("show_error", str(e)))

    def process_queue(self):
        """Process messages from background thread"""
        try:
            while True:
                msg = self.message_queue.get_nowait()

                # Handle both 2-tuple and 3-tuple messages
                if len(msg) == 2:
                    msg_type, msg_data = msg
                elif len(msg) == 3:
                    msg_type, msg_data, msg_extra = msg
                else:
                    continue

                if msg_type == "log":
                    if len(msg) == 3:
                        # 3-tuple: ("log", message, level)
                        self.log_message(msg_data, msg_extra)
                    else:
                        self.log_message(msg_data)

                elif msg_type == "status":
                    if len(msg) == 3:
                        # 3-tuple: ("status", message, color)
                        self.update_status(msg_data, msg_extra)
                    else:
                        self.update_status(msg_data)

                elif msg_type == "progress":
                    self.update_progress(msg_data)

                elif msg_type == "db_status":
                    if len(msg) == 3:
                        # 3-tuple: ("db_status", text, color)
                        self.db_status_label.config(text=msg_data, foreground=msg_extra)
                    else:
                        self.db_status_label.config(text=msg_data)

                elif msg_type == "enable_buttons":
                    self.convert_button.config(state="normal")
                    self.add_files_button.config(state="normal")
                    self.remove_files_button.config(state="normal")
                    self.clear_queue_button.config(state="normal")

                elif msg_type == "show_success":
                    messagebox.showinfo("Success", msg_data)

                elif msg_type == "show_error":
                    messagebox.showerror("Error", msg_data)

        except queue.Empty:
            pass

        # Schedule next check
        self.root.after(100, self.process_queue)


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


class ConnectionManagerDialog:
    def __init__(self, parent, main_app):
        self.main_app = main_app
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Manage Database Connections")
        self.dialog.geometry("700x500")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        # Load config
        self.load_config()

        # Main frame
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Connection list
        list_frame = ttk.LabelFrame(main_frame, text="Connections", padding="10")
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.conn_listbox = tk.Listbox(list_frame, height=15)
        self.conn_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.conn_listbox.bind('<<ListboxSelect>>', self.on_connection_select)

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.conn_listbox.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.conn_listbox.config(yscrollcommand=scrollbar.set)

        # Buttons for list
        list_btn_frame = ttk.Frame(list_frame)
        list_btn_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))

        ttk.Button(list_btn_frame, text="Add", command=self.add_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(list_btn_frame, text="Delete", command=self.delete_connection).pack(side=tk.LEFT, padx=5)

        # Connection details
        details_frame = ttk.LabelFrame(main_frame, text="Connection Details", padding="10")
        details_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        details_frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(details_frame, text="Connection Name:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(details_frame, textvariable=self.name_var, width=30)
        self.name_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)

        row += 1
        ttk.Label(details_frame, text="Server:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.server_var = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.server_var, width=30).grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)

        row += 1
        ttk.Label(details_frame, text="Database:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.database_var = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.database_var, width=30).grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)

        row += 1
        ttk.Label(details_frame, text="Username:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.username_var = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.username_var, width=30).grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)

        row += 1
        ttk.Label(details_frame, text="Password:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        ttk.Entry(details_frame, textvariable=self.password_var, width=30, show="*").grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)

        row += 1
        ttk.Label(details_frame, text="Driver:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.driver_var = tk.StringVar(value="{ODBC Driver 17 for SQL Server}")
        ttk.Entry(details_frame, textvariable=self.driver_var, width=30).grid(row=row, column=1, sticky=(tk.W, tk.E), pady=5)

        row += 1
        details_btn_frame = ttk.Frame(details_frame)
        details_btn_frame.grid(row=row, column=0, columnspan=2, pady=(20, 0))

        ttk.Button(details_btn_frame, text="Save", command=self.save_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(details_btn_frame, text="Test", command=self.test_current_connection).pack(side=tk.LEFT, padx=5)

        # Bottom buttons
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))

        ttk.Button(bottom_frame, text="Close", command=self.close_dialog).pack(side=tk.RIGHT, padx=5)

        # Load connections into list
        self.refresh_list()

    def load_config(self):
        """Load config from file"""
        try:
            with open('config.json', 'r') as f:
                self.config = json.load(f)
                # Ensure new format
                if 'connections' not in self.config:
                    # Convert old format to new
                    old_config = self.config.copy()
                    self.config = {
                        'default_connection': 'default',
                        'connections': {
                            'default': old_config
                        }
                    }
                    self.main_app.log_message("Converted legacy config format to new format", "INFO")
                self.main_app.log_message(f"Loaded config with {len(self.config.get('connections', {}))} connection(s)", "INFO")
        except FileNotFoundError:
            self.config = {
                'default_connection': 'production',
                'connections': {}
            }
            self.main_app.log_message("No config.json found, starting with empty configuration", "INFO")

    def save_config(self):
        """Save config to file"""
        try:
            with open('config.json', 'w') as f:
                json.dump(self.config, f, indent=4)
            self.main_app.log_message("Configuration saved successfully to config.json", "INFO")
            return True
        except Exception as e:
            error_msg = f"Failed to save config: {e}"
            self.main_app.log_message(error_msg, "ERROR")
            messagebox.showerror("Error", error_msg)
            return False

    def refresh_list(self):
        """Refresh the connection list"""
        self.conn_listbox.delete(0, tk.END)
        for conn_name in self.config.get('connections', {}).keys():
            self.conn_listbox.insert(tk.END, conn_name)

    def on_connection_select(self, event):
        """Handle connection selection"""
        selection = self.conn_listbox.curselection()
        if selection:
            conn_name = self.conn_listbox.get(selection[0])
            conn_data = self.config['connections'][conn_name]

            self.name_var.set(conn_name)
            self.name_entry.config(state='readonly')
            self.server_var.set(conn_data.get('server', ''))
            self.database_var.set(conn_data.get('database', ''))
            self.username_var.set(conn_data.get('username', ''))
            # Decrypt password for display
            decrypted_password = decrypt_password(conn_data.get('password', ''))
            self.password_var.set(decrypted_password)
            self.driver_var.set(conn_data.get('driver', '{ODBC Driver 17 for SQL Server}'))

            self.main_app.log_message(f"Selected connection: '{conn_name}' (Server: {conn_data.get('server', 'N/A')}, Database: {conn_data.get('database', 'N/A')})", "INFO")

    def add_connection(self):
        """Add new connection"""
        self.name_entry.config(state='normal')
        self.name_var.set('')
        self.server_var.set('')
        self.database_var.set('')
        self.username_var.set('')
        self.password_var.set('')
        self.driver_var.set('{ODBC Driver 17 for SQL Server}')
        self.conn_listbox.selection_clear(0, tk.END)
        self.main_app.log_message("New connection form opened", "INFO")

    def save_connection(self):
        """Save current connection"""
        conn_name = self.name_var.get().strip()
        if not conn_name:
            self.main_app.log_message("Save connection failed: Connection name is required", "ERROR")
            messagebox.showwarning("Warning", "Please enter a connection name")
            return

        if not self.server_var.get().strip():
            self.main_app.log_message("Save connection failed: Server name is required", "ERROR")
            messagebox.showwarning("Warning", "Please enter a server name")
            return

        is_new = conn_name not in self.config.get('connections', {})

        # Encrypt password before saving
        encrypted_password = encrypt_password(self.password_var.get())

        conn_data = {
            'server': self.server_var.get().strip(),
            'database': self.database_var.get().strip(),
            'username': self.username_var.get().strip(),
            'password': encrypted_password,
            'driver': self.driver_var.get().strip()
        }

        self.main_app.log_message(f"{'Creating' if is_new else 'Updating'} connection '{conn_name}' (Server: {conn_data['server']}, Database: {conn_data['database']}) with encrypted password", "INFO")

        self.config['connections'][conn_name] = conn_data

        # Set as default if it's the first connection
        if len(self.config['connections']) == 1:
            self.config['default_connection'] = conn_name
            self.main_app.log_message(f"Set '{conn_name}' as default connection", "INFO")

        if self.save_config():
            self.main_app.log_message(f"Connection '{conn_name}' {'created' if is_new else 'updated'} successfully", "SUCCESS")
            messagebox.showinfo("Success", f"Connection '{conn_name}' saved successfully")
            self.refresh_list()
            self.main_app.refresh_connections()
            self.name_entry.config(state='readonly')

    def delete_connection(self):
        """Delete selected connection"""
        selection = self.conn_listbox.curselection()
        if not selection:
            self.main_app.log_message("Delete failed: No connection selected", "ERROR")
            messagebox.showwarning("Warning", "Please select a connection to delete")
            return

        conn_name = self.conn_listbox.get(selection[0])

        self.main_app.log_message(f"Delete confirmation requested for connection '{conn_name}'", "INFO")

        if messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete connection '{conn_name}'?"):
            self.main_app.log_message(f"Deleting connection '{conn_name}'...", "INFO")
            del self.config['connections'][conn_name]

            # Update default if needed
            if self.config.get('default_connection') == conn_name:
                remaining = list(self.config['connections'].keys())
                new_default = remaining[0] if remaining else ''
                self.config['default_connection'] = new_default
                if new_default:
                    self.main_app.log_message(f"Updated default connection to '{new_default}'", "INFO")

            if self.save_config():
                self.main_app.log_message(f"Connection '{conn_name}' deleted successfully", "SUCCESS")
                messagebox.showinfo("Success", f"Connection '{conn_name}' deleted successfully")
                self.refresh_list()
                self.main_app.refresh_connections()
                self.add_connection()  # Clear form
        else:
            self.main_app.log_message(f"Delete cancelled for connection '{conn_name}'", "INFO")

    def test_current_connection(self):
        """Test the current connection settings"""
        # Validate required fields
        if not self.server_var.get().strip():
            self.main_app.log_message("Connection test failed: Server name is required", "ERROR")
            messagebox.showwarning("Warning", "Please enter a server name")
            return

        if not self.database_var.get().strip():
            self.main_app.log_message("Connection test failed: Database name is required", "ERROR")
            messagebox.showwarning("Warning", "Please enter a database name")
            return

        conn_name = self.name_var.get().strip() or "Unnamed"
        self.main_app.log_message(f"Testing connection '{conn_name}' (Server: {self.server_var.get()}, Database: {self.database_var.get()})...", "INFO")

        # Show testing message
        test_window = tk.Toplevel(self.dialog)
        test_window.title("Testing Connection")
        test_window.geometry("300x100")
        test_window.transient(self.dialog)

        label = ttk.Label(test_window, text="Testing connection...", padding=20)
        label.pack()

        progress = ttk.Progressbar(test_window, mode='indeterminate', length=250)
        progress.pack(pady=10)
        progress.start()

        def test_in_thread():
            try:
                import pyodbc
                # Password in the form is already decrypted, use it directly
                conn_str = (
                    f'DRIVER={self.driver_var.get()};'
                    f'SERVER={self.server_var.get()};'
                    f'DATABASE={self.database_var.get()};'
                    f'UID={self.username_var.get()};'
                    f'PWD={self.password_var.get()};'
                )
                conn = pyodbc.connect(conn_str, timeout=10)
                conn.close()

                # Close test window and show success
                test_window.destroy()
                self.main_app.log_message(f"Connection test successful for '{conn_name}'", "SUCCESS")
                messagebox.showinfo("Success", "Connection test successful!")
            except Exception as e:
                # Close test window and show error
                test_window.destroy()
                error_msg = str(e)
                # Make error message more readable
                if "ODBC Driver" in error_msg:
                    error_msg = "ODBC Driver not found. Please install the specified driver."
                    self.main_app.log_message(f"Connection test failed for '{conn_name}': ODBC Driver not found", "ERROR")
                elif "Login failed" in error_msg or "Login timeout" in error_msg:
                    error_msg = f"Authentication failed. Please check username and password.\n\nDetails: {error_msg}"
                    self.main_app.log_message(f"Connection test failed for '{conn_name}': Authentication failed", "ERROR")
                elif "Could not open a connection" in error_msg:
                    error_msg = f"Could not connect to server. Please check server name.\n\nDetails: {error_msg}"
                    self.main_app.log_message(f"Connection test failed for '{conn_name}': Could not connect to server", "ERROR")
                else:
                    error_msg = f"Connection failed:\n\n{error_msg}"
                    self.main_app.log_message(f"Connection test failed for '{conn_name}': {str(e)}", "ERROR")

                messagebox.showerror("Connection Failed", error_msg)

        # Start test in background thread
        threading.Thread(target=test_in_thread, daemon=True).start()

    def close_dialog(self):
        """Close the dialog"""
        self.main_app.log_message("Connection management dialog closed", "INFO")
        self.dialog.destroy()


def main():
    root = tk.Tk()
    app = FileToDBGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
