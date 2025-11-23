"""
Main GUI Window - File to Database Converter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import queue
from datetime import datetime
import os

from src.database import get_db_connection, create_table_from_dataframe, get_available_connections
from src.file_processor import get_dataframes
from src.utils import sanitize_name, setup_logging, logger
from src.dialogs import DataPreviewDialog, ConnectionManagerDialog


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

