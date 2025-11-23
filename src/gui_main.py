"""
Main GUI Window - File to Database Converter
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import threading
import queue
from datetime import datetime
import os
from PIL import ImageGrab
import subprocess

from src.database import get_db_connection, create_table_from_dataframe, get_available_connections
from src.file_processor import get_dataframes
from src.utils import sanitize_name, setup_logging, logger
from src.dialogs import DataPreviewDialog, ConnectionManagerDialog

# Get version from git tag or use default
def get_version():
    """Get version from git tag"""
    try:
        result = subprocess.run(['git', 'describe', '--tags', '--abbrev=0'],
                              capture_output=True, text=True, timeout=2)
        if result.returncode == 0:
            return result.stdout.strip()
    except:
        pass
    return "v1.0.0"  # Default version

VERSION = get_version()


class FileToDBGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("File to Database Table Converter")
        self.root.geometry("950x750")  # Increased window size
        self.root.minsize(800, 600)  # Increased minimum window size

        # Message queue for thread-safe GUI updates
        self.message_queue = queue.Queue()

        # Store column overrides: {file_path: {sheet_name: {'columns': {old_name: new_name}, 'types': {col_name: type}}}}
        self.column_overrides = {}

        # Store CSV delimiter preferences: {file_path: delimiter}
        self.csv_delimiters = {}

        # Main container with scrollbar support
        main_frame = ctk.CTkFrame(root)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)

        # Configure grid weights for responsive layout
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Title row
        title_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        title_frame.columnconfigure(0, weight=1)

        title_label = ctk.CTkLabel(
            title_frame,
            text=f"File to Database Table Converter {VERSION}",
            font=ctk.CTkFont(size=20, weight="bold")
        )
        title_label.pack(side=tk.LEFT)

        # File Selection Section
        file_frame = ctk.CTkFrame(main_frame)
        file_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        file_frame.columnconfigure(0, weight=1)
        file_frame.rowconfigure(1, weight=1)

        # Frame label
        ctk.CTkLabel(file_frame, text="File Queue", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky=tk.W, padx=10, pady=(10, 5))

        # File queue textbox - shows 3-4 files, scrolls for more
        self.file_queue_textbox = ctk.CTkTextbox(
            file_frame,
            height=100,  # Height for 3-4 files, scrollbar appears automatically for more
            font=ctk.CTkFont(size=11),
            wrap="none"  # Prevent text wrapping for long filenames
        )
        self.file_queue_textbox.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=10, pady=(0, 10))

        # Store file queue as list (we'll update textbox display)
        self.file_queue_listbox = None  # Keep for compatibility, will handle differently

        # File queue management buttons
        queue_btn_frame = ctk.CTkFrame(file_frame, fg_color="transparent")
        queue_btn_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), padx=10, pady=(0, 10))

        self.add_files_button = ctk.CTkButton(queue_btn_frame, text="➕ Add Files", command=self.add_files, width=120)
        self.add_files_button.pack(side=tk.LEFT, padx=(0, 8))

        self.remove_files_button = ctk.CTkButton(queue_btn_frame, text="➖ Remove Selected", command=self.remove_selected_files, width=140)
        self.remove_files_button.pack(side=tk.LEFT, padx=(0, 8))

        self.clear_queue_button = ctk.CTkButton(queue_btn_frame, text="🗑 Clear All", command=self.clear_file_queue, width=100)
        self.clear_queue_button.pack(side=tk.LEFT, padx=(0, 8))

        self.preview_button = ctk.CTkButton(queue_btn_frame, text="👁 Preview File", command=self.preview_selected_file, width=120, fg_color="#2fa572")
        self.preview_button.pack(side=tk.LEFT)

        # File queue
        self.file_queue = []
        self.file_queue_selection = None

        # Make textbox clickable to select files
        self.file_queue_textbox.bind("<Button-1>", self._on_file_queue_click)

        # Initialize the display
        self._update_file_queue_display()

        # Database Info Display
        db_frame = ctk.CTkFrame(main_frame)
        db_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        db_frame.columnconfigure(1, weight=1)

        # Frame label
        ctk.CTkLabel(db_frame, text="Database Connection", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(10, 5))

        # Connection selector
        connection_selector_frame = ctk.CTkFrame(db_frame, fg_color="transparent")
        connection_selector_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), padx=10, pady=(0, 10))

        ctk.CTkLabel(connection_selector_frame, text="Connection:").pack(side=tk.LEFT, padx=(0, 10))

        self.connection_var = tk.StringVar()
        self.connection_combo = ctk.CTkComboBox(
            connection_selector_frame,
            variable=self.connection_var,
            state="readonly",
            width=250
        )
        self.connection_combo.pack(side=tk.LEFT, padx=(0, 10))

        ctk.CTkButton(connection_selector_frame, text="🔄 Refresh", command=self.refresh_connections, width=90).pack(side=tk.LEFT, padx=(0, 8))
        ctk.CTkButton(connection_selector_frame, text="⚙ Manage", command=self.manage_connections, width=90).pack(side=tk.LEFT)

        # Status label
        self.db_status_label = ctk.CTkLabel(db_frame, text="Status: Not connected", text_color="gray")
        self.db_status_label.grid(row=2, column=0, sticky=tk.W, padx=10, pady=(0, 10))

        ctk.CTkButton(db_frame, text="✓ Test Connection", command=self.test_connection, width=140).grid(
            row=2, column=1, sticky=tk.E, padx=10, pady=(0, 10)
        )

        # Options Section
        options_frame = ctk.CTkFrame(main_frame)
        options_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 10))

        ctk.CTkLabel(options_frame, text="Options", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky=tk.W, padx=10, pady=(10, 5))

        self.drop_existing_var = tk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            options_frame,
            text="Drop existing tables if they exist",
            variable=self.drop_existing_var,
            font=ctk.CTkFont(size=12)
        ).grid(row=1, column=0, sticky=tk.W, padx=10, pady=(0, 10))

        # Progress Section
        progress_frame = ctk.CTkFrame(main_frame)
        progress_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)

        # Progress header with percentage
        progress_header_frame = ctk.CTkFrame(progress_frame, fg_color="transparent")
        progress_header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=10, pady=(10, 5))
        progress_header_frame.columnconfigure(0, weight=1)

        ctk.CTkLabel(progress_header_frame, text="Progress", font=ctk.CTkFont(size=13, weight="bold")).pack(side=tk.LEFT)

        self.progress_percentage_label = ctk.CTkLabel(
            progress_header_frame,
            text="0%",
            font=ctk.CTkFont(size=12, weight="bold"),
            text_color="#1565c0"
        )
        self.progress_percentage_label.pack(side=tk.RIGHT)

        # Progress bar
        self.progress_bar = ctk.CTkProgressBar(
            progress_frame,
            mode='determinate',
            height=20
        )
        self.progress_bar.grid(row=1, column=0, sticky=(tk.W, tk.E), padx=10, pady=(0, 8))
        self.progress_bar.set(0)
        self.current_progress = 0

        # Status label
        self.status_label = ctk.CTkLabel(progress_frame, text="Ready", text_color="#2e7d32", font=ctk.CTkFont(size=12, weight="bold"))
        self.status_label.grid(row=2, column=0, sticky=tk.W, padx=10, pady=(0, 10))

        # Log Output Section
        log_frame = ctk.CTkFrame(main_frame)
        log_frame.grid(row=5, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=3)  # Give more weight to log section

        ctk.CTkLabel(log_frame, text="Log Output", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, sticky=tk.W, padx=10, pady=(10, 5))

        self.log_text = ctk.CTkTextbox(
            log_frame,
            height=180,  # Fixed height for consistent display
            wrap="word",
            font=ctk.CTkFont(family="Courier", size=11)
        )
        self.log_text.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))
        self.log_text.configure(state='disabled')

        # Action Buttons
        button_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        button_frame.grid(row=6, column=0, sticky=(tk.W, tk.E), pady=(0, 0))
        button_frame.columnconfigure(0, weight=1)

        self.convert_button = ctk.CTkButton(
            button_frame,
            text="▶ Convert to Database",
            command=self.start_conversion,
            height=40,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#1f6aa5",
            hover_color="#144870"
        )
        self.convert_button.grid(row=0, column=0, sticky=tk.E, padx=(0, 10))

        self.clear_log_button = ctk.CTkButton(
            button_frame,
            text="🗑 Clear Log",
            command=self.clear_log,
            height=40,
            width=120,
            fg_color="#666666",
            hover_color="#555555"
        )
        self.clear_log_button.grid(row=0, column=1, sticky=tk.E)

        # Start message queue processor
        self.process_queue()

        # Load available connections after GUI is fully initialized
        self.refresh_connections()

        # Add initial log message
        self.log_message("Application started. Please select a file to begin.")

    def _update_file_queue_display(self):
        """Update the file queue textbox display"""
        self.file_queue_textbox.configure(state='normal')
        self.file_queue_textbox.delete("1.0", tk.END)

        if len(self.file_queue) == 0:
            self.file_queue_textbox.insert(tk.END, "No files in queue. Click 'Add Files' to begin.")
        else:
            for i, filepath in enumerate(self.file_queue):
                basename = os.path.basename(filepath)
                prefix = "▶ " if i == self.file_queue_selection else "   "
                # Show full filename without truncation
                self.file_queue_textbox.insert(tk.END, f"{prefix}{i+1}. {basename}\n")

        self.file_queue_textbox.configure(state='disabled')

    def _on_file_queue_click(self, event):
        """Handle click on file queue textbox"""
        try:
            # Get the line number that was clicked
            index = self.file_queue_textbox.index(f"@{event.x},{event.y}")
            line_num = int(index.split('.')[0]) - 1
            if 0 <= line_num < len(self.file_queue):
                self.file_queue_selection = line_num
                self._update_file_queue_display()
        except:
            pass

    def take_screenshot(self):
        """Take a screenshot of the application window"""
        try:
            # Get window position and size
            x = self.root.winfo_rootx()
            y = self.root.winfo_rooty()
            width = self.root.winfo_width()
            height = self.root.winfo_height()

            # Capture the window
            screenshot = ImageGrab.grab(bbox=(x, y, x + width, y + height))

            # Create screenshots directory if it doesn't exist
            screenshot_dir = "screenshots"
            if not os.path.exists(screenshot_dir):
                os.makedirs(screenshot_dir)

            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join(screenshot_dir, f"screenshot_{timestamp}.png")

            # Save screenshot
            screenshot.save(filename)

            self.log_message(f"Screenshot saved: {filename}", "SUCCESS")
            messagebox.showinfo("Screenshot Saved", f"Screenshot saved to:\n{filename}")
        except Exception as e:
            error_msg = f"Failed to take screenshot: {e}"
            self.log_message(error_msg, "ERROR")
            messagebox.showerror("Screenshot Error", error_msg)

    def refresh_connections(self):
        """Refresh the list of available connections"""
        connections = get_available_connections()
        self.connection_combo.configure(values=connections)

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
                    self.log_message(f"Added: {basename}", "INFO")
                    added_count += 1
                else:
                    self.log_message(f"Skipped duplicate: {os.path.basename(filename)}", "INFO")

            if added_count > 0:
                self.log_message(f"Added {added_count} file(s) to queue. Total: {len(self.file_queue)}")
                self._update_file_queue_display()
            else:
                self.log_message("No new files added (duplicates skipped)", "INFO")
        else:
            self.log_message("No files selected", "INFO")

    def remove_selected_files(self):
        """Remove selected files from queue"""
        if self.file_queue_selection is None:
            self.log_message("No files selected for removal", "INFO")
            messagebox.showinfo("No Selection", "Please click on a file to select it for removal")
            return

        if self.file_queue_selection < len(self.file_queue):
            filename = self.file_queue[self.file_queue_selection]
            self.file_queue.pop(self.file_queue_selection)
            self.log_message(f"Removed: {os.path.basename(filename)}")
            self.file_queue_selection = None
            self._update_file_queue_display()
            self.log_message(f"Files remaining in queue: {len(self.file_queue)}")

    def clear_file_queue(self):
        """Clear all files from queue"""
        if len(self.file_queue) == 0:
            self.log_message("Queue is already empty", "INFO")
            return

        count = len(self.file_queue)
        self.file_queue.clear()
        self.file_queue_selection = None
        self._update_file_queue_display()
        self.column_overrides.clear()
        self.log_message(f"Cleared {count} file(s) from queue")

    def preview_selected_file(self):
        """Open preview dialog for selected file"""
        if self.file_queue_selection is None:
            self.log_message("No file selected for preview", "INFO")
            messagebox.showinfo("No Selection", "Please click on a file to select it for preview")
            return

        file_index = self.file_queue_selection
        if file_index >= len(self.file_queue):
            return

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
        self.log_text.configure(state='normal')

        # Insert text (CTkTextbox doesn't support tags like regular Text widget)
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)

        # Disable editing again
        self.log_text.configure(state='disabled')

        # Also log to Python logger
        if level == "ERROR":
            logger.error(message)
        elif level == "SUCCESS":
            logger.info(f"SUCCESS: {message}")
        else:
            logger.info(message)

    def clear_log(self):
        """Clear the log text widget"""
        self.log_text.configure(state='normal')
        self.log_text.delete("1.0", tk.END)
        self.log_text.configure(state='disabled')
        self.log_message("Log cleared")

    def update_status(self, message, color="black"):
        """Update status label"""
        # Map common color names to more visible colors in light mode
        color_map = {
            "green": "#2e7d32",  # Darker green for better visibility
            "blue": "#1565c0",   # Darker blue
            "red": "#c62828",    # Darker red
            "black": "#000000"
        }
        actual_color = color_map.get(color, color)
        self.status_label.configure(text=message, text_color=actual_color)

    def update_progress(self, value):
        """Update progress bar and percentage label"""
        self.current_progress = value
        self.progress_bar.set(value / 100)  # CTkProgressBar expects 0.0 to 1.0
        self.progress_percentage_label.configure(text=f"{int(value)}%")

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
                self.message_queue.put(("db_status", "Status: Connected", "green"))
            except Exception as e:
                self.message_queue.put(("log", f"Connection failed: {e}", "ERROR"))
                self.message_queue.put(("status", "Connection failed", "red"))
                self.message_queue.put(("db_status", "Status: Connection failed", "red"))

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
        self.convert_button.configure(state="disabled")
        self.add_files_button.configure(state="disabled")
        self.remove_files_button.configure(state="disabled")
        self.clear_queue_button.configure(state="disabled")

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
                    # Get delimiter preference for CSV files
                    delimiter = self.csv_delimiters.get(file_path, ',')
                    dataframes = get_dataframes(file_path, delimiter=delimiter)
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

                    self.message_queue.put(("log", f"  [SUCCESS] {filename} completed successfully", "SUCCESS"))
                    successful_files += 1

                except Exception as e:
                    self.message_queue.put(("log", f"  [ERROR] Failed to process {filename}: {e}", "ERROR"))
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

            # Get delimiter preference for CSV files
            delimiter = self.csv_delimiters.get(file_path, ',')
            dataframes = get_dataframes(file_path, delimiter=delimiter)
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

                self.message_queue.put(("log", f"[SUCCESS] Table '{table_name}' created successfully", "SUCCESS"))
                current_progress = int(30 + ((idx + 1) * progress_per_sheet))
                self.message_queue.put(("progress", current_progress))

            cursor.close()
            conn.close()

            self.message_queue.put(("progress", 100))
            self.message_queue.put(("log", f"[SUCCESS] All {total_sheets} table(s) created successfully!", "SUCCESS"))
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
                        self.db_status_label.configure(text=msg_data, text_color=msg_extra)
                    else:
                        self.db_status_label.configure(text=msg_data)

                elif msg_type == "enable_buttons":
                    self.convert_button.configure(state="normal")
                    self.add_files_button.configure(state="normal")
                    self.remove_files_button.configure(state="normal")
                    self.clear_queue_button.configure(state="normal")

                elif msg_type == "show_success":
                    messagebox.showinfo("Success", msg_data)

                elif msg_type == "show_error":
                    messagebox.showerror("Error", msg_data)

        except queue.Empty:
            pass

        # Schedule next check
        self.root.after(100, self.process_queue)

