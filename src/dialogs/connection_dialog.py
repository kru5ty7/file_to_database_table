"""
Connection Manager Dialog - Manage database connections
"""

import tkinter as tk
from tkinter import messagebox
import customtkinter as ctk
import json
import threading
from src.database import get_available_connections
from src.utils import encrypt_password, decrypt_password


class ConnectionManagerDialog:
    def __init__(self, parent, main_app):
        self.main_app = main_app
        self.dialog = ctk.CTkToplevel(parent)
        self.dialog.title("Manage Database Connections")
        self.dialog.geometry("800x550")

        # Enable minimize and maximize buttons (remove transient to allow window controls)
        # self.dialog.transient(parent)  # Commented out to enable min/max buttons
        self.dialog.resizable(True, True)  # Allow window resizing
        self.dialog.grab_set()

        # Load config
        self.load_config()

        # Main frame
        main_frame = ctk.CTkFrame(self.dialog)
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        self.dialog.columnconfigure(0, weight=1)
        self.dialog.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)

        # Connection list
        list_frame = ctk.CTkFrame(main_frame)
        list_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)

        ctk.CTkLabel(list_frame, text="Connections", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(10, 5))

        # Use CTkTextbox for connection list (styled like listbox)
        self.conn_textbox = ctk.CTkTextbox(list_frame, height=300, width=200, font=ctk.CTkFont(size=11))
        self.conn_textbox.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=(0, 10))
        self.conn_textbox.bind("<Button-1>", self._on_connection_click)

        # Store selection
        self.selected_connection_index = None

        # Buttons for list
        list_btn_frame = ctk.CTkFrame(list_frame, fg_color="transparent")
        list_btn_frame.grid(row=2, column=0, pady=(0, 10), padx=10)

        ctk.CTkButton(list_btn_frame, text="➕ Add", command=self.add_connection, width=90).pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(list_btn_frame, text="🗑 Delete", command=self.delete_connection, width=90, fg_color="#d32f2f", hover_color="#b71c1c").pack(side=tk.LEFT, padx=5)

        # Connection details
        details_frame = ctk.CTkFrame(main_frame)
        details_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        details_frame.columnconfigure(1, weight=1)

        ctk.CTkLabel(details_frame, text="Connection Details", font=ctk.CTkFont(size=13, weight="bold")).grid(row=0, column=0, columnspan=2, sticky=tk.W, padx=10, pady=(10, 15))

        row = 1
        ctk.CTkLabel(details_frame, text="Connection Name:").grid(row=row, column=0, sticky=tk.W, padx=(10, 5), pady=5)
        self.name_var = tk.StringVar()
        self.name_entry = ctk.CTkEntry(details_frame, textvariable=self.name_var, width=250)
        self.name_entry.grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)

        row += 1
        ctk.CTkLabel(details_frame, text="Server:").grid(row=row, column=0, sticky=tk.W, padx=(10, 5), pady=5)
        self.server_var = tk.StringVar()
        ctk.CTkEntry(details_frame, textvariable=self.server_var, width=250).grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)

        row += 1
        ctk.CTkLabel(details_frame, text="Database:").grid(row=row, column=0, sticky=tk.W, padx=(10, 5), pady=5)
        self.database_var = tk.StringVar()
        ctk.CTkEntry(details_frame, textvariable=self.database_var, width=250).grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)

        row += 1
        ctk.CTkLabel(details_frame, text="Username:").grid(row=row, column=0, sticky=tk.W, padx=(10, 5), pady=5)
        self.username_var = tk.StringVar()
        ctk.CTkEntry(details_frame, textvariable=self.username_var, width=250).grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)

        row += 1
        ctk.CTkLabel(details_frame, text="Password:").grid(row=row, column=0, sticky=tk.W, padx=(10, 5), pady=5)
        self.password_var = tk.StringVar()
        ctk.CTkEntry(details_frame, textvariable=self.password_var, width=250, show="*").grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)

        row += 1
        ctk.CTkLabel(details_frame, text="Driver:").grid(row=row, column=0, sticky=tk.W, padx=(10, 5), pady=5)
        self.driver_var = tk.StringVar(value="{ODBC Driver 17 for SQL Server}")
        ctk.CTkEntry(details_frame, textvariable=self.driver_var, width=250).grid(row=row, column=1, sticky=(tk.W, tk.E), padx=(5, 10), pady=5)

        row += 1
        details_btn_frame = ctk.CTkFrame(details_frame, fg_color="transparent")
        details_btn_frame.grid(row=row, column=0, columnspan=2, pady=(20, 10))

        ctk.CTkButton(details_btn_frame, text="💾 Save", command=self.save_connection, width=100, fg_color="#2fa572", hover_color="#26734f").pack(side=tk.LEFT, padx=5)
        ctk.CTkButton(details_btn_frame, text="✓ Test Connection", command=self.test_current_connection, width=140).pack(side=tk.LEFT, padx=5)

        # Bottom buttons
        bottom_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        bottom_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))

        ctk.CTkButton(bottom_frame, text="Close", command=self.close_dialog, width=100).pack(side=tk.RIGHT, padx=5)

        # Store connection names list
        self.connection_names = []

        # Load connections into list
        self.refresh_list()

    def _update_connection_list_display(self):
        """Update the connection textbox display"""
        self.conn_textbox.configure(state='normal')
        self.conn_textbox.delete("1.0", tk.END)
        for i, conn_name in enumerate(self.connection_names):
            prefix = "▶ " if i == self.selected_connection_index else "  "
            self.conn_textbox.insert(tk.END, f"{prefix}{conn_name}\n")
        self.conn_textbox.configure(state='disabled')

    def _on_connection_click(self, event):
        """Handle click on connection textbox"""
        try:
            index = self.conn_textbox.index(f"@{event.x},{event.y}")
            line_num = int(index.split('.')[0]) - 1
            if 0 <= line_num < len(self.connection_names):
                self.selected_connection_index = line_num
                self._update_connection_list_display()
                self.on_connection_select()
        except:
            pass

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
        self.connection_names = list(self.config.get('connections', {}).keys())
        self.selected_connection_index = None
        self._update_connection_list_display()

    def on_connection_select(self):
        """Handle connection selection"""
        if self.selected_connection_index is not None and self.selected_connection_index < len(self.connection_names):
            conn_name = self.connection_names[self.selected_connection_index]
            conn_data = self.config['connections'][conn_name]

            self.name_var.set(conn_name)
            self.name_entry.configure(state='readonly')
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
        self.name_entry.configure(state='normal')
        self.name_var.set('')
        self.server_var.set('')
        self.database_var.set('')
        self.username_var.set('')
        self.password_var.set('')
        self.driver_var.set('{ODBC Driver 17 for SQL Server}')
        self.selected_connection_index = None
        self._update_connection_list_display()
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
            self.name_entry.configure(state='readonly')

    def delete_connection(self):
        """Delete selected connection"""
        if self.selected_connection_index is None or self.selected_connection_index >= len(self.connection_names):
            self.main_app.log_message("Delete failed: No connection selected", "ERROR")
            messagebox.showwarning("Warning", "Please select a connection to delete")
            return

        conn_name = self.connection_names[self.selected_connection_index]

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
