"""
Application entry point
"""

import customtkinter as ctk
from src.gui_main import FileToDBGUI

# Set appearance mode and default color theme
ctk.set_appearance_mode("Light")  # Changed to Light mode for better visibility
ctk.set_default_color_theme("blue")  # Themes: "blue" (default), "green", "dark-blue"


def main():
    """Main entry point for the GUI application"""
    root = ctk.CTk()
    app = FileToDBGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
