"""
Application entry point
"""

import tkinter as tk
from src.gui_main import FileToDBGUI


def main():
    """Main entry point for the GUI application"""
    root = tk.Tk()
    app = FileToDBGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
