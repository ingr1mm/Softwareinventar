# Description: This module is the entry point for the application. It creates an instance of the GUI and runs the application.

# Import the required modules
import logging
import tkinter as tk
from frontend import ApplicationGUI
from config import db_config


logging.basicConfig(level=logging.INFO)

if __name__ == '__main__':
    try:
        root = tk.Tk()
        app = ApplicationGUI(root, db_config)
        app.root.mainloop()
    except Exception as e:
        logging.error(f'Application failed to start: {e}')