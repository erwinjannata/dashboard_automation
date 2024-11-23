import configparser
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.messagebox import showinfo


def open_schedule_screen(rootWindow):
    # Initialize config file
    config = configparser.ConfigParser()
    config.read('config.ini')

    # Read required information
    base_urlApex11 = config.get('base_links', 'Apex11')
    base_urlApex12 = config.get('base_links', 'Apex12')
    base_urlApex16 = config.get('base_links', 'Apex16')

    # Generate new window
    new_window = tk.Toplevel(master=rootWindow)
    new_window.title('Jadwalkan proses')
    new_window.configure(bg='white')
