import configparser
import tkinter as tk
from tkinter import ttk, filedialog
from tkinter.messagebox import showinfo


def open_setting(rootWindow):
    # Initialize config file
    config = configparser.ConfigParser()
    config.read('config.ini')

    # Read required information
    base_url117 = config.get('base_links', '117')
    base_url141 = config.get('base_links', '141')
    base_url207 = config.get('base_links', '207')
    base_urlApex11 = config.get('base_links', 'Apex11')
    base_urlApex12 = config.get('base_links', 'Apex12')
    base_urlApex16 = config.get('base_links', 'Apex16')

    new_window = tk.Toplevel(master=rootWindow)
    new_window.title('Settings')
    new_window.configure(bg='white')
    new_window.resizable(0, 0)

    def save():
        config.set('base_links', '141', entry_141.get())
        config.set('base_links', '117', entry_117.get())
        config.set('base_links', '207', entry_207.get())
        config.set('base_links', 'Apex11', entry_apex11.get())
        config.set('base_links', 'Apex12', entry_apex12.get())
        config.set('base_links', 'Apex16', entry_apex16.get())

        with open('config.ini', 'w') as configfile:
            config.write(configfile)

        showinfo(title='Success', message="Saved")
        new_window.lift()

    # Dashboard 141
    label_141 = ttk.Label(new_window, text="Dashboard 141",
                          background="white", font="calibri 11 bold")
    label_141.grid(
        row=0, column=0, pady=5, padx=5, sticky='w')

    entry_141 = tk.Entry(new_window, background='white', width=50)
    entry_141.insert(0, base_url141)
    entry_141.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)

    # Dashboard 117
    label_117 = ttk.Label(new_window, text="Dashboard 117",
                          background="white", font="calibri 11 bold")
    label_117.grid(
        row=2, column=0, pady=5, padx=5, sticky='w')

    entry_117 = tk.Entry(new_window, background='white', width=50)
    entry_117.insert(0, base_url117)
    entry_117.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)

    # Dashboard 207
    label_207 = ttk.Label(new_window, text="Dashboard 207",
                          background="white", font="calibri 11 bold")
    label_207.grid(
        row=4, column=0, pady=5, padx=5, sticky='w')

    entry_207 = tk.Entry(new_window, background='white', width=50)
    entry_207.insert(0, base_url207)
    entry_207.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)

    # Apex 11
    label_apex11 = ttk.Label(new_window, text="Apex 11",
                             background="white", font="calibri 11 bold")
    label_apex11.grid(
        row=0, column=2, pady=5, padx=5, sticky='w')
    entry_apex11 = tk.Entry(new_window, background='white', width=50)
    entry_apex11.insert(0, base_urlApex11)
    entry_apex11.grid(row=1, column=2, pady=5, padx=5,
                      sticky='w', columnspan=2)

    # Apex 12
    label_apex12 = ttk.Label(new_window, text="Apex 12",
                             background="white", font="calibri 11 bold")
    label_apex12.grid(
        row=2, column=2, pady=5, padx=5, sticky='w')
    entry_apex12 = tk.Entry(new_window, background='white', width=50)
    entry_apex12.insert(0, base_urlApex12)
    entry_apex12.grid(row=3, column=2, pady=5, padx=5,
                      sticky='w', columnspan=2)

    # Apex 16
    label_apex16 = ttk.Label(new_window, text="Apex 16",
                             background="white", font="calibri 11 bold")
    label_apex16.grid(
        row=4, column=2, pady=5, padx=5, sticky='w')
    entry_apex16 = tk.Entry(new_window, background='white', width=50)
    entry_apex16.insert(0, base_urlApex16)
    entry_apex16.grid(row=5, column=2, pady=5, padx=5,
                      sticky='w', columnspan=2)

    # Save Button
    save_btn = ttk.Button(new_window, text="Save",
                          state=tk.NORMAL, width=10, command=lambda: save())
    save_btn.grid(row=6, column=3, pady=5, padx=5, sticky='e')

    new_window.mainloop()
