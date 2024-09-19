from pathlib import Path
from datetime import datetime
from tkinter import ttk, filedialog
from tkcalendar import DateEntry
from tkinter.messagebox import showinfo
from functions.db117_function import inbound_data117, outbound_data117
from functions.general_function import combine_files
from functions.db141_function import outbound_data141, inbound_data141
import tkinter.scrolledtext as tkst
import tkinter as tk
import threading
import configparser


# Interface configuration
root = tk.Tk()
root.configure(bg='white')
root.title('JNE AMI Dashboard Automation v.1.1')
root.resizable(0, 0)

# Variables
mode = tk.StringVar()
is_combine = tk.BooleanVar()


def main_process():
    mode = combo_box.current()
    username = username_entry.get()
    password = password_entry.get()
    calendar_date = calendar_date_entry.get()
    calendar_date2 = calendar_date_entry2.get()
    date = datetime.strptime(calendar_date, '%m/%d/%Y')
    date2 = datetime.strptime(calendar_date2, '%m/%d/%Y')
    diff = date2 - date
    jam_penarikan = datetime.now().strftime("%d-%b-%Y %H.%M %p")
    working_directory = Path.cwd()

    if (username and password) and diff.days >= 0:
        # 141 - Outbound
        if mode == 0:
            progressbar.start()
            outbound_data141(username=username, password=password, date_thru=date2, date_from=date, loop=(
                diff.days + 1), combine=is_combine.get(), penarikan=jam_penarikan, working_dir=rf"{working_directory}\141\Outbound", log=log_box)
        # 141 - Intracity
        elif mode == 1:
            progressbar.start()
            inbound_data141(mode=2, username=username, password=password, date_thru=date2, date_from=date, loop=(
                diff.days + 1), combine=is_combine.get(), penarikan=jam_penarikan, working_dir=rf"{working_directory}\141\Intracity", log=log_box)
        # 141 - Inbound End-To-End
        elif mode == 2:
            progressbar.start()
            inbound_data141(mode=0, username=username, password=password, date_thru=date2, date_from=date, loop=(
                diff.days + 1), combine=is_combine.get(), penarikan=jam_penarikan, working_dir=rf"{working_directory}\141\Inbound End-To-End", log=log_box)
        # 141 - Inbound Summary
        elif mode == 3:
            progressbar.start()
            inbound_data141(mode=1, username=username, password=password, date_thru=date2, date_from=date, loop=(
                diff.days + 1), combine=is_combine.get(), penarikan=jam_penarikan, working_dir=rf"{working_directory}\141\Inbound Summary", log=log_box)
        # ---------- # ---------- #
        # 117 - Inbound
        elif mode == 4:
            progressbar.start()
            inbound_data117(username=username, password=password, date_thru=date2, date_from=date, loop=(
                diff.days + 1), combine=is_combine.get(), penarikan=jam_penarikan, working_dir=rf"{working_directory}\117\Inbound", log=log_box)
        # 117 - Outbound
        elif mode == 5:
            progressbar.start()
            outbound_data117(username=username, password=password, date_thru=date2, date_from=date, loop=(
                diff.days + 1), combine=is_combine.get(), penarikan=jam_penarikan, working_dir=rf"{working_directory}\117\Outbound", log=log_box)
    elif not username or not password:
        if mode == 6:
            files = filedialog.askopenfilenames(filetypes=(
                ("Excel Workbook", "*.xlsx"),
                ("Comma Separated Values", "*.csv"),
                ("All Files", "*.*"),
            ))
            progressbar.start()
            if not files:
                showinfo(title='Warning', message='File Not Found')
            else:
                combine_files(files=files, start_date=date,
                              end_date=date2, is_standalone=True)
        else:
            showinfo(title='Warning', message="Invalid User")
    else:
        showinfo(title='Warning', message="Invalid Date")
    progressbar.stop()
    combo_box.state(['!disabled'])
    combine_check.config(state='normal')
    calendar_date_entry.state(['!disabled'])
    calendar_date_entry2.state(['!disabled'])
    username_entry.configure(state='normal')
    password_entry.configure(state='normal')
    show_password.config(state='normal')
    command_btn.state(['!disabled'])
    setting_btn.state(['!disabled'])


def start_thread(event):
    global main_thread
    main_thread = threading.Thread(target=main_process, daemon=True)
    main_thread.start()
    combo_box.state(['disabled'])
    combine_check.config(state='disabled')
    calendar_date_entry.state(['disabled'])
    calendar_date_entry2.state(['disabled'])
    username_entry.configure(state='disabled')
    password_entry.configure(state='disabled')
    show_password.config(state='disabled')
    command_btn.state(['disabled'])
    setting_btn.state(['disabled'])
    root.after(5, check_thread_process)


def check_thread_process():
    if main_thread.is_alive():
        root.after(5, check_thread_process)


def toggle_check():
    global password, show_password
    if show_password.var.get():
        password_entry.config(show="*")
    else:
        password_entry.config(show="")

def open_setting():
    config = configparser.ConfigParser()
    config.read('config.ini')
    
    base_link117 = config.get('base_links', '117')
    base_link141 = config.get('base_links', '141')
    base_link207 = config.get('base_links', '207')

    new_window = tk.Toplevel(master=root)
    new_window.title('Settings')
    new_window.configure(bg='white')
    new_window.resizable(0,0)

    def save():
        config.set('base_links', '141', entry_141.get())
        config.set('base_links', '117', entry_117.get())
        config.set('base_links', '207', entry_207.get())

        with open('config.ini', 'w') as configfile:
            config.write(configfile)
        
        showinfo(title='Success', message="Saved")
        new_window.lift()


    label_141 = ttk.Label(new_window, text="Dashboard 141", background="white", font="calibri 11 bold").grid(
    row=0, column=0, pady=5, padx=5, sticky='w')

    entry_141 = tk.Entry(new_window, background='white', width=50)
    entry_141.insert(0, base_link141)
    entry_141.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)
    
    label_117 = ttk.Label(new_window, text="Dashboard 117", background="white", font="calibri 11 bold").grid(
    row=2, column=0, pady=5, padx=5, sticky='w')

    entry_117 = tk.Entry(new_window, background='white', width=50)
    entry_117.insert(0, base_link117)
    entry_117.grid(row=3, column=0, pady=5, padx=5, sticky='w', columnspan=2)
    
    label_207 = ttk.Label(new_window, text="Dashboard 207", background="white", font="calibri 11 bold").grid(
    row=4, column=0, pady=5, padx=5, sticky='w')

    entry_207 = tk.Entry(new_window, background='white', width=50)
    entry_207.insert(0, base_link207)
    entry_207.grid(row=5, column=0, pady=5, padx=5, sticky='w', columnspan=2)

    save_btn = ttk.Button(new_window, text="Save", state=tk.NORMAL, width=10, command=lambda: save())
    save_btn.grid(row=6, column=1, pady=5, padx=5, sticky='e')

    new_window.mainloop()



# GUI
combo_label = ttk.Label(root, text="Download Data",
                        background="white", font="calibri 11 bold").grid(row=0, column=0, pady=5, padx=5, sticky='w')
combo_box = ttk.Combobox(root, textvariable=mode, width=53)
combo_box['value'] = (
    '141 - Outbound',   # 0
    '141 - Intracity',  # 1
    '141 - Inbound End-To-End', # 2
    '141 - Inbound Summary', # 3
    '117 - Inbound',    # 4
    '117 - Outbound',   # 5
    'Gabung Data',  # 6
    )  
combo_box.current(0)
combo_box.grid(row=1, column=0, pady=5, padx=5, sticky='w', columnspan=2)

combine_check = tk.Checkbutton(
    root, text="Gabung Data", background="white", variable=is_combine, onvalue=True, offvalue=False, state='normal')
combine_check.grid(row=2, column=0, pady=5, padx=5, sticky='w')

calendar_label = ttk.Label(root, text="Date From", background="white",
                           font="calibri 11 bold").grid(row=3, column=0, pady=5, padx=5, sticky='w')

calendar_date_entry = DateEntry(root, selectmode='day', locale='en_US',
                                date_pattern='MM/dd/yyyy', weekendbackground='white', weekendforeground='black', width=23)
calendar_date_entry.grid(row=4, column=0, pady=5, padx=5, sticky='w')

calendar_label2 = ttk.Label(root, text="Date Thru", background="white",
                            font="calibri 11 bold").grid(row=3, column=1, pady=5, padx=5, sticky='w')

calendar_date_entry2 = DateEntry(root, selectmode='day', locale='en_US',
                                 date_pattern='MM/dd/yyyy', weekendbackground='white', weekendforeground='black', width=23)
calendar_date_entry2.grid(row=4, column=1, pady=5, padx=5, sticky='e')

username_label = ttk.Label(root, text="Username", background="white", font="calibri 11 bold").grid(
    row=5, column=0, pady=5, padx=5, sticky='w')

username_entry = tk.Entry(root, background='white', width=26)
username_entry.grid(
    row=6, column=0, pady=5, padx=5, sticky='w')

password_label = ttk.Label(root, text="Password", background="white",
                           font="calibri 11 bold").grid(row=7, column=0, pady=5, padx=5, sticky='w')

password_entry = tk.Entry(root, background='white', width=26)
password_entry.default_show_val = password_entry.config(show="*")
password_entry.config(show="*")
password_entry.grid(row=8, column=0, pady=5, padx=5, sticky='w')

show_password = tk.Checkbutton(
    root, text="Tampilkan password", onvalue=False, offvalue=True, command=toggle_check, background='white')
show_password.var = tk.BooleanVar(value=True)
show_password['variable'] = show_password.var
show_password.grid(row=9, column=0, pady=5, padx=5, sticky='w')

command_btn = ttk.Button(root, text="Start",
                         command=lambda: start_thread(None), state=tk.NORMAL, width=26)
command_btn.grid(row=6, column=1, pady=5, padx=5, sticky='e')

setting_btn = ttk.Button(root, text="Settings", command=lambda: open_setting(), state=tk.NORMAL, width=26)
setting_btn.grid(row=7, column=1, pady=5, padx=5, sticky='e')

log_label = tk.Label(root, text='Log', background='white', font='calibri 11 bold').grid(
    row=0, column=2, pady=5, padx=5, sticky='w')
log_box = tkst.ScrolledText(root, width=50, height=15)
log_box.grid(row=1, column=2, pady=5,
             padx=5, sticky='w', rowspan=8)

progressbar = ttk.Progressbar(
    root, mode='indeterminate', orient='horizontal', length=420)
progressbar.grid(row=9, column=2, pady=5, padx=5, sticky='w')

root.mainloop()
