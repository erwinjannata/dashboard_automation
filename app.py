import os
import time
import schedule
import threading
import tkinter as tk
import tkinter.scrolledtext as tkst
from pathlib import Path
from datetime import datetime
from tkcalendar import DateEntry
from partials.settings import open_setting
from partials.updates import check_update_function
from tkinter import ttk, filedialog
from tkinter.messagebox import showinfo
from functions.general_function import combine_files
from functions.db141_function import DB141
from functions.db117_function import DB117
from tktimepicker import SpinTimePickerOld
from tktimepicker import constants
from threading import Event

if __name__ == "__main__":
    # Interface configuration
    root = tk.Tk()
    root.configure(bg='white')
    root.title(f'JNE AMI Dashboard Automation v.2.7')
    root.resizable(0, 0)
    root.columnconfigure(0, weight=4)
    root.columnconfigure(1, weight=1)

    # Variables
    mode = tk.StringVar()
    apex = tk.StringVar()
    is_combine = tk.BooleanVar()
    is_apex = tk.BooleanVar()
    is_scheduled = tk.BooleanVar()
    apex_names = ["Apex11", "Apex12", "Apex16"]

    scheduler_thread = None
    stop_scheduler = Event()

    def determine_process():
        mode = combo_box.current()
        date = datetime.strptime(calendar_date_entry.get(), '%m/%d/%Y')
        date2 = datetime.strptime(calendar_date_entry2.get(), '%m/%d/%Y')
        diff = date2 - date

        if (username_entry.get() and password_entry.get()) and diff.days >= 0:
            if is_scheduled.get() == True:
                global scheduler_thread, stop_scheduler
                stop_scheduler.set()
                if scheduler_thread and scheduler_thread.is_alive():
                    scheduler_thread.join(timeout=1)

                stop_scheduler.clear()

                scheduled_time = time_picker.time()
                formatted_time = f'{scheduled_time[0]:02}:{scheduled_time[1]:02}'

                schedule.clear()
                schedule.every().day.at(formatted_time).do(main_process).tag('download_task')

                log_box.insert(
                    tk.END, f"{datetime.now().strftime('%H:%M')} - Scheduled download at {formatted_time} \n")
                log_box.see("end")

                scheduler_thread = threading.Thread(
                    target=run_scheduler, daemon=True)
                scheduler_thread.start()
            else:
                main_process()
        elif mode == 6:
            if is_scheduled.get() == True:
                showinfo(title='Warning',
                         message="Proses ini tidak bisa dilakukan secara terjadwal")
                close_thread()
            else:
                main_process()
        elif not username_entry.get() or not password_entry.get():
            showinfo(title='Warning', message="Invalid User")
            close_thread()
        else:
            showinfo(title='Warning', message="Invalid Date")
            close_thread()

    def run_scheduler():
        while not stop_scheduler.is_set():
            schedule.run_pending()
            time.sleep(1)

    def abort_scheduled_job():
        global scheduler_thread, stop_scheduler
        stop_scheduler.set()
        schedule.clear('download_task')

        if scheduler_thread and scheduler_thread.is_alive():
            scheduler_thread.join(timeout=1)
            scheduler_thread = None

        stop_scheduler.clear()

        log_box.insert(
            tk.END, f"{datetime.now().strftime('%H:%M')} - Scheduled download cancelled \n")
        log_box.see("end")
        close_thread()

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
        apex_name = apex_names[apex_combo.current()]
        global apex_file_name
        apex_file_name = ""
        driver_location = rf"{working_directory}\webdriver\msedgedriver.exe"

        if (username and password) and diff.days >= 0:
            if is_apex.get() == True:
                try:
                    apex_file_name = filedialog.asksaveasfile(defaultextension=".csv", filetypes=[
                        ("CSV (Comma delimited)", "*.csv")]).name
                except:
                    showinfo(title="Error",
                             message="Pilih lokasi penyimpanan file apex")
                    close_thread()
                    return "Error"

            # Initialize Class Dashboard 141
            dashboard141 = DB141(mode=0,
                                 username_141=username,
                                 password_141=password,
                                 username_apex=username_apex_entry.get(),
                                 password_apex=password_apex_entry.get(),
                                 date_from=date,
                                 date_thru=date2,
                                 loop=(diff.days + 1),
                                 is_combine=is_combine.get(),
                                 is_apex=is_apex.get(),
                                 working_dir="",
                                 apex_type=apex_name,
                                 apex_file_name=apex_file_name,
                                 penarikan=jam_penarikan,
                                 log=log_box,
                                 driver=driver_location
                                 )

            # Initialize Class Dashboard 117
            dashboard117 = DB117(username_117=username,
                                 password_117=password,
                                 username_apex=username_apex_entry.get(),
                                 password_apex=password_apex_entry.get(),
                                 date_from=date,
                                 date_thru=date2,
                                 loop=(diff.days + 1),
                                 is_combine=is_combine.get(),
                                 is_apex=is_apex.get(),
                                 working_dir="",
                                 apex_type=apex_name,
                                 apex_file_name=apex_file_name,
                                 penarikan=jam_penarikan,
                                 log=log_box,
                                 driver=driver_location)

            # 141 - Outbound
            if mode == 0:
                progressbar.start()
                dashboard141.working_dir = rf"{working_directory}\141\Outbound"
                dashboard141.outbound_data141()
            # 141 - Intracity
            elif mode == 1:
                progressbar.start()
                dashboard141.mode = 2
                dashboard141.working_dir = rf"{working_directory}\141\Intracity"
                dashboard141.inbound_data141()
            # 141 - Inbound End-To-End
            elif mode == 2:
                progressbar.start()
                dashboard141.mode = 0
                dashboard141.working_dir = rf"{working_directory}\141\Inbound End-To-End"
                dashboard141.inbound_data141()
            # 141 - Inbound Summary
            elif mode == 3:
                progressbar.start()
                dashboard141.mode = 1
                dashboard141.working_dir = rf"{working_directory}\141\Inbound Summary"
                dashboard141.inbound_data141()
            # ---------- # ---------- #
            # 117 - Inbound
            elif mode == 4:
                progressbar.start()
                dashboard117.working_dir = rf"{working_directory}\117\Inbound"
                dashboard117.inbound_data117()
            # 117 - Outbound
            elif mode == 5:
                progressbar.start()
                dashboard117.working_dir = rf"{working_directory}\117\Outbound"
                dashboard117.outbound_data117()
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
        close_thread()
        return schedule.CancelJob

    def start_thread(event):
        global main_thread
        main_thread = threading.Thread(target=determine_process, daemon=True)
        main_thread.start()
        combo_box.state(['disabled'])
        apex_combo.state(['disabled'])
        combine_check.config(state='disabled')
        apex_check.config(state='disabled')
        calendar_date_entry.state(['disabled'])
        calendar_date_entry2.state(['disabled'])
        username_entry.configure(state='disabled')
        password_entry.configure(state='disabled')
        username_apex_entry.configure(state='disabled')
        password_apex_entry.configure(state='disabled')
        show_password.config(state='disabled')
        show_password_apex.config(state='disabled')
        command_btn.state(['disabled'])
        setting_btn.state(['disabled'])
        time_picker.configureAll(state='disabled')
        scheduler_check.config(state='disabled')
        if is_scheduled.get() == True:
            cancel_btn.state(['!disabled'])
        menu_bar.entryconfig('File', state='disabled')
        menu_bar.entryconfig('Help', state='disabled')
        root.after(5, check_thread_process)

    def close_thread():
        progressbar.stop()
        combo_box.state(['!disabled'])
        combine_check.config(state='normal')
        apex_check.config(state='normal')
        calendar_date_entry.state(['!disabled'])
        calendar_date_entry2.state(['!disabled'])
        username_entry.configure(state='normal')
        password_entry.configure(state='normal')
        show_password.config(state='normal')
        show_password_apex.config(state='normal')
        command_btn.state(['!disabled'])
        setting_btn.state(['!disabled'])
        scheduler_check.config(state='normal')
        if is_scheduled.get() == True:
            time_picker.configureAll(state='normal')
        if is_apex.get() == True:
            apex_combo.state(['!disabled'])
            username_apex_entry.configure(state='normal')
            password_apex_entry.configure(state='normal')
        cancel_btn.state(['disabled'])
        menu_bar.entryconfig('File', state='normal')
        menu_bar.entryconfig('Help', state='normal')

    def check_thread_process():
        if main_thread.is_alive():
            root.after(5, check_thread_process)

    def start_update_thread(event):
        global main_thread
        main_thread = threading.Thread(
            target=lambda: check_update_function(log_box=log_box, close_thread=close_thread), daemon=True)
        main_thread.start()
        progressbar.start()
        combo_box.state(['disabled'])
        apex_combo.state(['disabled'])
        combine_check.config(state='disabled')
        apex_check.config(state='disabled')
        calendar_date_entry.state(['disabled'])
        calendar_date_entry2.state(['disabled'])
        username_entry.configure(state='disabled')
        password_entry.configure(state='disabled')
        username_apex_entry.configure(state='disabled')
        password_apex_entry.configure(state='disabled')
        show_password.config(state='disabled')
        show_password_apex.config(state='disabled')
        command_btn.state(['disabled'])
        setting_btn.state(['disabled'])
        time_picker.configureAll(state='disabled')
        cancel_btn.state(['!disabled'])
        menu_bar.entryconfig('File', state='disabled')
        menu_bar.entryconfig('Help', state='disabled')
        root.after(5, check_thread_process)

    def toggle_check():
        global password, show_password
        if show_password.var.get():
            password_entry.config(show="")
        else:
            password_entry.config(show="*")

    def toggle_check2():
        global password, show_password_apex
        if show_password_apex.var.get():
            password_apex_entry.config(show="")
        else:
            password_apex_entry.config(show="*")

    def toggle_apex_option():
        if is_apex.get() == True:
            combine_check.select()
            apex_combo.state(['!disabled'])
            username_apex_entry.configure(state='normal')
            password_apex_entry.configure(state='normal')
            show_password_apex.config(state='normal')
        else:
            apex_combo.state(['disabled'])
            username_apex_entry.configure(state='disabled')
            password_apex_entry.configure(state='disabled')
            show_password_apex.config(state='disabled')

    def toggle_scheduled_option():
        if is_scheduled.get() == True:
            time_picker.configureAll(state='normal')
        else:
            time_picker.configureAll(state='disabled')

    def open_directory():
        os.startfile(Path.cwd())

    def launch_manual_book():
        file = rf"{Path.cwd()}\Manual Book Automasi Dashboard JNE AMI.pdf"
        os.startfile(file)

    # GUI
    # Download type option
    combo_label = ttk.Label(root, text="Download Data",
                            background="white", font="calibri 11 bold")
    combo_label.grid(row=0, column=0, pady=5, padx=5, sticky='w')
    combo_box = ttk.Combobox(root, textvariable=mode, width=23)
    combo_box['value'] = (
        '141 - Outbound',   # 0
        '141 - Intracity',  # 1
        '141 - Inbound End-To-End',  # 2
        '141 - Inbound Summary',  # 3
        '117 - Inbound',    # 4
        '117 - Outbound',   # 5
        'Gabung Data',  # 6
    )
    combo_box.current(0)
    combo_box.grid(row=1, column=0, pady=5, padx=5, sticky='w')

    # Apex Combobox
    apex_label = ttk.Label(root, text="Apex DB",
                           background="white", font="calibri 11 bold")
    apex_label.grid(row=0, column=2, pady=5, padx=5, sticky='w')
    apex_combo = ttk.Combobox(root, textvariable=apex,
                              width=23, state='disabled')
    apex_combo['value'] = (
        'Apex 11',
        'Apex 12',
        'Apex 16'
    )
    apex_combo.current(0)
    apex_combo.grid(row=1, column=2, pady=5, padx=5, sticky='e')

    # Combine data check
    combine_check = tk.Checkbutton(
        root, text="Gabung Data", background="white", variable=is_combine, onvalue=True, offvalue=False, state='normal')
    combine_check.grid(row=2, column=0, pady=5, padx=5, sticky='w')

    # Apex data check
    apex_check = tk.Checkbutton(root, text="Upload ke apex", background='white',
                                variable=is_apex, onvalue=True, offvalue=False, state='normal', command=toggle_apex_option)
    apex_check.grid(row=2, column=2, pady=5, padx=5, sticky='w')

    ttk.Separator(root, orient='horizontal').grid(
        row=3, column=0, padx=15, pady=5, sticky='ew', columnspan=3)

    # Date From
    calendar_label = ttk.Label(root, text="Date From", background="white",
                               font="calibri 11 bold")
    calendar_label.grid(row=4, column=0, pady=5, padx=5, sticky='w')

    calendar_date_entry = DateEntry(root, selectmode='day', locale='en_US',
                                    date_pattern='MM/dd/yyyy', weekendbackground='white', weekendforeground='black', width=23)
    calendar_date_entry.grid(row=5, column=0, pady=5, padx=5, sticky='w')

    # Date To
    calendar_label2 = ttk.Label(root, text="Date Thru", background="white",
                                font="calibri 11 bold")
    calendar_label2.grid(row=4, column=2, pady=5, padx=5, sticky='w')

    calendar_date_entry2 = DateEntry(root, selectmode='day', locale='en_US',
                                     date_pattern='MM/dd/yyyy', weekendbackground='white', weekendforeground='black', width=23)
    calendar_date_entry2.grid(row=5, column=2, pady=5, padx=5, sticky='e')

    ttk.Separator(root, orient='horizontal').grid(
        row=6, column=0, padx=15, pady=5, sticky='ew', columnspan=3)

    # Username Form
    username_label = ttk.Label(root, text="Username", background="white", font="calibri 11 bold").grid(
        row=7, column=0, pady=5, padx=5, sticky='w')

    username_entry = tk.Entry(root, background='white', width=26)
    username_entry.grid(
        row=8, column=0, pady=5, padx=5, sticky='w')

    # Password Form
    password_label = ttk.Label(root, text="Password", background="white",
                               font="calibri 11 bold").grid(row=9, column=0, pady=5, padx=5, sticky='w')

    password_entry = tk.Entry(root, background='white', width=26)
    password_entry.default_show_val = password_entry.config(show="*")
    password_entry.config(show="*")
    password_entry.grid(row=10, column=0, pady=5, padx=5, sticky='w')

    # Show Password
    show_password = tk.Checkbutton(
        root, text="Tampilkan password", onvalue=True, offvalue=False, command=toggle_check, background='white')
    show_password.var = tk.BooleanVar(value=False)
    show_password['variable'] = show_password.var
    show_password.grid(row=11, column=0, pady=5, padx=5, sticky='w')

    ttk.Separator(root, orient='vertical').grid(
        row=7, column=1, pady=10, padx=10, sticky='ns', rowspan=5)

    # Username Apex Form
    username_apex_label = ttk.Label(root, text="Username Apex", background="white", font="calibri 11 bold").grid(
        row=7, column=2, pady=5, padx=5, sticky='w')

    username_apex_entry = tk.Entry(
        root, background='white', width=26, state='disabled')
    username_apex_entry.grid(
        row=8, column=2, pady=5, padx=5, sticky='e')

    # Password Apex Form
    password_apex_label = ttk.Label(root, text="Password Apex", background="white",
                                    font="calibri 11 bold").grid(row=9, column=2, pady=0, padx=5, sticky='w')

    password_apex_entry = tk.Entry(
        root, background='white', width=26, state='disabled')
    password_apex_entry.default_show_val = password_apex_entry.config(show="*")
    password_apex_entry.config(show="*")
    password_apex_entry.grid(row=10, column=2, pady=0, padx=5, sticky='e')

    # Show Password Apex
    show_password_apex = tk.Checkbutton(
        root, text="Tampilkan password", onvalue=True, offvalue=False, command=toggle_check2, background='white', state='disabled')
    show_password_apex.var = tk.BooleanVar(value=False)
    show_password_apex['variable'] = show_password_apex.var
    show_password_apex.grid(row=11, column=2, pady=5, padx=5, sticky='w')

    ttk.Separator(root, orient='horizontal').grid(
        row=12, column=0, padx=15, pady=5, sticky='ew', columnspan=3)

    password_apex_entry = tk.Entry(
        root, background='white', width=26, state='disabled')
    password_apex_entry.default_show_val = password_apex_entry.config(show="*")
    password_apex_entry.config(show="*")
    password_apex_entry.grid(row=10, column=2, pady=0, padx=5, sticky='e')

    ttk.Separator(root, orient='horizontal').grid(
        row=12, column=0, padx=15, pady=5, sticky='ew', columnspan=3)

    # Scheduler
    ttk.Label(root, text="Penarikan Terjadwal", background="white",
              font="calibri 11 bold").grid(row=13, column=0, pady=0, padx=5, sticky='w')
    time_picker = SpinTimePickerOld(root)
    time_picker.addAll(constants.HOURS24)
    time_picker.configureAll(state='disabled')
    time_picker.grid(row=14, column=0, padx=10,
                     pady=5, sticky='ew', columnspan=3)

    scheduler_check = tk.Checkbutton(
        root, text="Jadwalkan Penarikan", onvalue=True, offvalue=False, command=toggle_scheduled_option, background='white', variable=is_scheduled)
    scheduler_check.grid(row=16, column=0, pady=5, padx=5, sticky='w')

    ttk.Separator(root, orient='horizontal').grid(
        row=17, column=0, padx=15, pady=5, sticky='ew', columnspan=3)

    # Start Process Button
    command_btn = ttk.Button(root, text="Start",
                             command=lambda: start_thread(None), state=tk.NORMAL, width=26)
    command_btn.grid(row=18, column=2, pady=5, padx=5, sticky='e')

    # Cancel Schedule Button
    cancel_btn = ttk.Button(root, text='Batalkan Penarikan Terjadwal',
                            command=lambda: abort_scheduled_job(), state=tk.DISABLED, width=26)
    cancel_btn.grid(row=19, column=2, pady=5, padx=5, sticky='e')

    # Setting Button
    setting_btn = ttk.Button(
        root, text="Pengaturan", command=lambda: open_setting(rootWindow=root), state=tk.NORMAL, width=26)
    setting_btn.grid(row=18, column=0, pady=5, padx=5, sticky='w')

    # Open Directory Button
    open_dir_btn = ttk.Button(root, text="Folder Unduhan",
                              state=tk.NORMAL, width=26, command=lambda: open_directory())
    open_dir_btn.grid(row=19, column=0, pady=5, padx=5, sticky='w')

    # Log Box
    log_label = tk.Label(root, text='Log', background='white', font='calibri 11 bold').grid(
        row=0, column=3, pady=5, padx=5, sticky='w')
    log_box = tkst.ScrolledText(root, width=55, height=18)
    log_box.grid(row=1, column=3, pady=0,
                 padx=5, sticky='ns', rowspan=18)

    # Progress Bar
    progressbar = ttk.Progressbar(
        root, mode='indeterminate', orient='horizontal', length=465)
    progressbar.grid(row=19, column=3, pady=5, padx=5, sticky='w')

    # Menu Bar
    menu_bar = tk.Menu(root)
    file = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="File", menu=file)
    file.add_command(label='Folder Unduhan', command=lambda: open_directory())
    file.add_command(label='Pengaturan',
                     command=lambda: open_setting(rootWindow=root))
    file.add_command(label='Manual Book', command=lambda: launch_manual_book())

    file.add_separator()
    file.add_command(label='Tutup Aplikasi', command=root.destroy)

    help = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Help", menu=help)
    help.add_command(label='Periksa Update',
                     command=lambda: start_update_thread(None))

    root.config(menu=menu_bar)
    root.mainloop()
