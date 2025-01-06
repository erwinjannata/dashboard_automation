import os
import gc
import sys
import datetime
import pandas as pd
import xlwings as xl
from tkinter import filedialog
from tkinter.messagebox import showinfo


def check_file(file, current_date, column_name):
    filename, file_extension = os.path.splitext(file)
    global df, data_date

    if file_extension == ".xlsx":
        df = pd.read_excel(file, usecols=[column_name], nrows=1)
        date_date = str(df.at[0, column_name])
        data_date = datetime.datetime.strptime(
            date_date, '%Y-%m-%d %H:%M:%S').date()
    else:
        df = pd.read_csv(
            file, usecols=[column_name], nrows=1, encoding='iso-8859-1')
        date_date = df.at[0, column_name]
        data_date = datetime.datetime.strptime(
            date_date, '%d-%b-%Y %H:%M').date()

    if data_date == current_date:
        del df
        gc.collect()
        return True
    else:
        os.remove(file)
        del df
        gc.collect()
        return False


def combine_files(files, start_date, end_date, is_standalone):
    path = os.path.split(files[0])[0]
    saved_as = ""
    if is_standalone == True:
        saved_as = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[
            ("Excel Workbook (.xlsx)", "*.xlsx")])
    else:
        saved_as = rf"{path}\{start_date.strftime('%d-%b-%Y')} - {end_date.strftime('%d-%b-%Y')}.xlsx"

    app = xl.App(visible=False)

    new_wb = xl.Book()
    new_ws = new_wb.sheets[0]
    limit = 1

    for i, file in enumerate(files):
        wb = xl.Book(file)
        ws = wb.sheets[0]
        try:
            # Count range of data to be copied
            max_col = ws.used_range[-1:, -1:].column
            max_row = ws.used_range[-1:, -1:].row

            # Copy & Paste data
            if i == 0:
                ws.range((1, 1), (max_row, max_col)).copy()
            else:
                ws.range((2, 1), (max_row, max_col)).copy()
            new_ws.range(f"A{limit}").paste(paste='values_and_number_formats')
            new_max_row = new_ws.used_range[-1:, -1:].row
            limit = (new_max_row + 1)
            wb.app.api.CutCopyMode = False

            # Close the workbook and proceed to next workbook
            wb.close()
        except Exception as e:
            wb.close()
            new_wb.close()
            app.quit()
            sys.exit()

    new_ws.autofit()
    new_wb.save(saved_as)
    new_wb.close()
    app.quit()
    if is_standalone == True:
        showinfo(title="Message", message="Proses Selesai")
    else:
        return saved_as
