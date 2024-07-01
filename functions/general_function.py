import os
import sys
import xlwings as xl
from tkinter.messagebox import showinfo
from tkinter import filedialog


def combine_files(files, start_date, end_date, is_standalone):
    path = os.path.split(files[0])[0]
    saved_as = ""
    if is_standalone == True:
        saved_as = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[
            ("Excel Workbook (.xlsx)", "*.xlsx")])
    else:
        saved_as = f"{path}/{start_date.strftime('%d-%b-%Y')} - {end_date.strftime('%d-%b-%Y')}.xlsx"

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
