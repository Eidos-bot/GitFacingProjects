import gc
from datetime import datetime, timedelta
import win32com.client
from pdf2image import convert_from_path


def converting_excel():
    o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False

    # Calendar excel would be here, but it can be any excel that you'd want included in an automatic email/
    wb_path = r""

    wb = o.Workbooks.Open(wb_path)

    ws_index_list = [1]

    # Say you want to print these sheets
    # The file name you want for the pdf version of the wb_path.
    path_to_pdf = r''

    wb.WorkSheets(ws_index_list).Select()

    wb.ActiveSheet.ExportAsFixedFormat(0, path_to_pdf)

    wb.Close(SaveChanges=True)
    o.Quit()

    converted_calendar = convert_from_path(path_to_pdf, poppler_path=r"external_items/poppler-24.08.0/Library/bin")
    new_image_path = r"Email_Code\converted_calendar.jpg"
    for image in converted_calendar:
        image.save(new_image_path)
    gc.collect()
    print(new_image_path)
    return new_image_path


def next_monday(start_date, days_from_now, weeknum, free_date):

    if start_date == "Today":
        date = datetime.today()

    else:
        ed_sate = start_date.split("/")
        date = datetime(int(ed_sate[2]),int(ed_sate[0]), int(ed_sate[1]))

    future_date = date + timedelta(days=days_from_now)

    # Find the next Monday
    while future_date.weekday() != weeknum and free_date == "No":  # 0 represents Monday
        future_date += timedelta(days=1)


    answer = f'{future_date.strftime("%A")}, {future_date.strftime("%m")}/{future_date.strftime("%d")}/{future_date.strftime("%Y")}'
    return answer


