from datetime import datetime

import win32com.client as win32
import time
import os
from typing import List, NamedTuple, Tuple


class InputArgs(NamedTuple):
    first_name: str
    last_name: str
    patronymic: str
    date: str


excel = win32.gencache.EnsureDispatch('Excel.Application')


def write_excel_file(data: List[tuple], filename='results.xlsx'):
    if filename in os.listdir():
        wb = excel.Workbooks.Open(os.path.join(os.getcwd(), filename))
    else:
        wb = excel.Workbooks.Add()
        wb.SaveAs(os.path.join(os.getcwd(), filename))
    new_sheet_name = str(time.time())[:31]

    # Создаст новый лист и сделает его активным
    wb.Sheets.Add().Name = new_sheet_name
    work_sh = wb.ActiveSheet

    for row, el in enumerate(data):
        if len(el) > 2:
            work_sh.Range(f"A{row + 2}:G{row + 2}").Value = el
        else:
            # No debts
            work_sh.Cells(row + 2, 1).Value = el[0]  # name
            work_sh.Range(f"B{row + 2}:G{row + 2}").Value = el[1]

    wb.Save()
    wb.Close(True)
    excel.Application.Quit()


def read_excel_file(filename='input.xlsx'):
    wb = excel.Workbooks.Open(os.path.join(os.getcwd(), filename))
    work_sh = wb.ActiveSheet

    filled_range = len(work_sh.UsedRange)
    row_number = filled_range // 4
    data = list()

    for i in range(2, row_number + 1):
        tmp: Tuple[tuple] = work_sh.Range(
            work_sh.Cells(i, 1),
            work_sh.Cells(i, 4)
        ).Value[0]

        raw_date = datetime.strptime(str(tmp[3]), "%Y-%m-%d 00:00:00+00:00")
        birthday = raw_date.strftime("%d.%m.%Y")

        data.append(InputArgs(last_name=str(tmp[0]),
                              first_name=str(tmp[1]),
                              patronymic=str(tmp[2]),
                              date=str(birthday)))

    wb.Close(True)
    excel.Application.Quit()

    return data


if __name__ == '__main__':
    print(read_excel_file())
