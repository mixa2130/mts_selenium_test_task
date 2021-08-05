"""This module is responsible for interaction with Excel using pywin32"""
import os
import json
import time
from datetime import datetime
from collections import deque
from typing import List, NamedTuple, Tuple, Deque
import win32com.client as win32


class InputArgs(NamedTuple):
    last_name: str  # Фамилия
    first_name: str  # Имя потенциального должника
    patronymic: str  # Отчество
    date: str = ''  # Дата рождения


excel = win32.gencache.EnsureDispatch('Excel.Application')

with open('files.json', 'r', encoding='utf-8') as json_file:
    json_data: dict = json.load(json_file)
    FILES: List[dict] = json_data['files']


def write_excel_file(data: List[tuple], file_desc: dict, filename='results.xlsx'):
    """
    Writes data to excel file.
    Each function call creates a new one Sheet.
    If file with such filename doesn't exist - creates it.

    Universal function, as for fssp, as for sudrf

    :param data: data to write
    :param file_desc: file from FILES
    :param filename: Name of the excel file to write, lying in the project root
    """
    if filename in os.listdir():
        wb = excel.Workbooks.Open(os.path.join(os.getcwd(), filename))
    else:
        wb = excel.Workbooks.Add()
        wb.SaveAs(os.path.join(os.getcwd(), filename))
    new_sheet_name = str(time.time())[:31]

    # Создаст новый лист и сделает его активным
    wb.Sheets.Add().Name = new_sheet_name
    work_sh = wb.ActiveSheet

    columns_cnt: int = len(file_desc['headers'])

    # Header
    work_sh.Range(
        work_sh.Cells(1, 1),
        work_sh.Cells(1, columns_cnt)
    ).Value = file_desc['headers']

    # Data
    for row, el in enumerate(data):
        if len(el) > 2:
            work_sh.Range(
                work_sh.Cells(row + 2, 1),
                work_sh.Cells(row + 2, columns_cnt)
            ).Value = el
        else:
            # No debts
            work_sh.Cells(row + 2, 1).Value = el[0]  # name
            work_sh.Range(
                work_sh.Cells(row + 2, 2),
                work_sh.Cells(row + 2, columns_cnt)
            ).Value = el[1]

    wb.Save()
    wb.Close(True)
    excel.Application.Quit()


def read_excel_file(file_desc: dict) -> Deque:
    """
    Reads data from excel file.
    Universal function, as for fssp, as for sudrf

    :param file_desc: Name of the excel file to read, lying in the project root

    :raise pywintypes.com_error: (-2147352567, ..): If there are no such file
    """
    wb = excel.Workbooks.Open(os.path.join(os.getcwd(), file_desc['filename']))
    work_sh = wb.ActiveSheet

    filled_range = len(work_sh.UsedRange)
    columns_cnt: int = file_desc['columns_cnt']
    row_number: int = filled_range // columns_cnt
    data = deque()

    for i in range(2, row_number + 1):
        tmp: Tuple[tuple] = work_sh.Range(
            work_sh.Cells(i, 1),
            work_sh.Cells(i, columns_cnt)
        ).Value[0]

        if columns_cnt == 4:
            # В выборке есть поле date
            raw_date = datetime.strptime(str(tmp[3]), "%Y-%m-%d 00:00:00+00:00")
            birthday: str = raw_date.strftime("%d.%m.%Y")

            data.append(InputArgs(last_name=str(tmp[0]),
                                  first_name=str(tmp[1]),
                                  patronymic=str(tmp[2]),
                                  date=birthday))
        else:
            data.append(InputArgs(last_name=str(tmp[0]),
                                  first_name=str(tmp[1]),
                                  patronymic=str(tmp[2])))

    wb.Close(True)
    excel.Application.Quit()

    return data
