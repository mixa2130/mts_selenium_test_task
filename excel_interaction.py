"""This module is responsible for interaction with Excel using pywin32"""
import time
import os
from datetime import datetime
from collections import deque
from typing import List, NamedTuple, Tuple, Deque
import win32com.client as win32


class InputArgs(NamedTuple):
    first_name: str  # Имя потенциального должника
    last_name: str  # Фамилия
    patronymic: str  # Отчество
    date: str  # Дата рождения


excel = win32.gencache.EnsureDispatch('Excel.Application')
fssp_column_names: tuple = (
    "Должник (физ. лицо: ФИО, дата и место рождения; юр. лицо: наименование, юр. адрес, фактический адрес)",
    "Исполнительное производство (номер, дата возбуждения)",
    "Реквизиты исполнительного документа (вид, дата принятия органом, номер, наименование органа,"
    " выдавшего исполнительный документ)",
    "Дата, причина окончания или прекращения ИП (статья, часть, пункт основания)",
    "Предмет исполнения, сумма непогашенной задолженности",
    "Отдел судебных приставов (наименование, адрес)",
    "Судебный пристав-исполнитель, телефон для получения информации"
)


def write_excel_file(data: List[tuple], filename='results.xlsx'):
    """
    Writes data to excel file.
    Each function call creates a new one Sheet.
    If file with such name doesn't exist - creates it.

    :param data: data to write
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

    # Header
    work_sh.Range("A1:G1").Value = fssp_column_names

    # Data
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


def read_excel_file(filename='input.xlsx') -> Deque:
    """
    Reads data from excel file.

    :param filename: Name of the excel file to read, lying in the project root

    :raise pywintypes.com_error: (-2147352567, ..): If there are no such file
    """
    wb = excel.Workbooks.Open(os.path.join(os.getcwd(), filename))
    work_sh = wb.ActiveSheet

    filled_range = len(work_sh.UsedRange)
    row_number: int = filled_range // 4
    data = deque()

    for i in range(2, row_number + 1):
        tmp: Tuple[tuple] = work_sh.Range(
            work_sh.Cells(i, 1),
            work_sh.Cells(i, 4)
        ).Value[0]

        raw_date = datetime.strptime(str(tmp[3]), "%Y-%m-%d 00:00:00+00:00")
        birthday: str = raw_date.strftime("%d.%m.%Y")

        data.append(InputArgs(last_name=str(tmp[0]),
                              first_name=str(tmp[1]),
                              patronymic=str(tmp[2]),
                              date=birthday))

    wb.Close(True)
    excel.Application.Quit()

    return data
