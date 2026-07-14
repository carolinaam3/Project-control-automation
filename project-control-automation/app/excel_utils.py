from pathlib import Path
from typing import Any

from win32com.client import DispatchEx


def read_cell_after_recalc(file_path: str | Path, sheet_name: str, cell_ref: str) -> Any:
    """
    Abre um arquivo Excel, força recálculo, lê uma célula e salva o arquivo.
    """
    file_path = Path(file_path).resolve()

    excel = None
    workbook = None

    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        workbook = excel.Workbooks.Open(str(file_path))

        try:
            workbook.RefreshAll()
        except Exception:
            pass

        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        excel.CalculateFullRebuild()

        value = workbook.Worksheets(sheet_name).Range(cell_ref).Value

        workbook.Save()
        return value

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()


def recalculate_and_save(file_path: str | Path) -> None:
    """
    Abre um arquivo Excel, recalcula tudo e salva.
    """
    file_path = Path(file_path).resolve()

    excel = None
    workbook = None

    try:
        excel = DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False

        workbook = excel.Workbooks.Open(str(file_path))

        try:
            workbook.RefreshAll()
        except Exception:
            pass

        try:
            excel.CalculateUntilAsyncQueriesDone()
        except Exception:
            pass

        excel.CalculateFullRebuild()
        workbook.Save()

    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()