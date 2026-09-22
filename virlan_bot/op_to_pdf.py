"""Exporta la hoja OP_OK de un .xlsx ya lleno a PDF usando Excel real
(win32com). No se usa LibreOffice: el archivo trae imágenes embebidas y un
área de impresión configurada que Excel respeta con fidelidad."""

from __future__ import annotations

from pathlib import Path

import win32com.client as win32

_HOJA = "OP_OK"
_XL_TYPE_PDF = 0


def exportar_op_a_pdf(xlsx_path: str | Path, pdf_path: str | Path, hoja: str = _HOJA) -> None:
    pdf_path = Path(pdf_path)
    pdf_path.parent.mkdir(parents=True, exist_ok=True)

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(Path(xlsx_path).resolve()))
        try:
            ws = wb.Worksheets(hoja)
            ws.ExportAsFixedFormat(_XL_TYPE_PDF, str(pdf_path.resolve()))
        finally:
            wb.Close(False)
    finally:
        excel.Quit()
