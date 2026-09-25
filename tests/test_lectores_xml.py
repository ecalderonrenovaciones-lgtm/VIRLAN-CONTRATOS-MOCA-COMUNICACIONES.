"""Los lectores propios de .docx/.xlsx (sin librerías) deben manejar lo que escriben otras herramientas."""

import zipfile

import openpyxl

from tests.conftest import DOM_FISCAL
from virlan_bot._xlsx_utils import fila_a_dict_por_encabezado, leer_filas
from virlan_bot.ficha_extractor import extraer_ficha


def _docx_xml_crudo(ruta, parrafos_xml):
    """Word real escapa & < > en el XML (&amp; &lt; &gt;)."""
    cuerpo = "".join(f"<w:p><w:r><w:t>{p}</w:t></w:r></w:p>" for p in parrafos_xml)
    xml = ('<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="x"><w:body>' + cuerpo + "</w:body></w:document>")
    with zipfile.ZipFile(ruta, "w") as z:
        z.writestr("word/document.xml", xml)
    return ruta


def test_razon_social_con_ampersand_se_lee_desescapada(tmp_path):
    ruta = _docx_xml_crudo(
        tmp_path / "f.docx",
        ["PRUEBA &amp; ASOCIADOS SA DE CV", "500000001", "JUAN PEREZ", DOM_FISCAL, "EPR010101AB1", "5512345678", "j@p.mx"],
    )
    assert extraer_ficha(ruta).razon_social == "PRUEBA & ASOCIADOS SA DE CV"


def test_xlsx_con_cadenas_en_linea_y_compartidas(tmp_path):
    """openpyxl guarda inlineStr; Excel guarda shared strings. Ambos deben leerse."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Plan", "Cuenta", "Empresa"])
    ws.append(["ATT Armalo 599", 500000001, "PRUEBA & CIA"])
    ruta = tmp_path / "a.xlsx"
    wb.save(ruta)
    filas = fila_a_dict_por_encabezado(leer_filas(ruta), 1)
    assert filas[2] == {"Plan": "ATT Armalo 599", "Cuenta": "500000001", "Empresa": "PRUEBA & CIA"}
