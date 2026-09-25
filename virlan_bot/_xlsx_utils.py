"""Utilidades internas para leer hojas .xlsx sin dependencias externas,
devolviendo cada fila como un dict {header: valor} usando la fila de
encabezados que se indique.

Se usa para SAE 2.2.xlsx, control de renovacion.xlsx y LAYOUT DE
VINCULACION, que son hojas simples de una sola pestaña. Para escribir en
la OP (que tiene imágenes/merges que hay que preservar) se usa openpyxl en
su lugar (ver op_filler.py).
"""

from __future__ import annotations

import html
import re
import zipfile
from pathlib import Path

_ROW_RE = re.compile(r"<row[^>]*r=\"(\d+)\"[^>]*>(.*?)</row>", re.S)
_CELL_RE = re.compile(r"<c\s+([^>/]*?)(?:/>|>(.*?)</c>)", re.S)
_REF_RE = re.compile(r'r="([A-Z]+)(\d+)"')
_TYPE_RE = re.compile(r't="(\w+)"')
_VALUE_RE = re.compile(r"<v>([^<]*)</v>")


def _shared_strings(z: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    out = []
    for si in re.findall(r"<si>(.*?)</si>", xml, re.S):
        textos = re.findall(r"<t[^>]*>([^<]*)</t>", si)
        out.append(html.unescape("".join(textos)))   # &amp; -> &
    return out


def leer_filas(xlsx_path: str | Path, sheet_index: int = 1) -> dict[int, dict[str, str]]:
    """Devuelve {numero_de_fila: {columna_letra: valor}} para
    xl/worksheets/sheet<sheet_index>.xml, resolviendo shared strings."""
    xlsx_path = Path(xlsx_path)
    with zipfile.ZipFile(xlsx_path) as z:
        shared = _shared_strings(z)
        xml = z.read(f"xl/worksheets/sheet{sheet_index}.xml").decode("utf-8")

    filas: dict[int, dict[str, str]] = {}
    for row_num_str, row_xml in _ROW_RE.findall(xml):
        row_num = int(row_num_str)
        celdas: dict[str, str] = {}
        for m in _CELL_RE.finditer(row_xml):
            attrs, inner = m.group(1), m.group(2) or ""
            ref_match = _REF_RE.search(attrs)
            if not ref_match:
                continue
            col = ref_match.group(1)
            t_match = _TYPE_RE.search(attrs)
            t = t_match.group(1) if t_match else None
            v_match = _VALUE_RE.search(inner)
            val = v_match.group(1) if v_match else ""
            if t == "inlineStr":
                # cadenas en línea (openpyxl, pandas, Excel Online...): <is><t>texto</t></is>
                val = html.unescape("".join(re.findall(r"<t[^>]*>([^<]*)</t>", inner)))
            elif t == "str":
                val = html.unescape(val)
            if t == "s" and val:
                try:
                    val = shared[int(val)]
                except (ValueError, IndexError):
                    pass
            if val:
                # Algunas celdas traen espacios de no separación (\xa0) u
                # otros espacios sobrantes por copy-paste desde otras
                # fuentes; se normalizan para evitar valores como
                # '6221854808\xa0' que rompen comparaciones exactas y el
                # renderizado en Excel.
                val = val.replace("\xa0", " ").strip()
            if val:
                celdas[col] = val
        if celdas:
            filas[row_num] = celdas
    return filas


def fila_a_dict_por_encabezado(
    filas: dict[int, dict[str, str]], fila_encabezado: int
) -> dict[str, dict[str, str]]:
    """Reindexado por nombre de columna en vez de letra, usando la fila de
    encabezados indicada. Devuelve {numero_de_fila: {nombre_columna: valor}}."""
    encabezados = filas.get(fila_encabezado, {})
    resultado: dict[str, dict[str, str]] = {}
    for num_fila, celdas in filas.items():
        if num_fila <= fila_encabezado:
            continue
        fila_nombrada = {}
        for col, valor in celdas.items():
            nombre = encabezados.get(col, col)
            fila_nombrada[nombre] = valor
        if fila_nombrada:
            resultado[num_fila] = fila_nombrada
    return resultado
