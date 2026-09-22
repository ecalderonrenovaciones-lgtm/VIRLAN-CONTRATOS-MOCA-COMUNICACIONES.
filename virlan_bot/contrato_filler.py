"""Llena la página 1 del CONTRATO PDF sustituyendo el texto de un cliente
anterior por el del cliente nuevo, en las coordenadas fijas calibradas en
config/contrato_fieldmap.v1.json.

Método: redacción (borra el texto viejo dentro del bbox) + inserción de
texto nuevo con la misma fuente/tamaño (PyMuPDF). Las páginas 2 y 3 se
copian sin modificar (verificado idénticas entre clientes distintos).

Nunca trunca en silencio: si un valor no cabe en su bbox ni reduciendo el
tamaño dentro de un rango razonable, se registra en `desbordes` para que la
capa de revisión humana lo marque.
"""

from __future__ import annotations

from dataclasses import dataclass, field as dc_field
from pathlib import Path

import fitz

_TAMANO_MINIMO_ABSOLUTO = 4.0  # nunca reducir una fuente por debajo de este tamaño en puntos
# (regla confirmada por el usuario 2026-09-22: ante desborde, reducir el tamaño
# hasta que quepa en el lugar, en vez de alertar con un tope fijo del 75%)

# El PDF original usa fuentes ArialMT/Arial-BoldMT embebidas; PyMuPDF solo
# admite insertar texto nuevo con sus 14 fuentes base. Helvetica tiene
# métricas prácticamente idénticas a Arial, así que se usa como sustituto.
_FONT_ALIASES = {
    "ArialMT": "helv",
    "Arial": "helv",
    "Arial-BoldMT": "hebo",
    "Arial-ItalicMT": "heit",
    "Arial-BoldItalicMT": "hebi",
}


def _fuente_pymupdf(fontname: str) -> str:
    return _FONT_ALIASES.get(fontname, fontname)


@dataclass
class ResultadoLlenado:
    campos_omitidos: list[str] = dc_field(default_factory=list)
    desbordes: list[str] = dc_field(default_factory=list)


def _ancho_texto(texto: str, fontname: str, fontsize: float) -> float:
    return fitz.get_text_length(texto, fontname=_fuente_pymupdf(fontname), fontsize=fontsize)


def _insertar_con_ajuste(
    page: "fitz.Page",
    bbox: list[float],
    texto: str,
    fontname: str,
    fontsize: float,
    nombre_campo: str,
    resultado: ResultadoLlenado,
) -> None:
    ancho_disponible = bbox[2] - bbox[0]
    size = fontsize
    while _ancho_texto(texto, fontname, size) > ancho_disponible and size > _TAMANO_MINIMO_ABSOLUTO:
        size -= 0.25

    if _ancho_texto(texto, fontname, size) > ancho_disponible:
        resultado.desbordes.append(
            f"{nombre_campo}: {texto!r} no cabe en el espacio disponible "
            f"({ancho_disponible:.1f}pt) ni reduciendo la fuente hasta el "
            f"mínimo legible ({_TAMANO_MINIMO_ABSOLUTO:.1f}pt); revisar manualmente."
        )

    baseline_y = bbox[3] - fontsize * 0.15
    page.insert_text(
        (bbox[0], baseline_y), texto, fontname=_fuente_pymupdf(fontname), fontsize=size, color=(0, 0, 0)
    )


def _insertar_rfc_en_casillas(
    page: "fitz.Page", campo_rfc: dict, rfc: str, resultado: ResultadoLlenado
) -> None:
    bbox = campo_rfc["bbox"]
    fontname = campo_rfc["font"]
    fontsize = campo_rfc["size"]

    rfc = rfc.strip().upper()
    # El RFC tiene 12 caracteres (persona moral) o 13 (persona física); el
    # machote solo calibró casillas para 12, así que la cantidad real de
    # casillas se ajusta al RFC recibido (nunca se truncan caracteres) y el
    # ancho de casilla se recalcula para que todos quepan en el mismo bbox
    # (regla confirmada por el usuario 2026-09-22).
    n = len(rfc)
    if n not in (12, 13):
        resultado.desbordes.append(
            f"rfc_cliente: {rfc!r} tiene {n} caracteres (se esperaban 12 o 13); "
            f"se insertará tal cual, revisar manualmente."
        )

    ancho_casilla = (bbox[2] - bbox[0]) / n
    baseline_y = bbox[3] - fontsize * 0.15
    for i, ch in enumerate(rfc):
        ancho_char = _ancho_texto(ch, fontname, fontsize)
        x = bbox[0] + i * ancho_casilla + (ancho_casilla - ancho_char) / 2
        page.insert_text(
            (x, baseline_y), ch, fontname=_fuente_pymupdf(fontname), fontsize=fontsize, color=(0, 0, 0)
        )


def _insertar_checkbox(page: "fitz.Page", checkbox: dict) -> None:
    bbox = checkbox["bbox"]
    fontname = checkbox["font"]
    fontsize = checkbox["size"]
    marca = checkbox["marca"]
    ancho_marca = _ancho_texto(marca, fontname, fontsize)
    x = bbox[0] + ((bbox[2] - bbox[0]) - ancho_marca) / 2
    baseline_y = bbox[3] - fontsize * 0.1
    page.insert_text(
        (x, baseline_y), marca, fontname=_fuente_pymupdf(fontname), fontsize=fontsize, color=(0, 0, 0)
    )


def llenar_contrato(
    plantilla_pdf: str | Path,
    salida_pdf: str | Path,
    fieldmap: dict,
    valores: dict[str, str],
    rfc_cliente: str | None,
    checkboxes_activos: list[str] | None = None,
) -> ResultadoLlenado:
    """Genera salida_pdf a partir de plantilla_pdf, reemplazando los campos
    variables de la página indicada en fieldmap['pagina'] con `valores`
    (nombre_campo -> texto nuevo) y marcando los checkboxes en
    `checkboxes_activos` (nombres de config['checkboxes']).

    Campos de fieldmap['campos_texto'] ausentes en `valores` se omiten (no
    se tocan) y se listan en resultado.campos_omitidos.
    """
    resultado = ResultadoLlenado()
    doc = fitz.open(str(plantilla_pdf))
    pagina_idx = fieldmap["pagina"]
    page = doc[pagina_idx]

    campos_texto = fieldmap["campos_texto"]
    bboxes_a_redactar = []

    for nombre, spec in campos_texto.items():
        if nombre in valores and valores[nombre]:
            bboxes_a_redactar.append(spec["bbox"])
        else:
            resultado.campos_omitidos.append(nombre)

    if rfc_cliente:
        bboxes_a_redactar.append(fieldmap["campo_rfc_cliente"]["bbox"])

    # El machote es un PDF real reciclado de un cliente anterior: sus
    # casillas (Tipo de Contratación, Acepto/No Acepto, etc.) pueden traer
    # una "X" de ese cliente. Se redactan TODAS las casillas conocidas del
    # fieldmap (se marquen o no para este cliente) para nunca dibujar una X
    # nueva encima de una vieja — regla confirmada por el usuario: solo debe
    # verse una X por casilla, nunca dos superpuestas.
    for checkbox in fieldmap["checkboxes"].values():
        if isinstance(checkbox, dict):  # se salta claves de metadata como "_nota"
            bboxes_a_redactar.append(checkbox["bbox"])

    for bbox in bboxes_a_redactar:
        page.add_redact_annot(fitz.Rect(bbox))
    if bboxes_a_redactar:
        page.apply_redactions()

    for nombre, spec in campos_texto.items():
        if nombre in valores and valores[nombre]:
            texto = spec.get("prefijo", "") + str(valores[nombre])
            _insertar_con_ajuste(
                page, spec["bbox"], texto, spec["font"], spec["size"], nombre, resultado
            )

    if rfc_cliente:
        _insertar_rfc_en_casillas(page, fieldmap["campo_rfc_cliente"], rfc_cliente, resultado)

    for nombre_cb in checkboxes_activos or []:
        checkbox = fieldmap["checkboxes"].get(nombre_cb)
        if checkbox is None:
            raise ValueError(
                f"El checkbox {nombre_cb!r} no está calibrado en "
                f"contrato_fieldmap.v1.json; no se puede marcar sin "
                f"coordenadas conocidas."
            )
        _insertar_checkbox(page, checkbox)

    Path(salida_pdf).parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(salida_pdf))
    doc.close()
    return resultado
