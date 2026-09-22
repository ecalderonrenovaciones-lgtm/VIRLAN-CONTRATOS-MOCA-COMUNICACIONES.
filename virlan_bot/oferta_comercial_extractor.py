"""Extrae datos del PDF "Formato de Autorizaciones Especiales" (oferta /
propuesta comercial) que llega adjunto al correo cuando el cliente tiene un
descuento o condición especial negociada aparte del catálogo general.

Descubierto 2026-09-22 con un caso real (CORPORATIVO EN FARMACIAS Y
SERVICIOS ZH, correo con T1 y T2): cuando este PDF existe, es la fuente MÁS
autorizada para varios campos que antes se dejaban en blanco o se
aproximaban con reglas genéricas:

  - Folio de Autorización (OP celda G12) = "FOLIO DE PRODUCTO" del PDF.
  - Meses Gratis (3 casillas/shapes de la OP, J12/K12/L12) = los 3 números
    de "MESES DE RENTA GRATIS" (en los 2 casos reales vistos hasta ahora
    siempre 7, 13, 19 — parece ser una promoción estándar, no específica
    del cliente, pero se lee del PDF de todas formas en vez de asumir).
  - Por línea, dentro de cada renglón de "OFERTA COMERCIAL":
      "Descuento multilínea unitario" (OP col O) = "<dmr>% DMR".
      "Pago inicial/Diferencial de equipo unitario" (OP col K) = el monto
      literal "$<precio_unitario_sin_iva> sin IVA c/u" — se usa tal cual,
      sin recalcular, es el precio ya negociado.

Verificado contra el mismo caso real: 'Costo de Equipo' de control de
renovación coincidió EXACTO ($23,447.41 y $1,300.75) con el "$... sin IVA
c/u" de este PDF para las 14 líneas — confirma que control de renovación
se llena a partir de este mismo documento cuando existe, y por eso puede
usarse de respaldo cuando no hay PDF de oferta comercial adjunto.

Formato de cada renglón de movimiento (una fila por grupo de líneas que
comparten plan+equipo+descuento+plazo), tal como lo extrae PyMuPDF:
  "<cantidad> ATT Armalo Negocios <plan> [CPP] [CTRL] // <equipo> Desc
  <desc_equipo>% ($<precio_sin_iva> sin IVA c/u) // <dmr>% DMR // <plazo>
  meses // ..."
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field as dc_field
from pathlib import Path

import fitz

from .equipo_matching import equipo_coincide

_PATRON_FOLIO = re.compile(r"\b(\d{6}[A-Z]{2,3}\d{4,8})\b")
_PATRON_MESES_GRATIS = re.compile(
    r"MESES DE RENTA GRATIS\s*\n?\s*(\d+)\s*\n?\s*(\d+)\s*\n?\s*(\d+)"
)
_PATRON_MOVIMIENTO = re.compile(
    r"(\d+)\s+ATT Armalo Negocios\s+(\d+)\s*(?:CPP)?\s*(?:CTRL)?\s*//\s*"
    r"(.+?)\s+Desc\s+(\d+)%\s*\(\$([\d,]+\.\d+)\s*sin IVA c/u\)\s*//\s*"
    r"(\d+)%\s*DMR\s*//\s*(\d+)\s*meses",
    re.IGNORECASE,
)


@dataclass
class MovimientoOferta:
    cantidad_lineas: int
    numero_plan: float
    equipo: str
    descuento_equipo_pct: float
    precio_unitario_sin_iva: float
    dmr_pct: float
    plazo_meses: int


@dataclass
class OfertaComercial:
    folio_producto: str | None = None
    meses_renta_gratis: list[int] = dc_field(default_factory=list)
    movimientos: list[MovimientoOferta] = dc_field(default_factory=list)


def extraer_oferta_comercial(pdf_path: str | Path) -> OfertaComercial:
    doc = fitz.open(pdf_path)
    try:
        texto = "\n".join(page.get_text() for page in doc)
    finally:
        doc.close()

    resultado = OfertaComercial()

    m_folio = _PATRON_FOLIO.search(texto)
    if m_folio:
        resultado.folio_producto = m_folio.group(1)

    m_meses = _PATRON_MESES_GRATIS.search(texto)
    if m_meses:
        resultado.meses_renta_gratis = [int(m_meses.group(i)) for i in (1, 2, 3)]

    texto_normalizado = re.sub(r"\s+", " ", texto)
    for m in _PATRON_MOVIMIENTO.finditer(texto_normalizado):
        resultado.movimientos.append(
            MovimientoOferta(
                cantidad_lineas=int(m.group(1)),
                numero_plan=float(m.group(2)),
                equipo=m.group(3).strip(),
                descuento_equipo_pct=float(m.group(4)),
                precio_unitario_sin_iva=float(m.group(5).replace(",", "")),
                dmr_pct=float(m.group(6)),
                plazo_meses=int(m.group(7)),
            )
        )

    return resultado


def _tokens(texto: str) -> Counter:
    return Counter((texto or "").strip().upper().split())


def buscar_movimiento_para_linea(
    modelo: str, numero_plan: float | None, movimientos: list[MovimientoOferta]
) -> MovimientoOferta | None:
    """Encuentra el movimiento de la oferta comercial que corresponde a una
    línea, cruzando por número de plan y equipo (mismo matching por
    subconjunto de tokens que catalogo_precios/calculadora_mpe, en ambas
    direcciones porque no se sabe de antemano cuál de las dos cadenas trae
    la marca). Si hay 0 o más de 1 coincidencia, devuelve None (ambiguo o
    sin match) para que el llamador use un respaldo y alerte."""
    if numero_plan is None:
        return None

    tokens_modelo = _tokens(modelo)
    candidatos = [
        mov
        for mov in movimientos
        if mov.numero_plan == numero_plan
        and (
            equipo_coincide(tokens_modelo, mov.equipo.strip().upper())
            or equipo_coincide(_tokens(mov.equipo), (modelo or "").strip().upper())
        )
    ]
    if len(candidatos) == 1:
        return candidatos[0]
    return None
