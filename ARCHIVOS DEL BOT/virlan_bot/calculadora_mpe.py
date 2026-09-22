"""Calcula el MPE (Mensualidad de Pago por Equipo) reproduciendo la fórmula
real de 'Calculadora MPE' (hoja 'Calculadora MPE' en
'140926 Cálculo de MPE VA.xlsx'), verificada con Excel (COM):

  MPE = TRUNC((Precio_Base - PIE) / Plazo_meses + 0.01, 2)

donde Precio_Base sale de la hoja 'Master' (columna 'Precio Base', fila
donde 'Modelo' coincide con el equipo) y PIE ('Pago Inicial de Equipo') es
un anticipo que en los ejemplos disponibles siempre fue 0 — se deja como
parámetro con ese valor por defecto hasta confirmar si alguna vez es
distinto de 0.

Estructura verificada de la hoja 'Master' (fila 1 encabezados):
  A Marca   B Modelo (llave de búsqueda)   C Precio inicial   D Precio Base
  E Baja de Precio   F Fecha Baja Precio

Si no hay coincidencia exacta, se busca un sustituto con
equipo_matching.mejor_sustituto (misma capacidad en GB y misma conectividad
obligatorias) — ver catalogo_precios.py para el mismo mecanismo aplicado a
la lista de precios. El resultado trae `es_sustituto=True` y debe alertarse.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from collections import Counter

from .equipo_matching import equipo_coincide, mejor_sustituto

_HOJA = "Master"
_COL_MODELO = 2  # B
_COL_PRECIO_BASE = 4  # D


class EquipoNoEncontradoError(Exception):
    pass


@dataclass
class ResultadoMPE:
    precio_base: float
    mpe: float
    fila: int
    es_sustituto: bool = False


def _truncar(valor: float, decimales: int = 2) -> float:
    factor = 10**decimales
    return math.trunc(valor * factor) / factor


def calcular_mpe(
    equipo: str,
    plazo_meses: int,
    calculo_mpe_xlsx: str | Path,
    pie: float = 0.0,
) -> ResultadoMPE:
    equipo_norm = equipo.strip().upper()
    equipo_tokens = Counter(equipo_norm.split())

    wb = openpyxl.load_workbook(calculo_mpe_xlsx, data_only=True, read_only=True)
    try:
        ws = wb[_HOJA]
        coincidencias = []
        candidatos = []  # (modelo_norm, (precio_base, fila)) — fallback de sustituto
        for fila in ws.iter_rows(min_row=2):
            modelo = fila[_COL_MODELO - 1].value
            if not modelo:
                continue
            modelo_norm = str(modelo).strip().upper()
            precio_base = fila[_COL_PRECIO_BASE - 1].value
            if precio_base is None:
                continue
            candidatos.append((modelo_norm, (float(precio_base), fila[0].row)))
            if equipo_coincide(equipo_tokens, modelo_norm):
                coincidencias.append((float(precio_base), fila[0].row))
    finally:
        wb.close()

    if len(coincidencias) > 1:
        raise EquipoNoEncontradoError(
            f"'{equipo}' tiene {len(coincidencias)} filas distintas en "
            f"'{Path(calculo_mpe_xlsx).name}' (filas "
            f"{[f for _, f in coincidencias]}); revisar manualmente."
        )

    es_sustituto = False
    if coincidencias:
        precio_base, fila_num = coincidencias[0]
    else:
        sustituto = mejor_sustituto(equipo_norm, candidatos)
        if sustituto is None:
            raise EquipoNoEncontradoError(
                f"No se encontró '{equipo}' en la hoja '{_HOJA}' de "
                f"'{Path(calculo_mpe_xlsx).name}', ni un sustituto con la "
                f"misma capacidad y conectividad."
            )
        precio_base, fila_num = sustituto.payload
        es_sustituto = True

    if not plazo_meses:
        raise ValueError("plazo_meses debe ser mayor a 0 para calcular el MPE.")

    mpe = _truncar((precio_base - pie) / int(plazo_meses) + 0.01, 2)
    return ResultadoMPE(precio_base=precio_base, mpe=mpe, fila=fila_num, es_sustituto=es_sustituto)
