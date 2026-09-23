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

Si no hay el MISMO modelo NO se usa el precio de otro equipo (regla del usuario
2026-09-23): se lanza EquipoNoEncontradoError con el motivo y una sugerencia
(que no se usa) para consultar.
"""

from __future__ import annotations

import math
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .equipo_matching import buscar_coincidencias, sugerencia_para_aviso

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

    wb = openpyxl.load_workbook(calculo_mpe_xlsx, data_only=True, read_only=True)
    try:
        ws = wb[_HOJA]
        candidatos = []  # (modelo_norm, (precio_base, fila))
        for fila in ws.iter_rows(min_row=2):
            modelo = fila[_COL_MODELO - 1].value
            if not modelo:
                continue
            modelo_norm = str(modelo).strip().upper()
            precio_base = fila[_COL_PRECIO_BASE - 1].value
            if precio_base is None:
                continue
            candidatos.append((modelo_norm, (float(precio_base), fila[0].row)))
    finally:
        wb.close()

    coincidencias = buscar_coincidencias(equipo_norm, candidatos)

    if len(coincidencias) > 1:
        raise EquipoNoEncontradoError(
            f"'{equipo}' tiene {len(coincidencias)} filas distintas en "
            f"'{Path(calculo_mpe_xlsx).name}' (filas "
            f"{[f for _, f in coincidencias]}); revisar manualmente."
        )

    if not coincidencias:
        raise EquipoNoEncontradoError(
            f"No se encontró el MISMO modelo de '{equipo}' en la hoja '{_HOJA}' de "
            f"'{Path(calculo_mpe_xlsx).name}'; no se usa el precio de un equipo "
            f"distinto, consultar por qué no aparece "
            f"({sugerencia_para_aviso(equipo, candidatos)})."
        )
    precio_base, fila_num = coincidencias[0]

    if not plazo_meses:
        raise ValueError("plazo_meses debe ser mayor a 0 para calcular el MPE.")

    mpe = _truncar((precio_base - pie) / int(plazo_meses) + 0.01, 2)
    return ResultadoMPE(precio_base=precio_base, mpe=mpe, fila=fila_num)


def mpe_desde_lista(precio_lista: float, pie: float, plazo_meses: int) -> float:
    """Criterio alternativo del MPE (regla del usuario 2026-09-23): (Precio de
    lista - Diferencial de equipo unitario) / plazo contratado, con la misma
    fórmula de la calculadora (TRUNC(x + 0.01, 2)). Verificado: coincide con la
    calculadora en 37 de 38 líneas ya generadas (la excepción fue un equipo cuyo
    precio en el Master difería del de la lista)."""
    return _truncar((precio_lista - pie) / int(plazo_meses) + 0.01, 2)
