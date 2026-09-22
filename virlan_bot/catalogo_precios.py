"""Busca el 'Precio de lista' de un equipo en la hoja 'Precios' del archivo
de lista de precios que llega periódicamente por correo (ej.
'Lista de Precios Septiembre 2026_ArmaloNegocios_W38.xlsx').

Descubierto revisando las fórmulas reales del 'Cotizador' con Excel (COM):
Cotizador!C10 ("PRECIO LISTA") es literalmente
`=VLOOKUP(equipo, Precios!E:F, 2, 0)` — es decir, el Cotizador es solo una
interfaz sobre esta tabla plana, así que no hace falta simular sus menús
desplegables: se puede leer la tabla directamente.

Estructura verificada de la hoja 'Precios' (fila 3 = encabezados reales,
datos desde la fila 4):
  C Marca   D Modelo   E Familia (texto completo, es la llave de búsqueda)
  F Full price (= 'Precio de lista')
  AP Fecha Inicio   AQ Fecha Fin  (vigencia; solo es válida la fila cuya
  vigencia cubre la fecha de hoy)

Verificado con un caso real: para 'APPLE IPHONE 18 PRO MAX 512GB 5G' el
'Full price' (36999) y el precio a 24 meses del plan 799 (32859) coinciden
exactamente con los valores ya usados por el equipo de VIRLAN en
'control de renovacion.xlsx' para ese mismo cliente.

La columna 'Modelo' de SAE 2.2.xlsx no siempre incluye la marca (ej. equipos
Samsung vienen como 'GALAXY A14 4G 128GB'), mientras que 'Familia' en la
lista de precios sí la incluye ('SAMSUNG GALAXY A14 4G 128GB'). Por eso el
match acepta también que 'Familia' termine en " " + equipo (con límite de
palabra), no solo la igualdad exacta.

Si no hay coincidencia exacta (ej. equipo descontinuado, ya no aparece en la
lista vigente), se busca un sustituto con equipo_matching.mejor_sustituto:
misma capacidad en GB y misma conectividad (4G/5G/LTE) son obligatorias, el
resto es "el más similar". Regla confirmada por el usuario — ver
equipo_matching.py. Un resultado sustituto siempre trae `es_sustituto=True`
y debe alertarse para revisión manual antes de enviar.
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from collections import Counter

from .equipo_matching import equipo_coincide, mejor_sustituto

_HOJA = "Precios"
_FILA_ENCABEZADOS = 3
_COL_FAMILIA = 5  # E
_COL_FULL_PRICE = 6  # F
_COL_FECHA_INICIO = 42  # AP
_COL_FECHA_FIN = 43  # AQ


class EquipoNoEncontradoError(Exception):
    pass


@dataclass
class PrecioEquipo:
    familia: str
    precio_lista: float
    fila: int
    es_sustituto: bool = False


def buscar_precio_lista(
    equipo: str, lista_precios_xlsx: str | Path, fecha: _dt.date | None = None
) -> PrecioEquipo:
    fecha = fecha or _dt.date.today()
    equipo_norm = equipo.strip().upper()
    equipo_tokens = Counter(equipo_norm.split())

    wb = openpyxl.load_workbook(lista_precios_xlsx, data_only=True, read_only=True)
    try:
        ws = wb[_HOJA]
        coincidencias = []
        vigentes = []  # (familia, PrecioEquipo) — para fallback de sustituto
        for fila in ws.iter_rows(min_row=_FILA_ENCABEZADOS + 1):
            familia = fila[_COL_FAMILIA - 1].value
            if not familia:
                continue
            familia_norm = str(familia).strip().upper()
            inicio = fila[_COL_FECHA_INICIO - 1].value
            fin = fila[_COL_FECHA_FIN - 1].value
            if isinstance(inicio, _dt.datetime):
                inicio = inicio.date()
            if isinstance(fin, _dt.datetime):
                fin = fin.date()
            if inicio and fin and not (inicio <= fecha <= fin):
                continue
            precio = fila[_COL_FULL_PRICE - 1].value
            if precio is None:
                continue
            candidato = PrecioEquipo(str(familia), float(precio), fila[0].row)
            vigentes.append((familia_norm, candidato))
            if equipo_coincide(equipo_tokens, familia_norm):
                coincidencias.append(candidato)
    finally:
        wb.close()

    if len(coincidencias) > 1:
        raise EquipoNoEncontradoError(
            f"'{equipo}' tiene {len(coincidencias)} filas vigentes distintas "
            f"al {fecha} en '{Path(lista_precios_xlsx).name}' (filas "
            f"{[c.fila for c in coincidencias]}); revisar manualmente."
        )
    if coincidencias:
        return coincidencias[0]

    sustituto = mejor_sustituto(equipo_norm, vigentes)
    if sustituto is not None:
        resultado = sustituto.payload
        resultado.es_sustituto = True
        return resultado

    raise EquipoNoEncontradoError(
        f"No se encontró '{equipo}' vigente al {fecha} en la hoja "
        f"'{_HOJA}' de '{Path(lista_precios_xlsx).name}', ni un sustituto "
        f"con la misma capacidad y conectividad."
    )
