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

Si no hay el MISMO modelo (ej. equipo descontinuado) NO se usa el precio de otro
equipo (regla del usuario 2026-09-23): se lanza EquipoNoEncontradoError con el
motivo y una sugerencia (que no se usa) para consultar.

Si se pasan `numero_plan` y `plazo_meses`, además se lee de la misma fila
'Pago inicial/Diferencial de equipo' (ver `_columna_diferencial`): la hoja
'Precios' trae, a la derecha de 'Full price', un bloque de 5 columnas
(12/18/24/36/48 meses) por cada número de plan (299/399/499/599/799/1299/
1499). Esta es la fuente correcta para la OP col K — antes se usaba
'Costo de Equipo' de control de renovacion.xlsx, que resultó traer errores
de captura (corregido 2026-09-22, ver op_filler.py).
"""

from __future__ import annotations

import datetime as _dt
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from .equipo_matching import buscar_coincidencias, sugerencia_para_aviso

_HOJA = "Precios"
_FILA_ENCABEZADOS = 3
_COL_FAMILIA = 5  # E
_COL_FULL_PRICE = 6  # F
_COL_FECHA_INICIO = 42  # AP
_COL_FECHA_FIN = 43  # AQ

# 'Pago inicial/Diferencial de equipo' (OP col K): la hoja 'Precios' trae,
# a la derecha de 'Full price', 7 bloques de 5 columnas cada uno (uno por
# número de plan: 299/399/499/599/799/1299/1499 — fila 1 trae el nombre
# '´Armalo Negocios <plan>' encabezando cada bloque, fila 2 el número de
# plan tal cual), y dentro de cada bloque una columna por plazo en meses
# (12/18/24/36/48, en ese orden). Verificado con un caso real (usuario
# 2026-09-22): plan 599 a 24 meses para 'APPLE IPHONE 18 PRO MAX 256GB 5G'
# da columna X (24) = 29679, que es el valor correcto de 'Pago inicial/
# Diferencial de equipo' — antes el bot usaba 'Costo de Equipo' de control
# de renovacion.xlsx para ese campo, que puede venir con error de captura
# (para ese mismo cliente traía 26679, incorrecto).
_PLANES_DIFERENCIAL = (299, 399, 499, 599, 799, 1299, 1499)
_PLAZOS_DIFERENCIAL = (12, 18, 24, 36, 48)
_COL_PRIMER_BLOQUE_DIFERENCIAL = 7  # G: inicio del bloque del plan 299


def _columna_diferencial(numero_plan: float | None, plazo_meses: int | None) -> int | None:
    if numero_plan is None or plazo_meses is None:
        return None
    numero_plan_int = int(numero_plan)
    plazo_meses_int = int(plazo_meses)
    if numero_plan_int not in _PLANES_DIFERENCIAL or plazo_meses_int not in _PLAZOS_DIFERENCIAL:
        return None
    bloque = _PLANES_DIFERENCIAL.index(numero_plan_int)
    termino = _PLAZOS_DIFERENCIAL.index(plazo_meses_int)
    return _COL_PRIMER_BLOQUE_DIFERENCIAL + 5 * bloque + termino


class EquipoNoEncontradoError(Exception):
    pass


@dataclass
class PrecioEquipo:
    familia: str
    precio_lista: float
    fila: int
    diferencial_equipo: float | None = None


def buscar_precio_lista(
    equipo: str,
    lista_precios_xlsx: str | Path,
    fecha: _dt.date | None = None,
    numero_plan: float | None = None,
    plazo_meses: int | None = None,
) -> PrecioEquipo:
    fecha = fecha or _dt.date.today()
    equipo_norm = equipo.strip().upper()
    columna_diferencial = _columna_diferencial(numero_plan, plazo_meses)

    wb = openpyxl.load_workbook(lista_precios_xlsx, data_only=True, read_only=True)
    try:
        ws = wb[_HOJA]
        vigentes = []  # (familia_normalizada, PrecioEquipo)
        # enumerate: la 1ª celda de una fila vacía es EmptyCell y no tiene .row
        for num_fila, fila in enumerate(ws.iter_rows(min_row=_FILA_ENCABEZADOS + 1), start=_FILA_ENCABEZADOS + 1):
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
            diferencial = None
            if columna_diferencial is not None:
                valor_diferencial = fila[columna_diferencial - 1].value
                # La lista trae '-' (u otro texto) cuando el equipo no se ofrece
                # en esa combinación plan/plazo: se trata como "sin dato".
                if isinstance(valor_diferencial, (int, float)):
                    diferencial = float(valor_diferencial)
            candidato = PrecioEquipo(
                str(familia), float(precio), num_fila, diferencial_equipo=diferencial
            )
            vigentes.append((familia_norm, candidato))
    finally:
        wb.close()

    coincidencias = buscar_coincidencias(equipo_norm, vigentes)

    if len(coincidencias) > 1:
        raise EquipoNoEncontradoError(
            f"'{equipo}' tiene {len(coincidencias)} filas vigentes distintas "
            f"al {fecha} en '{Path(lista_precios_xlsx).name}' (filas "
            f"{[c.fila for c in coincidencias]}); revisar manualmente."
        )
    if coincidencias:
        return coincidencias[0]

    raise EquipoNoEncontradoError(
        f"No se encontró el MISMO modelo de '{equipo}' vigente al {fecha} en la hoja "
        f"'{_HOJA}' de '{Path(lista_precios_xlsx).name}'; no se usa el precio de un "
        f"equipo distinto, consultar por qué no aparece ({sugerencia_para_aviso(equipo, vigentes)})."
    )
