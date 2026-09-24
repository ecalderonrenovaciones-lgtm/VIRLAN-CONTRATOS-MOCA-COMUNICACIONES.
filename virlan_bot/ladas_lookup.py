"""Determina la 'Ciudad DN' de la OP a partir de la lada del teléfono de la
línea, usando la tabla de referencia 'ladas_mexico (2).csv'.

Regla confirmada por el usuario: se toma la lada del número de teléfono y se
usa solo la PRIMERA PALABRA de la columna Ciudad_Region_Principal.

Las ladas mexicanas tienen 2 o 3 dígitos, así que se prueba primero el
prefijo de 3 dígitos y, si no hay coincidencia, el de 2 dígitos.
"""

from __future__ import annotations

import csv
import re
import unicodedata
from functools import lru_cache
from pathlib import Path


class LadaNoEncontradaError(Exception):
    pass


@lru_cache(maxsize=8)
def _cargar_tabla(csv_path: str) -> dict[str, list[tuple[str, str]]]:
    tabla: dict[str, list[tuple[str, str]]] = {}
    with open(csv_path, encoding="utf-8") as f:
        for fila in csv.DictReader(f):
            lada = fila["Lada"].strip()
            ciudad = fila["Ciudad_Region_Principal"].strip()
            estado = (fila.get("Estado") or "").strip()
            tabla.setdefault(lada, []).append((ciudad, estado))
    return tabla


def primera_palabra_ciudad(ciudad_region: str) -> str:
    primera = re.split(r"[\s(,]", ciudad_region.strip(), maxsplit=1)[0]
    return primera


def ciudad_dn_de_region(ciudad_region: str, estado: str = "") -> str:
    """Ciudad DN a partir de Ciudad_Region_Principal. Regla del usuario
    2026-09-23: 'Ciudad de Mexico' se abrevia CDMX y 'Estado de Mexico' EDOMEX
    (antes salía solo 'Ciudad'); las ladas 55 y 56 son ambas CDMX. El resto
    sigue la regla de la primera palabra."""
    norm = "".join(
        c for c in unicodedata.normalize("NFD", ciudad_region.strip().upper())
        if unicodedata.category(c) != "Mn"
    )
    if norm.startswith("CIUDAD DE MEXICO"):
        return "CDMX"
    if norm.startswith("ESTADO DE MEXICO"):
        return "EDOMEX"
    # 'San Martin Texmelucan' no cabe en la casilla y su primera palabra ('San')
    # no dice nada: en ese caso se pone el ESTADO, siempre "Puebla" (pedido del
    # usuario 2026-09-24; la lada 223 aparece en la hoja también con Tlaxcala y el
    # usuario indicó dejarla como Puebla). Cualquier otra ciudad que no quepa en
    # la casilla reduce su propia letra (ver op_filler._set_con_ajuste).
    if norm.startswith("SAN MARTIN TEXMELUCAN"):
        return "Puebla"
    return primera_palabra_ciudad(ciudad_region)


def ciudad_dn_por_telefono(telefono: str, csv_path: str | Path) -> str:
    """Devuelve la primera palabra de la ciudad correspondiente a la lada del
    teléfono. Lanza LadaNoEncontradaError si no hay coincidencia — nunca se
    debe adivinar una ciudad."""
    digitos = re.sub(r"\D", "", telefono)
    if len(digitos) < 10:
        raise LadaNoEncontradaError(
            f"El teléfono {telefono!r} no tiene 10 dígitos, no se puede "
            f"determinar la lada."
        )

    tabla = _cargar_tabla(str(csv_path))

    for largo in (3, 2):
        prefijo = digitos[:largo]
        if prefijo in tabla:
            ciudades = {ciudad_dn_de_region(c, e) for c, e in tabla[prefijo]}
            if len(ciudades) > 1:
                raise LadaNoEncontradaError(
                    f"La lada {prefijo} tiene varias ciudades posibles en "
                    f"ladas_mexico: {sorted(ciudades)}; revisar manualmente "
                    f"el teléfono {telefono!r}."
                )
            return next(iter(ciudades))

    raise LadaNoEncontradaError(
        f"No se encontró la lada del teléfono {telefono!r} en la tabla de "
        f"ladas_mexico."
    )
