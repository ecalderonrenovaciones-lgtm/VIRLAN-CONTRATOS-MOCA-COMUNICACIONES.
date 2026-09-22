"""Búsqueda de un equipo "similar" cuando no hay coincidencia exacta en un
catálogo (lista de precios o cálculo de MPE).

Regla confirmada por el usuario (2026-09-22, cliente CORPORATIVO EN
FARMACIAS Y SERVICIOS ZH, equipo 'GALAXY A14 4G 128GB' descontinuado de la
lista de precios de septiembre 2026): quien revisa a veces encuentra que el
equipo exacto ya no está en el catálogo vigente. En ese caso se debe usar el
equipo más similar, con la restricción de que SIEMPRE deben coincidir:

  - La capacidad en GB.
  - La conectividad (4G / 5G, rara vez LTE).

Entre los candidatos que cumplen esa restricción, se prioriza el que tenga
el número de modelo numéricamente más cercano (ej. A14 -> A13 antes que
A16), y como desempate adicional la mayor similitud de texto completo.

Esto es una sustitución, no el equipo real vendido: siempre se debe generar
una alerta explícita para que se revise manualmente antes de enviar.
"""

from __future__ import annotations

import difflib
import re
from collections import Counter
from dataclasses import dataclass
from typing import Generic, TypeVar

T = TypeVar("T")

_RE_GB = re.compile(r"(\d+)\s*GB", re.IGNORECASE)
_RE_CONECTIVIDAD = re.compile(r"\b(5G|4G|LTE)\b", re.IGNORECASE)


def equipo_coincide(equipo_tokens: Counter, candidato_norm: str) -> bool:
    """True si todas las palabras de 'equipo' aparecen en 'candidato' (sin
    importar el orden ni si 'candidato' trae palabras extra, típicamente la
    marca: 'GALAXY A14 4G 128GB' coincide con 'SAMSUNG GALAXY A14 128GB 4G')."""
    candidato_tokens = Counter(candidato_norm.split())
    return all(candidato_tokens[t] >= c for t, c in equipo_tokens.items())


def extraer_gb(texto: str) -> int | None:
    m = _RE_GB.search(texto)
    return int(m.group(1)) if m else None


def extraer_conectividad(texto: str) -> str | None:
    m = _RE_CONECTIVIDAD.search(texto)
    return m.group(1).upper() if m else None


def _extraer_numero_modelo(texto: str, gb: int | None, conectividad: str | None) -> int | None:
    limpio = texto
    if gb is not None:
        limpio = re.sub(rf"{gb}\s*GB", " ", limpio, flags=re.IGNORECASE)
    if conectividad:
        limpio = re.sub(rf"\b{conectividad}\b", " ", limpio, flags=re.IGNORECASE)
    numeros = re.findall(r"\d+", limpio)
    return int(numeros[0]) if numeros else None


@dataclass
class Sustituto(Generic[T]):
    familia_original: str
    payload: T


def mejor_sustituto(
    equipo: str, candidatos: list[tuple[str, T]]
) -> Sustituto[T] | None:
    """`candidatos` es una lista de (familia_tal_como_aparece_en_catalogo, payload).
    Devuelve el mejor sustituto cuya capacidad en GB y conectividad coincidan
    exactamente con las de `equipo`, o None si `equipo` no trae capacidad/
    conectividad identificable, o si ningún candidato coincide en ambas."""
    equipo_norm = equipo.strip().upper()
    gb = extraer_gb(equipo_norm)
    conectividad = extraer_conectividad(equipo_norm)
    if gb is None or conectividad is None:
        return None
    numero_modelo = _extraer_numero_modelo(equipo_norm, gb, conectividad)

    elegibles = []
    for familia, payload in candidatos:
        familia_norm = familia.strip().upper()
        if extraer_gb(familia_norm) != gb or extraer_conectividad(familia_norm) != conectividad:
            continue
        distancia_modelo = None
        if numero_modelo is not None:
            n = _extraer_numero_modelo(familia_norm, gb, conectividad)
            if n is not None:
                distancia_modelo = abs(n - numero_modelo)
        similitud = difflib.SequenceMatcher(None, equipo_norm, familia_norm).ratio()
        # None (sin número de modelo comparable) ordena al final.
        clave_distancia = distancia_modelo if distancia_modelo is not None else 10**9
        elegibles.append((clave_distancia, -similitud, familia, payload))

    if not elegibles:
        return None

    elegibles.sort(key=lambda t: (t[0], t[1]))
    _, _, familia_elegida, payload_elegido = elegibles[0]
    return Sustituto(familia_original=familia_elegida, payload=payload_elegido)
