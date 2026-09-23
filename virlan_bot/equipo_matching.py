"""Búsqueda de un equipo (nombre que trae SAE) en un catálogo (lista de
precios o cálculo de MPE).

REGLA DEL USUARIO (2026-09-23): siempre debe ser el MISMO modelo; jamás se usa
el precio de un equipo distinto (esto reemplaza la regla del 2026-09-22 que
permitía tomar el modelo vecino, ej. A14 -> A13). Los nombres cambian entre
archivos, pero dos o más palabras parecidas bastan para inferir que es el mismo
equipo. Si no hay ningún candidato que sea el mismo modelo, el dato queda en N/A
y se consulta el porqué: `mejor_sustituto` solo se usa para SUGERIR en el aviso
cuál es el más parecido (nunca para llenar valores).

Se considera el mismo equipo: subconjunto de palabras (la marca puede faltar o
venir en otro orden), desempate por el más exacto, tolerancia a que el catálogo
omita la conectividad o la capacidad, mismo nombre de modelo aunque cambie el
espaciado ('ZFOLD 8 ULTRA' = 'Z FOLD8 ULTRA') y las equivalencias confirmadas
por el usuario en ALIAS_EQUIPOS.
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
_RE_TOKEN_GB = re.compile(r"\d+GB", re.IGNORECASE)

_MARCAS = {
    "SAMSUNG", "APPLE", "HONOR", "MOTOROLA", "XIAOMI", "OPPO", "ZTE", "HUAWEI",
    "REALME", "NOKIA", "ALCATEL", "TCL", "LG", "SONY", "TECNO", "INFINIX", "VIVO",
}


def _clave_modelo(nombre_norm: str) -> str:
    """Nombre del modelo sin marca, capacidad ni conectividad y sin espacios:
    'SAMSUNG GALAXY ZFOLD 8 ULTRA 512GB 5G' y 'SAMSUNG GALAXY Z FOLD8 ULTRA'
    dan ambos 'GALAXYZFOLD8ULTRA' (mismo modelo escrito con otro espaciado)."""
    tokens = [
        t
        for t in nombre_norm.split()
        if t not in _MARCAS and not _RE_TOKEN_GB.fullmatch(t) and not _RE_CONECTIVIDAD.fullmatch(t)
    ]
    return "".join(tokens)


def equipo_coincide(equipo_tokens: Counter, candidato_norm: str) -> bool:
    """True si todas las palabras de 'equipo' aparecen en 'candidato' (sin
    importar el orden ni si 'candidato' trae palabras extra, típicamente la
    marca: 'GALAXY A14 4G 128GB' coincide con 'SAMSUNG GALAXY A14 128GB 4G')."""
    candidato_tokens = Counter(candidato_norm.split())
    return all(candidato_tokens[t] >= c for t, c in equipo_tokens.items())


# Equivalencias confirmadas por el usuario entre el nombre que trae SAE y el
# nombre en los catálogos (lista de precios / Master MPE), cuando no se pueden
# deducir por tokens. Clave: nombre SAE normalizado (sin color, espacios
# colapsados); valor: nombres a probar, en orden (cada catálogo lo escribe
# distinto).
ALIAS_EQUIPOS: dict[str, tuple[str, ...]] = {
    # Confirmado 2026-09-23 (cliente ENERGY SOLUTIONS): el diferencial de la
    # lista (plan 799/24m = $2,869) coincide con el de control de renovación.
    "HONOR MAGIC8 LITE BDLCHOICEH5G": ("HONOR MAGIC 8 LITE BUNDLE", "HONOR MAGIC8 LITE BUNDLE"),
}


def buscar_coincidencias(equipo_norm: str, candidatos: list[tuple[str, T]]) -> list[T]:
    """Igual que `_buscar_coincidencias_base`, pero si no hay resultado y el
    equipo tiene un alias confirmado (`ALIAS_EQUIPOS`), prueba con sus nombres."""
    encontrados = _buscar_coincidencias_base(equipo_norm, candidatos)
    if encontrados:
        return encontrados
    for alias in ALIAS_EQUIPOS.get(" ".join(equipo_norm.split()), ()):
        encontrados = _buscar_coincidencias_base(alias, candidatos)
        if encontrados:
            return encontrados
    return []


def _buscar_coincidencias_base(equipo_norm: str, candidatos: list[tuple[str, T]]) -> list[T]:
    """Devuelve los payloads de `candidatos` [(nombre_normalizado, payload)] que
    corresponden al equipo, en dos niveles (se queda con el primero que dé
    resultados):

      1. Subconjunto de tokens (`equipo_coincide`). Si hay más de uno se
         desempata por el más exacto: primero tokens idénticos, luego
         idénticos ignorando la conectividad ('APPLE IPHONE 17 PRO 256GB' de
         SAE coincide con '... 17 PRO 256GB 5G' y con '... 17 PRO MAX 256GB
         5G', pero solo el primero es idéntico salvo el '5G').
      2. Si no hubo ninguno: igual, pero ignorando la conectividad (4G/5G/LTE)
         y/o la capacidad (GB) del equipo cuando ESE candidato no trae ninguna
         en su nombre (la hoja 'Master' de MPE llama 'APPLE IPHONE 17 256GB' a
         lo que SAE llama 'APPLE IPHONE 17 256GB 5G' y 'Honor X5d' a 'HONOR X5D
         128GB 4G'). Ausencia no contradice; una diferencia sí.

    Más de un resultado sigue significando ambigüedad (el llamador alerta)."""
    tokens = Counter(equipo_norm.split())

    def _sin_conectividad(c: Counter) -> Counter:
        return Counter({t: n for t, n in c.items() if not _RE_CONECTIVIDAD.fullmatch(t)})

    hallados = [
        (nombre, payload)
        for nombre, payload in candidatos
        if equipo_coincide(tokens, nombre)
    ]
    if len(hallados) > 1:
        for clave in (lambda c: c, _sin_conectividad):
            exactos = [h for h in hallados if clave(Counter(h[0].split())) == clave(tokens)]
            if exactos:
                hallados = exactos
                break
    if hallados:
        return [p for _, p in hallados]

    relajados = []
    for nombre, payload in candidatos:
        requeridos = Counter(
            {
                t: n
                for t, n in tokens.items()
                if not (extraer_conectividad(nombre) is None and _RE_CONECTIVIDAD.fullmatch(t))
                and not (extraer_gb(nombre) is None and _RE_TOKEN_GB.fullmatch(t))
            }
        )
        if requeridos != tokens and requeridos and equipo_coincide(requeridos, nombre):
            relajados.append((nombre, payload, requeridos))
    if len(relajados) > 1:
        exactos = [r for r in relajados if Counter(r[0].split()) == r[2]]
        if exactos:
            relajados = exactos
    if relajados:
        return [p for _, p, _ in relajados]

    # Nivel 3: mismo nombre de modelo aunque cambie el espaciado entre palabras
    # ('ZFOLD 8 ULTRA' vs 'Z FOLD8 ULTRA'). Exige igualdad completa del nombre
    # del modelo, así que 'Z FOLD8' NO empata con 'Z FOLD8 ULTRA'.
    clave = _clave_modelo(equipo_norm)
    if len(clave) >= 4:
        return [p for nombre, p in candidatos if _clave_modelo(nombre) == clave]
    return []


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
    """SOLO SUGERENCIA para el aviso (regla del usuario 2026-09-23: jamás se usa
    el precio de un equipo distinto). `candidatos` es una lista de
    (nombre_tal_como_aparece_en_catalogo, payload). Devuelve el candidato más
    cercano con la misma capacidad en GB y conectividad, o None si `equipo` no
    trae ninguna de las dos o ningún candidato coincide en ambas."""
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


def sugerencia_para_aviso(equipo: str, candidatos: list[tuple[str, T]]) -> str:
    """Texto para el aviso cuando NO hay el mismo modelo: nombres parecidos y el
    candidato más cercano por capacidad/conectividad, ambos marcados como NO
    usados."""
    parecidos = nombres_parecidos(equipo, [nombre for nombre, _ in candidatos])
    sust = mejor_sustituto(equipo, candidatos)
    extra = f"; el más cercano por capacidad y conectividad: {sust.familia_original!r}" if sust else ""
    return f"nombres parecidos (NO usados): {parecidos}{extra}"


def nombres_parecidos(equipo: str, nombres: list[str], n: int = 3) -> list[str]:
    """Nombres del catálogo más parecidos a `equipo`, SOLO como pista para que
    una persona consulte por qué no hubo coincidencia. Nunca se usan para llenar
    datos (regla del usuario 2026-09-23)."""
    equipo_norm = equipo.strip().upper()
    puntuados = sorted(
        {n_.strip() for n_ in nombres},
        key=lambda x: -difflib.SequenceMatcher(None, equipo_norm, x.upper()).ratio(),
    )
    return puntuados[:n]
