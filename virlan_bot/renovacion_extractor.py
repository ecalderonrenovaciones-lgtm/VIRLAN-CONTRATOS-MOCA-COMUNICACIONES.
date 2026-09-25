"""Extrae las líneas a renovar (Plan tarifario, Plazo, Marca/Modelo/Color)
desde los adjuntos 'SAE 2.2.xlsx' y/o 'control de renovacion.xlsx' del
correo, buscando por número de cuenta.

Formato verificado (fila de encabezados real, columna 'Cuenta'/'Cuenta Sub
Cuenta' trae el número de cuenta):

  control de renovacion.xlsx (fila 7 encabezados, datos desde fila 8):
    No, EMPRESA, Cuenta/Sub cuenta, Región, Línea de Renovar, Plan, Modelo,
    Color, Plazo, Costo de Equipo

  SAE 2.2.xlsx (fila 8 encabezados, datos desde fila 9):
    No, Empresa, Cuenta / Sub Cuenta, STREAMING, REDES SOCIALES,
    PROTECCIÓN DE EQUPO, Líneas a Renovar, Fecha de Vencimiento Actual,
    Modalidad MPP/CPP, Modelo, Color, Plan, Plazo, Addon Extra

Si ambos archivos están disponibles y traen filas para la misma cuenta con
datos distintos (plan/modelo/plazo), se reporta como alerta en vez de
elegir uno arbitrariamente.
"""

from __future__ import annotations

import re

from pathlib import Path

from ._xlsx_utils import fila_a_dict_por_encabezado, leer_filas
from .models import ClienteContrato, LineaRenovacion

_CUENTA_KEYS_CONTROL = ("Cuenta/Sub cuenta",)
_CUENTA_KEYS_SAE = ("Cuenta                          Sub Cuenta", "Cuenta / Sub Cuenta")


def _solo_digitos(texto: str) -> str:
    return re.sub(r"\D", "", texto or "")


def _es_telefono(texto: str) -> bool:
    return len(_solo_digitos(texto)) == 10


def _valor_columna(fila: dict[str, str], fragmento: str) -> str:
    """Valor de la primera columna cuyo encabezado (con espacios raros
    colapsados) contiene `fragmento`: el encabezado real es 'Modalidad      MPP
    / CPP' con espacios de largo variable, y por eso la búsqueda por nombre
    exacto no lo encontraba."""
    for nombre, valor in fila.items():
        if fragmento in re.sub(r"\s+", " ", nombre or "").upper():
            return (valor or "").strip()
    return ""


def _buscar_valor(fila: dict[str, str], posibles_nombres: tuple[str, ...]) -> str | None:
    for nombre in posibles_nombres:
        if nombre in fila:
            return fila[nombre]
    return None


def _extraer_control_renovacion(path: Path, numero_cuenta: str) -> list[LineaRenovacion]:
    filas = leer_filas(path)
    nombradas = fila_a_dict_por_encabezado(filas, fila_encabezado=7)
    lineas = []
    for fila in nombradas.values():
        cuenta = _buscar_valor(fila, _CUENTA_KEYS_CONTROL)
        if cuenta != numero_cuenta:
            continue
        if not fila.get("Línea de Renovar", "").strip():
            # Fila sin teléfono (ej. fila formateada suelta al final de la hoja,
            # visto en MAQUINAS INTELIGENTES): no es una línea real.
            continue
        costo_equipo_raw = fila.get("Costo de Equipo")
        try:
            costo_equipo = float(costo_equipo_raw) if costo_equipo_raw not in (None, "") else None
        except ValueError:
            costo_equipo = None
        lineas.append(
            LineaRenovacion(
                telefono=fila.get("Línea de Renovar", ""),
                plan_tarifario=fila.get("Plan", ""),
                plazo_meses=fila.get("Plazo", ""),
                modelo=fila.get("Modelo", ""),
                marca_modelo_color=", ".join(
                    p for p in (fila.get("Modelo", ""), fila.get("Color", "")) if p and p.strip().upper() not in ("NA", "N/A")
                ),
                costo_equipo=costo_equipo,
                fuente=path.name,
            )
        )
    return lineas


def _extraer_sae(path: Path, numero_cuenta: str) -> list[LineaRenovacion]:
    filas = leer_filas(path)
    nombradas = fila_a_dict_por_encabezado(filas, fila_encabezado=8)
    lineas = []
    for fila in nombradas.values():
        cuenta = _buscar_valor(fila, _CUENTA_KEYS_SAE)
        if cuenta != numero_cuenta:
            continue
        lineas.append(
            LineaRenovacion(
                telefono=fila.get("Líneas a Renovar", ""),
                plan_tarifario=fila.get("Plan", ""),
                plazo_meses=(fila.get("Plazo", "") or "").replace("MESES", "").strip(),
                modelo=fila.get("Modelo", ""),
                marca_modelo_color=", ".join(
                    p for p in (fila.get("Modelo", ""), fila.get("Color", "")) if p and p.strip().upper() not in ("NA", "N/A")
                ),
                modalidad_mpp_cpp=_valor_columna(fila, "MPP") or None,
                addon_extra=fila.get("Addon Extra") or None,
                redes_sociales=fila.get("REDES SOCIALES") or None,
                streaming=fila.get("STREAMING") or None,
                fuente=path.name,
            )
        )
    return lineas


def extraer_lineas_renovacion(
    cliente: ClienteContrato,
    control_renovacion_path: str | Path | None = None,
    sae_path: str | Path | None = None,
) -> None:
    """Rellena cliente.lineas a partir de los archivos disponibles,
    filtrando por cliente.numero_cuenta. Modifica el cliente en sitio."""
    lineas_control = (
        _extraer_control_renovacion(Path(control_renovacion_path), cliente.numero_cuenta)
        if control_renovacion_path
        else []
    )
    lineas_sae = (
        _extraer_sae(Path(sae_path), cliente.numero_cuenta) if sae_path else []
    )

    if lineas_control and lineas_sae:
        if len(lineas_control) != len(lineas_sae):
            cliente.agregar_alerta(
                f"control de renovacion.xlsx trae {len(lineas_control)} línea(s) "
                f"y SAE trae {len(lineas_sae)} línea(s) para la cuenta "
                f"{cliente.numero_cuenta}; revisar manualmente."
            )
        # Se empareja por TELÉFONO (control y SAE pueden traer las líneas en distinto
        # orden); solo las líneas sin teléfono válido (SIM por portabilidad, 'NA') se
        # emparejan por posición. Emparejar todo por posición asignaba el costo de
        # equipo a la línea equivocada cuando el orden difería.
        por_telefono = {}
        for a in lineas_control:
            if _es_telefono(a.telefono):
                por_telefono.setdefault(_solo_digitos(a.telefono), a)
        pares = []
        for i, b in enumerate(lineas_sae):
            if _es_telefono(b.telefono):
                a = por_telefono.get(_solo_digitos(b.telefono))
                if a is None:
                    cliente.agregar_alerta(
                        f"La línea {b.telefono} de SAE no aparece en control de renovacion.xlsx; "
                        f"su 'Costo de Equipo' no se pudo determinar de ahí."
                    )
                    continue
            else:
                a = lineas_control[i] if i < len(lineas_control) else None
                if a is None:
                    continue
            pares.append((a, b))
        for a, b in pares:
            if a.plan_tarifario != b.plan_tarifario or a.plazo_meses != b.plazo_meses:
                cliente.agregar_alerta(
                    f"Discrepancia entre control de renovacion.xlsx "
                    f"(plan={a.plan_tarifario!r}, plazo={a.plazo_meses!r}) y SAE "
                    f"(plan={b.plan_tarifario!r}, plazo={b.plazo_meses!r}) para "
                    f"la línea {a.telefono or b.telefono}; se usará SAE por traer "
                    f"más columnas, pero debe revisarse."
                )
            # 'Costo de Equipo' solo viene en control de renovacion.xlsx;
            # se traspasa a la línea de SAE que se usará como principal.
            b.costo_equipo = a.costo_equipo
        cliente.lineas = lineas_sae
    elif lineas_sae:
        cliente.lineas = lineas_sae
    elif lineas_control:
        cliente.lineas = lineas_control
    else:
        cliente.agregar_alerta(
            f"No se encontraron líneas a renovar para la cuenta "
            f"{cliente.numero_cuenta} en SAE ni en control de renovación."
        )

    if not cliente.lineas:
        cliente.campos_faltantes.append("lineas")
