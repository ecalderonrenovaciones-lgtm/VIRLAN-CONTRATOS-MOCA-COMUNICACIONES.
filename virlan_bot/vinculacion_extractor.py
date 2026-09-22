"""Extrae el RFC y el tipo/número de identificación oficial desde el
adjunto 'LAYOUT DE VINCULACION EMPRESARIAL CURP.xlsx', buscando por número
de cuenta. Esta es la fuente oficial de estos dos datos (confirmado por el
usuario) — no la ficha .docx ni los documentos de identidad escaneados.

Formato verificado (fila 1 encabezados, datos desde fila 2):
  Consecutivo, ATTUID analista SAE, Tipo Movimiento, # cuenta AMDOCS, RFC,
  CURP, Tipo identificación, # de identificación, Fecha de Validacion,
  Observaciones
"""

from __future__ import annotations

from pathlib import Path

from ._xlsx_utils import fila_a_dict_por_encabezado, leer_filas
from .models import ClienteContrato

_CUENTA_KEY = "# cuenta AMDOCS"


def extraer_vinculacion(cliente: ClienteContrato, layout_path: str | Path) -> None:
    """Rellena cliente.rfc, tipo_identificacion y numero_identificacion.
    Modifica el cliente en sitio."""
    layout_path = Path(layout_path)
    filas = leer_filas(layout_path)
    nombradas = fila_a_dict_por_encabezado(filas, fila_encabezado=1)

    encontrada = None
    for fila in nombradas.values():
        if fila.get(_CUENTA_KEY) == cliente.numero_cuenta:
            encontrada = fila
            break

    if encontrada is None:
        cliente.agregar_alerta(
            f"No se encontró la cuenta {cliente.numero_cuenta!r} en "
            f"'{layout_path.name}'. El RFC y la identificación oficial deben "
            f"capturarse manualmente antes de generar el CONTRATO (el RFC "
            f"siempre debe llenarse)."
        )
        cliente.campos_faltantes.extend(["rfc", "tipo_identificacion", "numero_identificacion"])
        return

    rfc = encontrada.get("RFC", "").strip()
    if rfc and rfc != cliente.rfc:
        if cliente.rfc:
            cliente.agregar_alerta(
                f"El RFC de la ficha ({cliente.rfc!r}) no coincide con el de "
                f"LAYOUT DE VINCULACION ({rfc!r}); se usará este último por "
                f"ser la fuente oficial."
            )
        cliente.rfc = rfc
        cliente.origen["rfc"] = layout_path.name
    elif not rfc:
        cliente.agregar_alerta(
            f"LAYOUT DE VINCULACION no trae RFC para la cuenta "
            f"{cliente.numero_cuenta}; el RFC siempre debe llenarse, revisar "
            f"manualmente."
        )
        cliente.campos_faltantes.append("rfc")

    cliente.tipo_identificacion = encontrada.get("Tipo identificación", "").strip()
    cliente.numero_identificacion = encontrada.get("# de identificación", "").strip()
    if cliente.tipo_identificacion:
        cliente.origen["tipo_identificacion"] = layout_path.name
    else:
        cliente.campos_faltantes.append("tipo_identificacion")
    if cliente.numero_identificacion:
        cliente.origen["numero_identificacion"] = layout_path.name
    else:
        cliente.campos_faltantes.append("numero_identificacion")
