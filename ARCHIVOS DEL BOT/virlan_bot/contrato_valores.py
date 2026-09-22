"""Arma el diccionario de valores que necesita contrato_filler.llenar_contrato
a partir de un ClienteContrato, combinando:
  - datos extraídos del cliente (ficha, vinculación, líneas)
  - valores por defecto del vendedor/ejecutivo (config.VENDEDOR_POR_DEFECTO)
    — en los dos ejemplos reales disponibles fueron siempre los mismos y no
    vienen en ningún archivo del correo
  - la persona autorizada para recibir equipos: se recibe como parámetro
    explícito y, si no se da, se usa el representante legal del cliente
    (regla confirmada por el usuario 2026-09-22: cuando no se especifica a
    nombre de quién se recibe, va el representante legal)
"""

from __future__ import annotations

import datetime as _dt

from . import config
from .models import ClienteContrato


def _fecha_espaciada(fecha: _dt.date) -> str:
    return f"{fecha.day:02d}          {fecha.month:02d}        {fecha.year}"


def construir_valores_contrato(
    cliente: ClienteContrato,
    persona_autorizada_recibir_equipos: str | None = None,
    fecha_contratacion: _dt.date | None = None,
) -> tuple[dict[str, str], list[str]]:
    """Devuelve (valores, alertas). `valores` trae todos los campos de
    contrato_fieldmap.v1.json['campos_texto'] que se pudieron determinar;
    los que falten se omiten (contrato_filler los deja sin tocar)."""
    alertas: list[str] = []
    fecha_contratacion = fecha_contratacion or _dt.date.today()

    if not cliente.lineas:
        alertas.append("Sin líneas de renovación: no se pudo determinar el plazo (meses) del contrato.")
        plazo = ""
    else:
        plazos = {l.plazo_meses for l in cliente.lineas if l.plazo_meses}
        if len(plazos) > 1:
            alertas.append(
                f"Las líneas a renovar tienen plazos distintos ({sorted(plazos)}); "
                f"el CONTRATO solo admite un plazo único, revisar manualmente."
            )
        plazo = cliente.lineas[0].plazo_meses

    if not persona_autorizada_recibir_equipos:
        persona_autorizada_recibir_equipos = cliente.representante_legal
        alertas.append(
            "No se indicó 'persona autorizada para recibir equipos'; se usó el "
            f"representante legal ({cliente.representante_legal}) por defecto."
        )

    fecha_maxima_entrega = fecha_contratacion + _dt.timedelta(
        days=config.ENTREGA_POR_DEFECTO["dias_habiles_entrega"]
    )
    alertas.append(
        f"Fecha máxima de entrega calculada como fecha de contratación + "
        f"{config.ENTREGA_POR_DEFECTO['dias_habiles_entrega']} días naturales "
        f"(regla aproximada, confirmar con el usuario)."
    )

    valores: dict[str, str] = dict(config.VENDEDOR_POR_DEFECTO)
    valores["fecha_contratacion"] = _fecha_espaciada(fecha_contratacion)
    valores["fecha_maxima_entrega"] = _fecha_espaciada(fecha_maxima_entrega)
    valores["hora_entrega"] = config.ENTREGA_POR_DEFECTO["hora_entrega"]

    valores["numero_cuenta"] = cliente.numero_cuenta
    valores["razon_social"] = cliente.razon_social
    valores["representante_legal_1"] = cliente.representante_legal
    valores["representante_legal_2"] = cliente.representante_legal
    valores["telefono_1"] = cliente.telefono
    valores["telefono_2"] = cliente.telefono
    valores["domicilio_calle_numero"] = cliente.domicilio_calle_numero()
    valores["domicilio_colonia"] = cliente.domicilio_colonia
    valores["domicilio_ciudad_estado_cp"] = cliente.domicilio_ciudad_estado_cp()
    valores["tipo_identificacion_oficial"] = cliente.tipo_identificacion
    valores["numero_identificacion_oficial"] = cliente.numero_identificacion
    valores["correo_electronico"] = cliente.correo
    valores["plazo_meses"] = plazo

    if persona_autorizada_recibir_equipos:
        valores["persona_autorizada_recibir_equipos"] = persona_autorizada_recibir_equipos

    valores = {k: v for k, v in valores.items() if v}
    return valores, alertas
