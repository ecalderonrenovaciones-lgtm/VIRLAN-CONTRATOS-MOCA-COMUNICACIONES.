"""Determina el tipo de contratación (NUEVA / ADICIÓN / RENOVACIÓN) a partir
del asunto del correo, confirmado por el usuario: 'Cuando sea diferente,
ADICIÓN O VENTA NUEVA, lo notarás en el asunto del correo; todo lo demás se
copia igual'. Si el asunto no trae ninguna palabra clave reconocible, se
asume RENOVACIÓN (el caso por defecto en los ejemplos disponibles) y se
registra una alerta para que se confirme en la revisión humana.
"""

from __future__ import annotations

import re

# Acepta plural ("ADICIONES", visto en CREA IMPRENTA 2026-09-24).
_PATRON_ADICION = re.compile(r"\bADICI[OÓ]N(?:ES)?\b", re.I)
_PATRON_NUEVA = re.compile(r"\b(SUSCRIPCI[OÓ]N\s+NUEVA|VENTA\s+NUEVA|ALTA\s+NUEVA|\bNUEVA\b)", re.I)
_PATRON_RENOVACION = re.compile(r"\bRENOVACI[OÓ]N(?:ES)?\b", re.I)

# nombre del checkbox en contrato_fieldmap.v1.json -> valor de TIPO DE VENTA en la OP
MAPA_TIPO = {
    "RENOVACION": ("tipo_contratacion_renovacion", "RENOVACION"),
    "ADICION": ("tipo_contratacion_adicion", "ADICION"),
    "NUEVA": ("tipo_contratacion_suscripcion_nueva", "NUEVA"),
}


def detectar_tipo_desde_asunto(asunto: str) -> tuple[str, list[str]]:
    """Devuelve (tipo, alertas) donde tipo es una clave de MAPA_TIPO."""
    alertas: list[str] = []

    if _PATRON_ADICION.search(asunto):
        return "ADICION", alertas
    if _PATRON_NUEVA.search(asunto):
        return "NUEVA", alertas
    if _PATRON_RENOVACION.search(asunto):
        return "RENOVACION", alertas

    alertas.append(
        f"No se encontró 'RENOVACIÓN', 'ADICIÓN' ni 'NUEVA' en el asunto del "
        f"correo ({asunto!r}); se asumió RENOVACIÓN por defecto — confirmar "
        f"manualmente el tipo de contratación correcto."
    )
    return "RENOVACION", alertas
