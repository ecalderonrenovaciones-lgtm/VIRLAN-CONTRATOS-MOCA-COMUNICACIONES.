"""Extrae los datos base del cliente desde la ficha .docx que llega como
adjunto del correo (ej. "ACABADOS TERRACOTA SA DE CV.docx").

Formato verificado: un documento Word simple (sin tablas). Los primeros 3
párrafos con texto siempre son, en orden:

    0: Razón Social
    1: Número de Cuenta
    2: Representante Legal

A partir de ahí, los campos restantes (Domicilio, RFC, Teléfono, Correo) se
identifican por contenido (no por posición fija), porque se ha visto variar
entre clientes:

- Formato "simple" (ej. ACABADOS TERRACOTA): 9 párrafos, 7 con texto, cada
  campo aparece una sola vez.
- Formato "con duplicados" (ej. CORPORATIVO EN FARMACIAS Y SERVICIOS ZH):
  trae el domicilio, el representante y el correo repetidos, y agrega un
  segundo domicilio distinto al final (domicilio de entrega de la persona
  autorizada, aún no diferenciado del domicilio fiscal — ver
  domicilio_entrega_raw en models.py).

Reglas de detección sobre los párrafos con texto a partir del índice 3:
    - Correo: contiene "@" (se admite el prefijo "Correo Electrónico : ").
    - RFC: coincide con el patrón de RFC de persona moral/física.
    - Teléfono: 10 dígitos.
    - Domicilio: 5 o más comas (8 componentes esperados, orden confirmado
      contra 3 clientes reales: calle, número, colonia, ciudad, municipio
      (repite o varía la ciudad), estado, CP, país). La PRIMERA coincidencia
      es el domicilio fiscal; una segunda coincidencia distinta se guarda
      como domicilio de entrega (informativo).
    - Duplicados exactos del representante legal se ignoran.

Si no se puede identificar alguno de los campos requeridos (domicilio, RFC,
teléfono, correo), se lanza DatosIncompletosError en vez de adivinar.
"""

from __future__ import annotations

import re
import zipfile
from pathlib import Path

from .models import ClienteContrato, DatosIncompletosError

_W_T = re.compile(r"<w:t[^>]*>([^<]*)</w:t>")
_W_P = re.compile(r"<w:p[ >].*?</w:p>", re.S)

_CAMPOS_MINIMOS = 7  # párrafos no vacíos mínimos esperados

_RE_RFC = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")
_RE_TELEFONO = re.compile(r"^\d{10}$")
_RE_CORREO_LABEL = re.compile(r"^correo\s*electr[oó]nico\s*:?\s*(.+)$", re.IGNORECASE)


def _parrafos(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    parrafos = []
    for p in _W_P.findall(xml):
        texto = "".join(_W_T.findall(p))
        parrafos.append(texto)
    return parrafos


def extraer_ficha(docx_path: str | Path) -> ClienteContrato:
    docx_path = Path(docx_path)
    parrafos = _parrafos(docx_path)
    no_vacios = [p.strip() for p in parrafos if p.strip()]

    if len(no_vacios) < _CAMPOS_MINIMOS:
        raise DatosIncompletosError(
            f"'{docx_path.name}' tiene {len(no_vacios)} párrafos con texto "
            f"(se esperaban al menos {_CAMPOS_MINIMOS}). No se puede mapear "
            f"con confianza; revisar el documento manualmente. "
            f"Párrafos: {no_vacios!r}"
        )

    razon_social, numero_cuenta, representante = no_vacios[:3]

    domicilios: list[str] = []
    rfc = None
    telefono = None
    correo = None
    sin_identificar: list[str] = []

    for p in no_vacios[3:]:
        m_label = _RE_CORREO_LABEL.match(p)
        valor = m_label.group(1).strip() if m_label else p

        if "@" in valor:
            if correo is None:
                correo = valor
            continue
        if _RE_RFC.match(p):
            if rfc is None:
                rfc = p
            continue
        if _RE_TELEFONO.match(p):
            if telefono is None:
                telefono = p
            continue
        if p == representante:
            continue  # representante repetido
        if p.count(",") >= 5:
            if p not in domicilios:
                domicilios.append(p)
            continue
        sin_identificar.append(p)

    faltantes = [
        nombre
        for nombre, valor in (
            ("domicilio", domicilios),
            ("RFC", rfc),
            ("teléfono", telefono),
            ("correo", correo),
        )
        if not valor
    ]
    if faltantes:
        raise DatosIncompletosError(
            f"No se pudieron identificar los campos {faltantes} en "
            f"'{docx_path.name}'. Párrafos: {no_vacios!r}"
        )

    if not numero_cuenta.isdigit():
        raise DatosIncompletosError(
            f"'{docx_path.name}': se esperaba que el párrafo 2 "
            f"('{numero_cuenta}') fuera el número de cuenta (solo dígitos)."
        )

    domicilio_raw = domicilios[0]

    cliente = ClienteContrato(
        razon_social=razon_social,
        numero_cuenta=numero_cuenta,
        representante_legal=representante,
        rfc=rfc,
        telefono=telefono,
        correo=correo,
    )
    cliente.origen["razon_social"] = docx_path.name
    cliente.origen["numero_cuenta"] = docx_path.name
    cliente.origen["representante_legal"] = docx_path.name
    cliente.origen["rfc_ficha"] = docx_path.name
    cliente.origen["telefono"] = docx_path.name
    cliente.origen["correo"] = docx_path.name

    componentes = [c.strip() for c in domicilio_raw.split(",")]
    if len(componentes) == 8:
        (
            cliente.domicilio_calle,
            cliente.domicilio_numero,
            cliente.domicilio_colonia,
            cliente.domicilio_ciudad,
            cliente.domicilio_municipio,
            cliente.domicilio_estado,
            cliente.domicilio_cp,
            cliente.domicilio_pais,
        ) = componentes
        cliente.origen["domicilio"] = docx_path.name
    else:
        cliente.domicilio_calle = domicilio_raw
        cliente.agregar_alerta(
            f"El domicilio tiene {len(componentes)} componentes separados por "
            f"coma (se esperaban 8); se guardó sin separar en "
            f"domicilio_calle. Revisar manualmente: {domicilio_raw!r}"
        )

    if len(domicilios) > 1:
        cliente.domicilio_entrega_raw = domicilios[1]
        cliente.agregar_alerta(
            "La ficha trae un segundo domicilio distinto al fiscal "
            f"(posible domicilio de entrega): {domicilios[1]!r}. Por ahora "
            "se usa solo el domicilio fiscal en contrato/OP; revisar si "
            "corresponde diferenciar entrega vs. fiscal."
        )

    if sin_identificar:
        cliente.agregar_alerta(
            f"Párrafos de la ficha no identificados (ignorados): {sin_identificar!r}"
        )

    if not re.search(r"@", cliente.correo):
        cliente.agregar_alerta(f"El correo extraído no parece válido: {cliente.correo!r}")

    return cliente
