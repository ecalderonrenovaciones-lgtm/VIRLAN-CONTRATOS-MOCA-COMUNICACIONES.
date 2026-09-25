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

Formato vigente desde 2026-09-23 (el usuario acordó con quien redacta la
ficha que a partir de ahora siempre venga así; los clientes anteriores no
lo traen y no se reprocesan):

    - Una línea "EL PAQUETE SE ENVIA A ESTE DOMICILIO:FISCAL" o
      "...:ENTREGA" indica a cuál domicilio se envía (cliente.envio_a).
    - Rótulos "DOMICILIO FISCAL" y "DOMICILIO DE ENTREGA" (a veces solo
      "ENTREGA") encabezan el bloque de cada domicilio; cada domicilio se
      asigna al bloque en que aparece, sin depender de su orden.

Si no se puede identificar el domicilio o el correo (o si el envío es a ENTREGA y
no hay domicilio de entrega), se lanza DatosIncompletosError en vez de adivinar. El
RFC (fuente oficial: LAYOUT DE VINCULACION) y el teléfono pueden faltar en la ficha:
quedan en blanco y se avisa. Los caracteres XML (&amp; &lt;...) se desescapan.
"""

from __future__ import annotations

import html
import re
import unicodedata
import zipfile
from pathlib import Path

from .models import ClienteContrato, DatosIncompletosError

_W_T = re.compile(r"<w:t[^>]*>([^<]*)</w:t>")
_W_P = re.compile(r"<w:p[ >].*?</w:p>", re.S)

_CAMPOS_MINIMOS = 7  # párrafos no vacíos mínimos esperados

_RE_RFC = re.compile(r"^[A-ZÑ&]{3,4}\d{6}[A-Z0-9]{3}$")
_RE_TELEFONO = re.compile(r"^\d{10}$")
_RE_CORREO_LABEL = re.compile(r"^correo\s*electr[oó]nico\s*:?\s*(.+)$", re.IGNORECASE)


_RE_ENVIO = re.compile(
    r"EL PAQUETE SE ENVIA A ESTE DOMICILIO\s*:?\s*(?:DOMICILIO\s+DE\s+)?(FISCAL|ENTREGA)"
)


def _sin_acentos(texto: str) -> str:
    return "".join(
        c for c in unicodedata.normalize("NFD", texto) if unicodedata.category(c) != "Mn"
    )


def _clasificar_etiqueta(p: str) -> str | None:
    """None si `p` es un dato; si es un rótulo/banner devuelve 'envio' (línea
    'EL PAQUETE SE ENVIA...'), 'fiscal' o 'entrega' (encabezado de bloque)."""
    u = _sin_acentos(p.strip().upper())
    if u.startswith("EL PAQUETE SE ENVIA"):
        return "envio"
    if u in {"DOMICILIO FISCAL", "FISCAL"}:
        return "fiscal"
    if u in {"ENTREGA", "DOMICILIO DE ENTREGA", "DOMICILIO ENTREGA"}:
        return "entrega"
    return None


# Títulos que NO deben aparecer junto al nombre del representante legal (pedido
# del usuario 2026-09-23: "solo el nombre completo"). Se quitan solo al INICIO
# y como palabra completa ('Ingrid' o 'Drake' no se tocan).
_RE_TITULOS = re.compile(
    r"^\s*(?:(?:sr|sra|srita|se[nñ]or|se[nñ]ora|lic|licda|ing|dr|dra|arq|mtro|mtra|prof|profa|c\.p)\b\.?\s*)+",
    re.IGNORECASE,
)


def quitar_titulos(nombre: str) -> str:
    """'Sr. FRANCISCO OLLIVIER ROMERO' -> 'FRANCISCO OLLIVIER ROMERO'."""
    limpio = _RE_TITULOS.sub("", nombre or "").strip()
    return limpio or (nombre or "").strip()


def _asignar_domicilio(cliente, raw: str, prefijo: str, nombre: str) -> None:
    """Separa un domicilio de 8 componentes (calle, número, colonia, ciudad,
    municipio, estado, CP, país) en los atributos `<prefijo>_<componente>`."""
    componentes = [c.strip() for c in raw.split(",")]
    if len(componentes) == 8:
        for campo, valor in zip(
            ("calle", "numero", "colonia", "ciudad", "municipio", "estado", "cp", "pais"),
            componentes,
        ):
            setattr(cliente, f"{prefijo}_{campo}", valor)
    else:
        setattr(cliente, f"{prefijo}_calle", raw)
        cliente.agregar_alerta(
            f"El domicilio {nombre} tiene {len(componentes)} componentes separados por "
            f"coma (se esperaban 8); se guardó sin separar en "
            f"{prefijo}_calle. Revisar manualmente: {raw!r}"
        )


def _parrafos(docx_path: Path) -> list[str]:
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml").decode("utf-8")
    parrafos = []
    for p in _W_P.findall(xml):
        texto = html.unescape("".join(_W_T.findall(p)))   # Word escribe & como &amp;
        parrafos.append(texto)
    return parrafos


def extraer_ficha(docx_path: str | Path) -> ClienteContrato:
    docx_path = Path(docx_path)
    parrafos = _parrafos(docx_path)
    # Se quitan los rótulos/banners antes de asignar posiciones fijas y se
    # recuerda a qué bloque (fiscal/entrega) pertenece cada línea de datos.
    items: list[tuple[str, str | None]] = []
    seccion: str | None = None
    banners: list[str] = []
    for p in (x.strip() for x in parrafos):
        if not p:
            continue
        tipo = _clasificar_etiqueta(p)
        if tipo == "envio":
            banners.append(p)
        elif tipo in ("fiscal", "entrega"):
            seccion = tipo
        else:
            items.append((p, seccion))
    no_vacios = [p for p, _ in items]
    seccion_de = {}
    for p, sec in items:
        seccion_de.setdefault(p, sec)

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

    dom_entrega = [d for d in domicilios if seccion_de.get(d) == "entrega"]
    dom_fiscal = [d for d in domicilios if seccion_de.get(d) != "entrega"] or domicilios[:1]
    if dom_fiscal[0] in dom_entrega:
        dom_entrega = [d for d in dom_entrega if d != dom_fiscal[0]]
    domicilio_raw = dom_fiscal[0]

    cliente = ClienteContrato(
        razon_social=razon_social,
        numero_cuenta=numero_cuenta,
        representante_legal=quitar_titulos(representante),
        rfc=rfc or "",
        telefono=telefono or "",
        correo=correo,
    )
    # RFC: la fuente oficial es LAYOUT DE VINCULACION (vinculacion_extractor),
    # así que puede faltar en la ficha (visto en FERTI GREEN, 2026-09-23).
    # Teléfono: solo viene en la ficha (los teléfonos de SAE/control son las
    # líneas a renovar, no un contacto del cliente); si falta se deja en
    # blanco con alerta, nunca se inventa (visto en SERVICIOS PROFESIONALES
    # HERNANDEZ NUÑO, 2026-09-23).
    if not telefono:
        cliente.campos_faltantes.append("telefono")
        cliente.agregar_alerta(
            "La ficha no trae teléfono del cliente; el CONTRATO queda sin teléfono "
            "(hay que capturarlo manualmente)."
        )
    cliente.origen["razon_social"] = docx_path.name
    cliente.origen["numero_cuenta"] = docx_path.name
    cliente.origen["representante_legal"] = docx_path.name
    cliente.origen["rfc_ficha"] = docx_path.name
    cliente.origen["telefono"] = docx_path.name
    cliente.origen["correo"] = docx_path.name

    _asignar_domicilio(cliente, domicilio_raw, "domicilio", "fiscal")
    cliente.origen["domicilio"] = docx_path.name

    if dom_entrega:
        cliente.domicilio_entrega_raw = dom_entrega[0]
        _asignar_domicilio(cliente, dom_entrega[0], "domicilio_entrega", "de entrega")
    elif len(dom_fiscal) > 1:
        # Ficha sin rótulos con un 2º domicilio (formato anterior a 2026-09-23).
        cliente.domicilio_entrega_raw = dom_fiscal[1]
        cliente.agregar_alerta(
            f"La ficha trae {len(dom_fiscal) - 1} domicilio(s) más sin rotular "
            f"({dom_fiscal[1:]!r}); se ignoran, se usa solo el primero como fiscal."
        )

    destino = None
    for b in banners:
        m = _RE_ENVIO.search(_sin_acentos(b.upper()))
        if m:
            destino = m.group(1)
            break
    if destino is None:
        cliente.envio_a = "FISCAL"
        cliente.agregar_alerta(
            "La ficha no indica a qué domicilio se envía el paquete (línea "
            "'EL PAQUETE SE ENVIA A ESTE DOMICILIO:...'); se asumió FISCAL."
        )
    else:
        cliente.envio_a = destino
    if cliente.envio_a == "ENTREGA" and not cliente.domicilio_entrega_raw:
        raise DatosIncompletosError(
            f"'{docx_path.name}' indica que el paquete se envía al domicilio de "
            f"ENTREGA pero no se encontró ese domicilio en la ficha. "
            f"Párrafos: {no_vacios!r}"
        )

    if sin_identificar:
        cliente.agregar_alerta(
            f"Párrafos de la ficha no identificados (ignorados): {sin_identificar!r}"
        )

    if not re.search(r"@", cliente.correo):
        cliente.agregar_alerta(f"El correo extraído no parece válido: {cliente.correo!r}")

    return cliente
