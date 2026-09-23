"""Extrae los adjuntos de un correo (.eml o .msg de Outlook) a una carpeta
y los identifica por nombre de archivo, según el patrón observado en los
correos reales del cliente (ficha .docx, SAE 2.2.xlsx, control de
renovacion.xlsx, LAYOUT DE VINCULACION xlsx).

Por decisión explícita del usuario, en esta primera versión el bot no se
conecta a Outlook automáticamente: el correo se indica manualmente (ya sea
guardado desde Outlook, o reenviado).

Soporta dos formatos, detectados por extensión (`.eml` / `.msg`) — agregado
2026-09-22 porque al exportar el proyecto a otra máquina, Outlook ahí solo
ofrecía guardar los correos como `.msg` (formato binario nativo de
Outlook/OLE), no como `.eml` (texto plano RFC822). `.msg` se lee con la
librería externa `extract-msg` (no requiere tener Outlook instalado);
ambos formatos terminan en la misma `AdjuntosCorreo` vía
`_clasificar_adjuntos`, así que el resto del pipeline no distingue cuál se
usó. Verificado contra un .msg real (mismo correo de CORPORATIVO EN
FARMACIAS que ya se había procesado como .eml): mismos 13 adjuntos."""

from __future__ import annotations

import re
from dataclasses import dataclass
from email import policy
from email.parser import BytesParser
from pathlib import Path


class CorreoIncompletoError(Exception):
    pass


@dataclass
class AdjuntosCorreo:
    asunto: str = ""
    ficha_docx: Path | None = None
    sae_xlsx: Path | None = None
    control_renovacion_xlsx: Path | None = None
    vinculacion_xlsx: Path | None = None
    oferta_comercial_pdf: Path | None = None
    # PDF escaneado de la identificación oficial ("INE_<nombre>.pdf"); trae a
    # mano "Cotejado contra original", el nombre de quien coteja y la FECHA DE
    # COTEJO, que es la fecha de contratación (regla del usuario 2026-09-23).
    ine_pdf: Path | None = None
    # Personas autorizadas para recibir equipos, leídas del cuerpo del correo
    # (ver `extraer_personas_autorizadas`); vacío si el correo no las trae.
    personas_autorizadas: list[str] = None
    otros: list[Path] = None

    def __post_init__(self):
        if self.otros is None:
            self.otros = []
        if self.personas_autorizadas is None:
            self.personas_autorizadas = []


_PATRONES = {
    "sae_xlsx": re.compile(r"^\d*\s*SAE", re.I),
    "control_renovacion_xlsx": re.compile(r"control\s*de\s*renovaci[oó]n", re.I),
    "vinculacion_xlsx": re.compile(r"vinculaci[oó]n", re.I),
    "ine_pdf": re.compile(r"^(?:\d+\s*)?INE[_\s]", re.I),
    # Nombre del archivo no es fijo (se genera ad-hoc por deal, ej.
    # 'PdfPropuestaComercial PDF CORPORATIVO EN F FARMACIA.pdf'), así que
    # el patrón por nombre es un intento best-effort; si no matchea, hay
    # un respaldo por contenido más abajo (busca el texto fijo del PDF).
    "oferta_comercial_pdf": re.compile(r"propuesta|oferta\s*comercial", re.I),
}


def _leer_crudos_eml(correo_path: Path) -> tuple[str, str, list[tuple[str, bytes]]]:
    with open(correo_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)
    crudos = []
    for part in msg.iter_attachments():
        nombre = part.get_filename()
        if not nombre:
            continue
        crudos.append((nombre, part.get_payload(decode=True)))
    cuerpo_part = msg.get_body(preferencelist=("plain", "html"))
    cuerpo = cuerpo_part.get_content() if cuerpo_part is not None else ""
    if cuerpo_part is not None and cuerpo_part.get_content_subtype() == "html":
        cuerpo = re.sub(r"<[^>]+>", " ", cuerpo)
    return msg["subject"] or "", cuerpo, crudos


def _leer_crudos_msg(correo_path: Path) -> tuple[str, str, list[tuple[str, bytes]]]:
    import extract_msg

    with extract_msg.openMsg(str(correo_path)) as msg:
        asunto = (msg.subject or "").replace("\x00", "")
        cuerpo = (msg.body or "").replace("\x00", "")
        crudos = []
        for adjunto in msg.attachments:
            # extract-msg (0.56.1) devuelve el nombre con un '\x00' final en
            # TODOS los adjuntos de este .msg real (propiedad OLE de largo
            # fijo sin recortar) — sin este strip, open() falla con
            # "embedded null character" al crear el archivo.
            nombre = (adjunto.getFilename() or "").replace("\x00", "").strip()
            if nombre and isinstance(adjunto.data, (bytes, bytearray)):
                crudos.append((nombre, adjunto.data))
    return asunto, cuerpo, crudos


# Los nombres de quienes reciben los equipos vienen en el cuerpo del correo
# (acordado con el usuario 2026-09-23), con redacción variable. Formatos vistos
# en correos reales: "PERSONAS AUTORIZADAS A RECIBIR: A, B", "PERSONAS : A, B",
# "...personas autorizadas para recibir: A", "Reciben: A y B". Algunos correos no las traen.
_RE_PERSONAS = re.compile(
    r"(?:personas?\b[^:\n]{0,60}|recib(?:e|en)\b[^:\n]{0,40}):[ \t]*(.+)", re.I
)
_RE_SEPARADOR_PERSONAS = re.compile(r"\s*(?:,|;|/|\by\b)\s*", re.I)


def extraer_personas_autorizadas(cuerpo: str) -> list[str]:
    # Solo el texto anterior al primer encabezado de mensaje reenviado ("De:").
    cuerpo = re.split(r"(?im)^\s*de\s*:", cuerpo or "", maxsplit=1)[0]
    m = _RE_PERSONAS.search(cuerpo)
    if not m:
        return []
    nombres = [n.strip(" .\t\r") for n in _RE_SEPARADOR_PERSONAS.split(m.group(1))]
    return [n for n in nombres if len(n) > 2]


def _clasificar_adjuntos(
    correo_path: Path,
    destino_dir: Path,
    asunto: str,
    crudos: list[tuple[str, bytes]],
    cuerpo: str = "",
) -> AdjuntosCorreo:
    resultado = AdjuntosCorreo(asunto=asunto, personas_autorizadas=extraer_personas_autorizadas(cuerpo))
    for nombre, datos in crudos:
        destino = destino_dir / nombre
        with open(destino, "wb") as out:
            out.write(datos)

        asignado = False
        for campo, patron in _PATRONES.items():
            if patron.search(nombre) and getattr(resultado, campo) is None:
                setattr(resultado, campo, destino)
                asignado = True
                break
        if not asignado and nombre.lower().endswith(".docx") and resultado.ficha_docx is None:
            resultado.ficha_docx = destino
            asignado = True
        if not asignado:
            resultado.otros.append(destino)

    if resultado.oferta_comercial_pdf is None:
        for candidato in list(resultado.otros):
            if candidato.suffix.lower() != ".pdf":
                continue
            try:
                import fitz

                with fitz.open(candidato) as doc:
                    texto = doc[0].get_text() if doc.page_count else ""
            except Exception:
                continue
            if "Formato de Autorizaciones Especiales" in texto:
                resultado.oferta_comercial_pdf = candidato
                resultado.otros.remove(candidato)
                break

    faltantes = [
        campo
        for campo in ("ficha_docx", "sae_xlsx", "control_renovacion_xlsx", "vinculacion_xlsx")
        if getattr(resultado, campo) is None
    ]
    if "ficha_docx" in faltantes or (
        "sae_xlsx" in faltantes and "control_renovacion_xlsx" in faltantes
    ):
        raise CorreoIncompletoError(
            f"No se pudieron identificar todos los adjuntos esperados en "
            f"'{correo_path.name}'. Faltantes: {faltantes}. Adjuntos "
            f"encontrados: {[p.name for p in destino_dir.iterdir()]}"
        )

    return resultado


_LECTORES_POR_EXTENSION = {
    ".eml": _leer_crudos_eml,
    ".msg": _leer_crudos_msg,
}


def extraer_adjuntos(correo_path: str | Path, destino_dir: str | Path) -> AdjuntosCorreo:
    correo_path = Path(correo_path)
    destino_dir = Path(destino_dir)
    destino_dir.mkdir(parents=True, exist_ok=True)

    lector = _LECTORES_POR_EXTENSION.get(correo_path.suffix.lower())
    if lector is None:
        raise ValueError(
            f"Formato de correo no soportado: '{correo_path.suffix}' "
            f"(se espera .eml o .msg) — {correo_path.name}"
        )
    asunto, cuerpo, crudos = lector(correo_path)
    return _clasificar_adjuntos(correo_path, destino_dir, asunto, crudos, cuerpo)
