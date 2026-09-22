"""Extrae los adjuntos de un correo .eml a una carpeta y los identifica por
nombre de archivo, según el patrón observado en los correos reales del
cliente (ficha .docx, SAE 2.2.xlsx, control de renovacion.xlsx, LAYOUT DE
VINCULACION xlsx).

Por decisión explícita del usuario, en esta primera versión el bot no se
conecta a Outlook automáticamente: el .eml se indica manualmente (ya sea
guardado desde Outlook, o reenviado)."""

from __future__ import annotations

import email
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
    otros: list[Path] = None

    def __post_init__(self):
        if self.otros is None:
            self.otros = []


_PATRONES = {
    "sae_xlsx": re.compile(r"^\d*\s*SAE", re.I),
    "control_renovacion_xlsx": re.compile(r"control\s*de\s*renovaci[oó]n", re.I),
    "vinculacion_xlsx": re.compile(r"vinculaci[oó]n", re.I),
}


def extraer_adjuntos(eml_path: str | Path, destino_dir: str | Path) -> AdjuntosCorreo:
    eml_path = Path(eml_path)
    destino_dir = Path(destino_dir)
    destino_dir.mkdir(parents=True, exist_ok=True)

    with open(eml_path, "rb") as f:
        msg = BytesParser(policy=policy.default).parse(f)

    resultado = AdjuntosCorreo(asunto=msg["subject"] or "")
    for part in msg.iter_attachments():
        nombre = part.get_filename()
        if not nombre:
            continue
        datos = part.get_payload(decode=True)
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
            f"'{eml_path.name}'. Faltantes: {faltantes}. Adjuntos "
            f"encontrados: {[p.name for p in destino_dir.iterdir()]}"
        )

    return resultado
