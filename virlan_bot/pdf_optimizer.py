"""Reduce el peso de los PDF generados (pedido del usuario 2026-09-23: máximo
700 KB por archivo, sin perder legibilidad).

El CONTRATO salía de ~2.2 MB aunque sus imágenes suman ~40 KB: el peso venía
de objetos y streams sin comprimir arrastrados del machote. Una reescritura
SIN PÉRDIDA (limpieza + deflate) lo deja en ~330 KB con texto y píxeles
idénticos (verificado comparando ambos). Solo si aun así excede el límite se
recomprimen las imágenes bajando resolución/calidad de forma gradual.

El archivo solo se reemplaza si el resultado pesa menos que el original (en
la OP, ya menor al límite, la reescritura lo dejaba más grande).
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

import fitz

# (dpi objetivo, calidad JPEG) de menos a más agresivo; solo se usan si la
# reescritura sin pérdida no alcanza el límite.
_NIVELES_IMAGEN = ((150, 70), (120, 60), (96, 50))


@dataclass
class ResultadoOptimizacion:
    archivo: str
    kb_antes: int
    kb_despues: int
    alerta: str | None = None


def _guardar_sin_perdida(doc: "fitz.Document", destino: Path) -> None:
    doc.save(
        str(destino), garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True
    )


def optimizar_pdf(pdf_path: str | Path, limite_kb: int) -> ResultadoOptimizacion:
    pdf_path = Path(pdf_path)
    kb_antes = pdf_path.stat().st_size // 1024
    tmp = pdf_path.with_name(pdf_path.stem + "._opt.pdf")

    try:
        with fitz.open(str(pdf_path)) as doc:
            _guardar_sin_perdida(doc, tmp)
        mejor = tmp.stat().st_size if tmp.stat().st_size < pdf_path.stat().st_size else None

        if (mejor or pdf_path.stat().st_size) // 1024 > limite_kb:
            for dpi, calidad in _NIVELES_IMAGEN:
                with fitz.open(str(pdf_path)) as doc:
                    doc.rewrite_images(dpi_threshold=dpi + 1, dpi_target=dpi, quality=calidad)
                    _guardar_sin_perdida(doc, tmp)
                mejor = tmp.stat().st_size
                if mejor // 1024 <= limite_kb:
                    break

        if mejor is not None and mejor < pdf_path.stat().st_size:
            os.replace(tmp, pdf_path)
    finally:
        if tmp.exists():
            tmp.unlink()

    kb_despues = pdf_path.stat().st_size // 1024
    alerta = None
    if kb_despues > limite_kb:
        alerta = (
            f"{pdf_path.name} pesa {kb_despues} KB y no se logró bajar del límite de "
            f"{limite_kb} KB; revisar manualmente."
        )
    return ResultadoOptimizacion(pdf_path.name, kb_antes, kb_despues, alerta)
