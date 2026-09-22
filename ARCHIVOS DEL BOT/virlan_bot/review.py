"""Arma un paquete de revisión humana: una página HTML con los datos
extraídos (y su origen), todas las alertas del pipeline, y una imagen por
página de cada PDF generado (CONTRATO y OP) para verificación visual rápida.

El pipeline nunca debe marcar un CONTRATO/OP como "listo para enviar" sin
que una persona revise este paquete — es un documento contractual."""

from __future__ import annotations

import html
from dataclasses import asdict
from pathlib import Path

import fitz

from .models import ClienteContrato


def _renderizar_paginas(pdf_path: Path, salida_dir: Path, prefijo: str, dpi: int = 130) -> list[str]:
    nombres = []
    doc = fitz.open(str(pdf_path))
    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi)
        nombre = f"{prefijo}_p{i + 1}.png"
        pix.save(str(salida_dir / nombre))
        nombres.append(nombre)
    doc.close()
    return nombres


def _fila_tabla(campo: str, valor: str, origen: str) -> str:
    return (
        f"<tr><td>{html.escape(campo)}</td><td>{html.escape(str(valor))}</td>"
        f"<td class='origen'>{html.escape(origen)}</td></tr>"
    )


def generar_paquete_revision(
    cliente: ClienteContrato,
    contrato_pdf: str | Path,
    op_pdf: str | Path,
    salida_dir: str | Path,
    alertas_extra: list[str] | None = None,
) -> Path:
    salida_dir = Path(salida_dir)
    salida_dir.mkdir(parents=True, exist_ok=True)

    imagenes_contrato = _renderizar_paginas(Path(contrato_pdf), salida_dir, "contrato")
    imagenes_op = _renderizar_paginas(Path(op_pdf), salida_dir, "op")

    campos_relevantes = [
        "razon_social", "numero_cuenta", "representante_legal", "rfc",
        "tipo_identificacion", "numero_identificacion", "telefono", "correo",
        "domicilio_calle", "domicilio_numero", "domicilio_colonia",
        "domicilio_ciudad", "domicilio_estado", "domicilio_cp",
    ]
    filas_cliente = "".join(
        _fila_tabla(campo, getattr(cliente, campo), cliente.origen.get(campo, "—"))
        for campo in campos_relevantes
    )

    filas_lineas = "".join(
        f"<tr><td>{html.escape(l.telefono)}</td><td>{html.escape(l.plan_tarifario)}</td>"
        f"<td>{html.escape(l.plazo_meses)}</td><td>{html.escape(l.marca_modelo_color)}</td>"
        f"<td class='origen'>{html.escape(l.fuente)}</td></tr>"
        for l in cliente.lineas
    )

    todas_alertas = list(cliente.alertas) + list(alertas_extra or [])
    lista_alertas = (
        "".join(f"<li>{html.escape(a)}</li>" for a in todas_alertas)
        if todas_alertas
        else "<li class='ok'>Sin alertas.</li>"
    )
    faltantes = (
        ", ".join(cliente.campos_faltantes) if cliente.campos_faltantes else "Ninguno"
    )

    imgs_contrato_html = "".join(
        f"<img src='{n}' alt='CONTRATO página {i+1}'>" for i, n in enumerate(imagenes_contrato)
    )
    imgs_op_html = "".join(
        f"<img src='{n}' alt='OP página {i+1}'>" for i, n in enumerate(imagenes_op)
    )

    html_doc = f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<title>Revisión — {html.escape(cliente.razon_social)} ({html.escape(cliente.numero_cuenta)})</title>
<style>
  body {{ font-family: Segoe UI, Arial, sans-serif; margin: 24px; color: #1a1a1a; }}
  h1 {{ font-size: 20px; }}
  h2 {{ font-size: 16px; margin-top: 32px; border-bottom: 2px solid #333; padding-bottom: 4px; }}
  table {{ border-collapse: collapse; width: 100%; margin-top: 8px; }}
  td, th {{ border: 1px solid #ccc; padding: 6px 10px; text-align: left; font-size: 13px; }}
  .origen {{ color: #666; font-size: 12px; }}
  ul.alertas {{ background: #fff3cd; border: 1px solid #ffe69c; padding: 12px 12px 12px 32px; }}
  ul.alertas li {{ margin-bottom: 4px; }}
  .ok {{ color: #2a7a2a; }}
  img {{ width: 100%; max-width: 900px; border: 1px solid #999; margin-bottom: 16px; display: block; }}
</style></head>
<body>
  <h1>Revisión — {html.escape(cliente.razon_social)} (cuenta {html.escape(cliente.numero_cuenta)})</h1>
  <p><b>Campos faltantes:</b> {html.escape(faltantes)}</p>

  <h2>Alertas del pipeline</h2>
  <ul class="alertas">{lista_alertas}</ul>

  <h2>Datos del cliente</h2>
  <table><tr><th>Campo</th><th>Valor</th><th>Origen</th></tr>{filas_cliente}</table>

  <h2>Líneas a renovar</h2>
  <table><tr><th>Teléfono</th><th>Plan</th><th>Plazo</th><th>Equipo</th><th>Origen</th></tr>{filas_lineas}</table>

  <h2>CONTRATO generado</h2>
  {imgs_contrato_html}

  <h2>OP generada</h2>
  {imgs_op_html}
</body></html>
"""
    ruta_html = salida_dir / "revision.html"
    ruta_html.write_text(html_doc, encoding="utf-8")
    return ruta_html
