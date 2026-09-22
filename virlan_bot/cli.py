"""Punto de entrada manual del bot.

Uso:
    py -m virlan_bot.cli procesar --eml "ruta\\al\\correo.eml" [--tipo-venta RENOVACION] [--persona-autorizada "Nombre"]

Por decisión explícita del usuario, esta primera versión no se conecta a
Outlook: el .eml se indica manualmente. Genera el CONTRATO y la OP en
salida/<cuenta>_<razon_social>/ junto con un paquete de revisión humana
(revision.html) — el pipeline nunca debe considerarse terminado sin que una
persona revise ese paquete. Además copia los 3 archivos finales (CONTRATO
PDF, OP PDF, OP Excel) a CONTRATOS TERMINADOS/<cuenta>_<razon_social>/
(carpeta hermana del proyecto, ver config.CONTRATOS_TERMINADOS_DIR) — esa
carpeta es solo para visualizar/entregar el resultado, salida/ sigue
siendo la carpeta de trabajo real con todo (incluida revision.html).
"""

from __future__ import annotations

import argparse
import hashlib
import re
import shutil
import sys
from pathlib import Path

from . import config
from .contrato_fieldmap import cargar_fieldmap
from .contrato_filler import llenar_contrato
from .contrato_valores import construir_valores_contrato
from .eml_reader import extraer_adjuntos
from .ficha_extractor import extraer_ficha
from .models import DatosIncompletosError
from .oferta_comercial_extractor import extraer_oferta_comercial
from .op_filler import llenar_op
from .op_to_pdf import exportar_op_a_pdf
from .renovacion_extractor import extraer_lineas_renovacion
from .review import generar_paquete_revision
from .tipo_contratacion import MAPA_TIPO, detectar_tipo_desde_asunto
from .vinculacion_extractor import extraer_vinculacion

_CHECKBOXES_COMUNES = [
    "aviso_privacidad_acepto",
    "autorizo_factura_correo_acepto",
    "tipo_plazo_minimo",
    "domicilio_fiscal_es_dom_entrega",
    "domicilio_fiscal_es_correspondencia",
]


def _slug(texto: str) -> str:
    texto = re.sub(r"[^\w\s-]", "", texto, flags=re.UNICODE).strip()
    return re.sub(r"[\s-]+", "_", texto)


# Archivos que constituyen el "producto terminado" de un cliente (pedido
# por el usuario 2026-09-22): se copian de carpeta_salida a
# CONTRATOS_TERMINADOS_DIR/<cuenta>_<razón_social>/ en cada corrida, sin
# las imágenes de previsualización ni revision.html (esos son solo para
# la revisión humana dentro de salida/, no para el entregable final).
_ARCHIVOS_TERMINADOS = ["contrato_borrador.pdf", "op_borrador.pdf", "op_borrador.xlsx"]


def _copiar_a_contratos_terminados(carpeta_salida: Path, nombre_carpeta: str) -> None:
    destino = config.CONTRATOS_TERMINADOS_DIR / nombre_carpeta
    destino.mkdir(parents=True, exist_ok=True)
    for nombre_archivo in _ARCHIVOS_TERMINADOS:
        origen = carpeta_salida / nombre_archivo
        if origen.exists():
            shutil.copy2(origen, destino / nombre_archivo)


def procesar(eml_path: str, tipo_venta: str | None, persona_autorizada: str | None) -> Path:
    eml_path = Path(eml_path)
    if not eml_path.exists():
        raise FileNotFoundError(f"No existe el archivo .eml: {eml_path}")

    nombre_corto = _slug(eml_path.stem)[:40] + "_" + hashlib.sha1(str(eml_path).encode()).hexdigest()[:8]
    carpeta_temporal = config.SALIDA_DIR / "_tmp_adjuntos" / nombre_corto
    adjuntos = extraer_adjuntos(eml_path, carpeta_temporal)

    alertas_tipo: list[str] = []
    if tipo_venta is None:
        tipo_venta, alertas_tipo = detectar_tipo_desde_asunto(adjuntos.asunto)
    elif tipo_venta not in MAPA_TIPO:
        raise ValueError(f"--tipo-venta debe ser uno de {list(MAPA_TIPO)}, se recibió {tipo_venta!r}")

    cliente = extraer_ficha(adjuntos.ficha_docx)
    extraer_lineas_renovacion(
        cliente,
        control_renovacion_path=adjuntos.control_renovacion_xlsx,
        sae_path=adjuntos.sae_xlsx,
    )
    if adjuntos.vinculacion_xlsx:
        extraer_vinculacion(cliente, adjuntos.vinculacion_xlsx)
    else:
        cliente.agregar_alerta(
            "No se encontró el adjunto de LAYOUT DE VINCULACION en el correo; "
            "RFC e identificación oficial no se pudieron completar."
        )

    carpeta_salida = config.SALIDA_DIR / f"{cliente.numero_cuenta}_{_slug(cliente.razon_social)}"
    carpeta_salida.mkdir(parents=True, exist_ok=True)

    valores_contrato, alertas_contrato = construir_valores_contrato(
        cliente, persona_autorizada_recibir_equipos=persona_autorizada
    )

    checkbox_tipo_contratacion, tipo_venta_op = MAPA_TIPO[tipo_venta]

    fieldmap = cargar_fieldmap(config.CONTRATO_FIELDMAP_VERSION)
    contrato_pdf = carpeta_salida / "contrato_borrador.pdf"
    resultado_contrato = llenar_contrato(
        plantilla_pdf=config.MACHOTE_CONTRATO_PDF,
        salida_pdf=contrato_pdf,
        fieldmap=fieldmap,
        valores=valores_contrato,
        rfc_cliente=cliente.rfc or None,
        checkboxes_activos=[checkbox_tipo_contratacion] + _CHECKBOXES_COMUNES,
    )

    lista_precios = config.lista_precios_xlsx()
    calculo_mpe = config.calculo_mpe_xlsx()
    ladas_csv = config.ladas_csv()
    if not lista_precios:
        cliente.agregar_alerta(
            "No se encontró ningún 'Lista de Precios*.xlsx' en DOCUMENTOS CONSULTA; "
            "Precio de lista quedará en blanco en la OP."
        )
    if not calculo_mpe:
        cliente.agregar_alerta(
            "No se encontró ningún '*Cálculo de MPE*.xlsx' en DOCUMENTOS CONSULTA; "
            "MPE quedará en blanco en la OP."
        )
    if not ladas_csv:
        cliente.agregar_alerta(
            "No se encontró ningún 'ladas_mexico*.csv' en DOCUMENTOS CONSULTA; "
            "Ciudad DN quedará en blanco en la OP."
        )

    oferta_comercial = None
    if adjuntos.oferta_comercial_pdf:
        oferta_comercial = extraer_oferta_comercial(adjuntos.oferta_comercial_pdf)
        if not oferta_comercial.folio_producto and not oferta_comercial.movimientos:
            cliente.agregar_alerta(
                f"Se encontró un adjunto de oferta comercial "
                f"('{adjuntos.oferta_comercial_pdf.name}') pero no se pudo "
                f"extraer ni el folio de producto ni los movimientos; "
                f"revisar el formato del PDF manualmente."
            )

    op_xlsx = carpeta_salida / "op_borrador.xlsx"
    op_pdf = carpeta_salida / "op_borrador.pdf"
    resultado_op = llenar_op(
        machote_xlsx=config.MACHOTE_OP_XLSX,
        salida_xlsx=op_xlsx,
        cliente=cliente,
        ladas_csv_path=ladas_csv,
        tipo_venta=tipo_venta_op,
        lista_precios_xlsx=lista_precios,
        calculo_mpe_xlsx=calculo_mpe,
        oferta_comercial=oferta_comercial,
    )
    exportar_op_a_pdf(op_xlsx, op_pdf)

    todas_alertas = (
        alertas_tipo
        + alertas_contrato
        + resultado_contrato.campos_omitidos
        + [f"CONTRATO — desborde: {d}" for d in resultado_contrato.desbordes]
        + resultado_op.alertas
    )

    ruta_revision = generar_paquete_revision(
        cliente=cliente,
        contrato_pdf=contrato_pdf,
        op_pdf=op_pdf,
        salida_dir=carpeta_salida,
        alertas_extra=todas_alertas,
    )

    _copiar_a_contratos_terminados(carpeta_salida, carpeta_salida.name)

    print(f"CONTRATO generado: {contrato_pdf}")
    print(f"OP generada: {op_pdf}")
    print(f"Paquete de revisión: {ruta_revision}")
    print(f"Copia en CONTRATOS TERMINADOS: {config.CONTRATOS_TERMINADOS_DIR / carpeta_salida.name}")
    print(f"Total de alertas: {len(cliente.alertas) + len(todas_alertas)} — revisar {ruta_revision.name} antes de enviar.")
    return ruta_revision


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="virlan_bot")
    sub = parser.add_subparsers(dest="comando", required=True)

    p_procesar = sub.add_parser("procesar", help="Genera CONTRATO y OP a partir de un .eml")
    p_procesar.add_argument("--eml", required=True, help="Ruta al correo .eml del cliente")
    p_procesar.add_argument(
        "--tipo-venta",
        default=None,
        choices=["NUEVA", "ADICION", "RENOVACION"],
        help="Si se omite, se detecta automáticamente del asunto del correo",
    )
    p_procesar.add_argument(
        "--persona-autorizada", default=None, help="Nombre(s) autorizados para recibir equipos"
    )

    args = parser.parse_args(argv)

    if args.comando == "procesar":
        try:
            procesar(args.eml, args.tipo_venta, args.persona_autorizada)
        except DatosIncompletosError as e:
            print(f"ERROR — datos incompletos: {e}", file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
