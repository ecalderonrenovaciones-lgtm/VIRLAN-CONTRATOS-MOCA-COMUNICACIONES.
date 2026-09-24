"""Punto de entrada manual del bot.

Uso:
    py -m virlan_bot.cli procesar --eml "ruta\\al\\correo.eml" [--tipo-venta RENOVACION] [--persona-autorizada "Nombre" [--persona-autorizada "Otro"]] [--fecha-contratacion DD-MM-AAAA]

--eml acepta tanto .eml como .msg (formato nativo de Outlook) — ver
eml_reader.py.

Por decisión explícita del usuario, esta primera versión no se conecta a
Outlook: el correo se indica manualmente. Genera el CONTRATO y la OP en
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
import dataclasses
import datetime as _dt
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
from .ficha_extractor import extraer_ficha, quitar_titulos
from .models import DatosIncompletosError
from .oferta_comercial_extractor import extraer_oferta_comercial
from .op_filler import _trae_equipo as trae_equipo
from .op_filler import llenar_op
from .pdf_optimizer import optimizar_pdf
from .op_to_pdf import exportar_op_a_pdf
from .renovacion_extractor import extraer_lineas_renovacion
from .review import generar_paquete_revision
from .tipo_contratacion import MAPA_TIPO, detectar_tipo_desde_asunto
from .vinculacion_extractor import extraer_vinculacion

_CHECKBOXES_COMUNES = [
    "aviso_privacidad_acepto",
    "autorizo_factura_correo_acepto",
    "tipo_plazo_minimo",
]

# Casillas de la fila de Domicilio Fiscal según a qué domicilio se envía el
# paquete (regla del usuario 2026-09-23): FISCAL => 'Dom. Entrega' y
# 'Correspondencia' de la fila fiscal, como siempre; ENTREGA => se quitan esas
# dos X y se marca 'Correspondencia' en la fila Domicilio Entrega.
_CHECKBOXES_POR_ENVIO = {
    "FISCAL": ["domicilio_fiscal_es_dom_entrega", "domicilio_fiscal_es_correspondencia"],
    "ENTREGA": ["domicilio_entrega_es_correspondencia"],
}


def _slug(texto: str) -> str:
    texto = re.sub(r"[^\w\s-]", "", texto, flags=re.UNICODE).strip()
    return re.sub(r"[\s-]+", "_", texto)


# Archivos que constituyen el "producto terminado" de un cliente (pedido
# por el usuario 2026-09-22): se copian de carpeta_salida a
# CONTRATOS_TERMINADOS_DIR/<cuenta>_<razón_social>/ en cada corrida, sin
# las imágenes de previsualización ni revision.html (esos son solo para
# la revisión humana dentro de salida/, no para el entregable final).
# Nombres posibles de las OP (regla del usuario 2026-09-24: SIM y equipos NO se
# mezclan; si el cliente trae ambos se generan "OP EQUIPOS" y "OP SIM").
_NOMBRES_OP_CONOCIDOS = ["op_borrador", "OP EQUIPOS", "OP SIM"]


def _limpiar_op_obsoletas(carpeta: Path, vigentes: list[str]) -> list[str]:
    """Borra las OP de corridas anteriores que ya no aplican (ej. la OP mezclada
    'op_borrador' cuando ahora hay 'OP EQUIPOS' y 'OP SIM'); devuelve avisos."""
    avisos: list[str] = []
    for nombre in _NOMBRES_OP_CONOCIDOS:
        if nombre in vigentes:
            continue
        for ext in (".pdf", ".xlsx"):
            viejo = carpeta / f"{nombre}{ext}"
            if viejo.exists():
                try:
                    viejo.unlink()
                except OSError:
                    avisos.append(f"No se pudo borrar el archivo obsoleto {viejo} (¿abierto?); bórralo a mano.")
    if "op_borrador" not in vigentes:
        for png in carpeta.glob("op_p*.png"):
            png.unlink(missing_ok=True)
    return avisos


def _copiar_a_contratos_terminados(carpeta_salida: Path, nombre_carpeta: str, nombres_op: list[str]) -> list[str]:
    destino = config.CONTRATOS_TERMINADOS_DIR / nombre_carpeta
    destino.mkdir(parents=True, exist_ok=True)
    avisos = _limpiar_op_obsoletas(destino, nombres_op)
    archivos = ["contrato_borrador.pdf"] + [f"{n}{e}" for n in nombres_op for e in (".pdf", ".xlsx")]
    for nombre_archivo in archivos:
        origen = carpeta_salida / nombre_archivo
        if origen.exists():
            shutil.copy2(origen, destino / nombre_archivo)
    return avisos


def _parsear_fecha(texto: str) -> _dt.date:
    for formato in ("%d-%m-%Y", "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return _dt.datetime.strptime(texto.strip(), formato).date()
        except ValueError:
            pass
    raise ValueError(f"--fecha-contratacion debe ser DD-MM-AAAA (ej. 23-09-2026), se recibió {texto!r}")


def procesar(
    eml_path: str,
    tipo_venta: str | None,
    persona_autorizada: str | None,
    fecha_contratacion: _dt.date | None = None,
) -> Path:
    eml_path = Path(eml_path)
    if not eml_path.exists():
        raise FileNotFoundError(f"No existe el archivo de correo (.eml/.msg): {eml_path}")

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

    if not persona_autorizada and adjuntos.personas_autorizadas:
        persona_autorizada = " Y ".join(
            quitar_titulos(n).upper() for n in adjuntos.personas_autorizadas
        )

    # Regla del usuario 2026-09-23: la fecha de contratación (CONTRATO y OP) es la
    # fecha de cotejo escrita a mano en el INE_*.pdf. Es manuscrita y no se lee
    # de forma automática (un OCR de letra a mano podría equivocarse sin avisar
    # en un documento contractual), así que se indica con --fecha-contratacion.
    alertas_fecha: list[str] = []
    if fecha_contratacion is None:
        fecha_contratacion = _dt.date.today()
        alertas_fecha.append(
            f"No se indicó la fecha de cotejo del INE; se usó la fecha de hoy "
            f"({fecha_contratacion:%d/%m/%Y}) como fecha de contratación en CONTRATO y OP. "
            f"Verificar contra la fecha manuscrita del INE_*.pdf (se muestra en revision.html) "
            f"y, si difiere, repetir con --fecha-contratacion DD-MM-AAAA."
        )
    if adjuntos.ine_pdf is None:
        alertas_fecha.append("No se encontró un adjunto INE_*.pdf; no se puede cotejar la fecha de contratación.")

    valores_contrato, alertas_contrato = construir_valores_contrato(
        cliente,
        persona_autorizada_recibir_equipos=persona_autorizada,
        fecha_contratacion=fecha_contratacion,
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
        checkboxes_activos=[checkbox_tipo_contratacion]
        + _CHECKBOXES_COMUNES
        + _CHECKBOXES_POR_ENVIO[cliente.envio_a],
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

    # Regla del usuario 2026-09-24: NO se pueden mezclar SIM y equipos en una OP.
    # Si el cliente trae ambos: dos OP ("OP EQUIPOS" y "OP SIM"), mismo CONTRATO.
    con_equipo = [l for l in cliente.lineas if trae_equipo(l)]
    sin_equipo = [l for l in cliente.lineas if not trae_equipo(l)]
    if con_equipo and sin_equipo:
        paquetes_op = [("OP EQUIPOS", con_equipo), ("OP SIM", sin_equipo)]
    else:
        paquetes_op = [("op_borrador", cliente.lineas)]
    nombres_op = [n for n, _ in paquetes_op]

    alertas_op: list[str] = []
    if len(paquetes_op) > 1:
        alertas_op.append(
            "El cliente trae SIM y equipos: no se mezclan; se generaron dos OP separadas "
            "(OP EQUIPOS y OP SIM) con el mismo CONTRATO."
        )
    alertas_op += _limpiar_op_obsoletas(carpeta_salida, nombres_op)
    ops_generadas: list[tuple[str, Path]] = []
    for nombre_op, lineas_op in paquetes_op:
        op_xlsx = carpeta_salida / f"{nombre_op}.xlsx"
        op_pdf = carpeta_salida / f"{nombre_op}.pdf"
        resultado_op = llenar_op(
            machote_xlsx=config.MACHOTE_OP_XLSX,
            salida_xlsx=op_xlsx,
            cliente=dataclasses.replace(cliente, lineas=lineas_op),
            ladas_csv_path=ladas_csv,
            tipo_venta=tipo_venta_op,
            lista_precios_xlsx=lista_precios,
            calculo_mpe_xlsx=calculo_mpe,
            oferta_comercial=oferta_comercial,
            fecha_contratacion=fecha_contratacion,
        )
        exportar_op_a_pdf(op_xlsx, op_pdf)
        prefijo = f"[{nombre_op}] " if len(paquetes_op) > 1 else ""
        alertas_op += [prefijo + a for a in resultado_op.alertas]
        ops_generadas.append((nombre_op, op_pdf))

    alertas_peso: list[str] = []
    for pdf in [contrato_pdf] + [p for _, p in ops_generadas]:
        opt = optimizar_pdf(pdf, config.LIMITE_PDF_KB)
        print(f"Tamaño {opt.archivo}: {opt.kb_antes} KB -> {opt.kb_despues} KB (límite {config.LIMITE_PDF_KB} KB)")
        if opt.alerta:
            alertas_peso.append(opt.alerta)

    todas_alertas = (
        alertas_fecha
        + alertas_tipo
        + alertas_contrato
        + resultado_contrato.campos_omitidos
        + [f"CONTRATO — desborde: {d}" for d in resultado_contrato.desbordes]
        + alertas_op
        + alertas_peso
    )

    ruta_revision = generar_paquete_revision(
        cliente=cliente,
        contrato_pdf=contrato_pdf,
        op_pdf=ops_generadas if len(ops_generadas) > 1 else ops_generadas[0][1],
        salida_dir=carpeta_salida,
        alertas_extra=todas_alertas,
        ine_pdf=adjuntos.ine_pdf,
        fecha_contratacion=f"{fecha_contratacion:%d/%m/%Y}",
    )

    todas_alertas += _copiar_a_contratos_terminados(carpeta_salida, carpeta_salida.name, nombres_op)

    print(f"CONTRATO generado: {contrato_pdf}")
    for nombre_op, ruta_op in ops_generadas:
        print(f"{nombre_op} generada: {ruta_op}")
    print(f"Paquete de revisión: {ruta_revision}")
    print(f"Copia en CONTRATOS TERMINADOS: {config.CONTRATOS_TERMINADOS_DIR / carpeta_salida.name}")
    print(f"Total de alertas: {len(cliente.alertas) + len(todas_alertas)} — revisar {ruta_revision.name} antes de enviar.")
    return ruta_revision


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="virlan_bot")
    sub = parser.add_subparsers(dest="comando", required=True)

    p_procesar = sub.add_parser("procesar", help="Genera CONTRATO y OP a partir de un correo .eml o .msg")
    p_procesar.add_argument("--eml", required=True, help="Ruta al correo (.eml o .msg) del cliente")
    p_procesar.add_argument(
        "--tipo-venta",
        default=None,
        choices=["NUEVA", "ADICION", "RENOVACION"],
        help="Si se omite, se detecta automáticamente del asunto del correo",
    )
    p_procesar.add_argument(
        "--persona-autorizada",
        action="append",
        default=None,
        help="Persona autorizada para recibir equipos; repetir la opción por cada persona "
        "(se unen con ' Y ' en mayúsculas, como en el contrato de ejemplo)",
    )

    p_procesar.add_argument(
        "--fecha-contratacion",
        default=None,
        help="Fecha de cotejo manuscrita en el INE_*.pdf (DD-MM-AAAA); es la fecha de "
        "contratación del CONTRATO y la OP. Si se omite se usa hoy, con alerta.",
    )

    args = parser.parse_args(argv)

    if args.comando == "procesar":
        try:
            personas = " Y ".join(quitar_titulos(n).upper() for n in args.persona_autorizada or [] if n.strip())
            fecha = _parsear_fecha(args.fecha_contratacion) if args.fecha_contratacion else None
            procesar(args.eml, args.tipo_venta, personas or None, fecha)
        except DatosIncompletosError as e:
            print(f"ERROR — datos incompletos: {e}", file=sys.stderr)
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
