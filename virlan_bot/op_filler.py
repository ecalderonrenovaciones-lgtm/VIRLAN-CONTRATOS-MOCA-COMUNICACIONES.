"""Llena la ORDEN DE PROGRAMACIÓN (OP) en Excel a partir de un ClienteContrato.

Descubrimiento clave (verificado abriendo el machote con Excel vía COM): las
casillas de "TIPO DE VENTA" (NUEVA / ADICIÓN / RENOVACIÓN) NO son celdas de
Excel, son 3 formas (shapes) rectangulares independientes superpuestas sobre
la hoja ("Rectangle 1"=NUEVA, "Rectangle 2"=ADICIÓN, "Rectangle 3"=RENOVACIÓN,
de izquierda a derecha), que se marcan escribiendo "X" en
shape.TextFrame2.TextRange.Text. openpyxl no garantiza preservar estas formas
al reguardar, así que la OP se edita con Excel real vía win32com (no
openpyxl), igual que la exportación a PDF (ver op_to_pdf.py).

Celdas de encabezado (hoja 'OP_OK', verificadas contra el machote y contra un
ejemplo ya terminado):
  C9  Nombre/Razón Social      J9  Cuenta          N9  RFC
  S9  Folio de Contrato (se deja en blanco, lo asigna AT&T después)
  N12 Identificación Oficial, formato "<tipo> //<numero>" (ej. "INE //2090776014")
  S12 Dirección de Entrega
  X6  Fecha (el machote trae la del ejemplo como valor fijo, no =TODAY(); se
      sobreescribe con la fecha de contratación = fecha de cotejo del INE, igual
      que el CONTRATO; ver cli.py)
  G12 Folio de Autorización (`_folio_autorizacion`) — un solo folio para
      toda la OP (no es por línea), según regla dada por el usuario
      2026-09-22: 240729MKT0305 si el cliente trae equipo (mayoría de los
      casos, no llega PDF de "oferta comercial"); si es solo SIM sin
      equipo, 240729MKT0605 a 12 meses o 240717MKT0205 a 24 meses. Se deja
      en blanco + alerta si no se puede determinar (sin líneas, mezcla de
      con/sin equipo, o plazo sin equipo distinto de 12/24).

Tabla de líneas (encabezados en fila 15, datos desde fila 16):
  C Núm. de líneas (=1 por fila)   D Ciudad DN   E Plan tarifario
  F Plazo (mínimo)                 H Marca/Modelo/Color/Capacidad
  T DN (renovación) = teléfono
  U MPP/CPP = se lee del sufijo del nombre del Plan tarifario ('CPP' tal
    cual, o las letras sueltas 'C M' que significan Controlado+MPP y se
    interpretan como 'MPP'); si el nombre no trae ninguno de los dos se usa
    de respaldo la columna 'Modalidad MPP/CPP' del SAE (`_modalidad_mpp_cpp`).
  J Precio de lista  (catalogo_precios.buscar_precio_lista)
  L Cuotas = Plazo
  K Pago inicial/Diferencial de equipo (PIE) = columna de la hoja 'Precios'
    de la lista de precios vigente que corresponde al número de plan y
    plazo de la línea (catalogo_precios.buscar_precio_lista,
    `diferencial_equipo`; ver `_columna_diferencial` ahí) — fuente
    corregida 2026-09-22: antes se usaba 'Costo de Equipo' de control de
    renovacion.xlsx, que el usuario confirmó que puede traer errores de
    captura (verificado con un caso real: control de renovación traía
    $26,679 y el valor correcto según la lista de precios, plan 599 a 24
    meses, era $29,679). 'Costo de Equipo' solo se usa como respaldo si no
    se puede determinar el diferencial desde la lista de precios (plan o
    plazo fuera del catálogo conocido, o lista de precios no disponible),
    y si ambas fuentes existen y difieren se genera una alerta para
    revisión manual. Se deja "N/A" si ninguna fuente da un valor > 0.
  N MPE: criterio principal la calculadora (calculadora_mpe.calcular_mpe);
    criterio alternativo (usuario 2026-09-23) (Precio de lista - Diferencial
    unitario) / plazo con TRUNC(x+0.01, 2) (`mpe_desde_lista`), que se usa si
    la calculadora falla, y gana si ambos difieren (el precio correcto es el de
    la lista de precios). Usa como PIE el mismo valor que
    terminó en la columna K (el PIE del MPE debe ser siempre el mismo
    'Diferencial de equipo unitario' de la columna K, regla confirmada por
    el usuario).
  P Addon CTRL = "X" y S Pago mensual de servicios adicionales = 50,
    SOLO si el nombre del Plan tarifario trae el sufijo "CTRL" (regla
    confirmada por el usuario; precio fijo de $50).
  Y Pago total mensual unitario = número del plan (299/399/.../1499,
    extraído del nombre del Plan tarifario) + 50 si aplica Addon CTRL; si
    la línea matchea un movimiento de la oferta comercial (ver abajo), el
    plan y el Addon CTRL se multiplican antes por (1 - %DMR).
  Fila 31 (Total a Pagar) ya trae la fórmula =SUM(Y16:Y30) en el machote:
    basta con llenar Y correctamente, no hace falta tocarla.
Oferta comercial (PDF "Formato de Autorizaciones Especiales", opcional,
ver oferta_comercial_extractor.py) — cuando llega adjunta al correo, es la
fuente más autorizada para (descubierto 2026-09-22 con un caso real):
  G12 Folio de Autorización = "FOLIO DE PRODUCTO" del PDF (reemplaza el
      folio fijo de `_folio_autorizacion`).
  Meses Gratis (shapes "Rectangle 4"/"Rectangle 9"/"Rectangle 10" en
      J12/K12/L12) = los 3 números de "MESES DE RENTA GRATIS".
  Por línea, si matchea un movimiento (mismo plan + equipo, ver
      `buscar_movimiento_para_linea`): O Descuento multilínea unitario =
      "%DMR"; K Pago inicial/Diferencial de equipo = el monto "$... sin
      IVA c/u" (tiene prioridad sobre la lista de precios y sobre 'Costo
      de Equipo' de control de renovación).
Observaciones (celdas C35:C37, " PLAN <número>: REDES: ... STREAMING: ..."):
'REDES SOCIALES' y 'STREAMING' de SAE 2.2.xlsx varían por plan tarifario
(regla confirmada por el usuario), así que se escribe una fila por cada
combinación de plan distinta presente en el cliente (máximo 3, las únicas
filas libres que deja el machote antes de la tabla de Abrev./Addón).
Firma del suscriptor (shape 'TextBox 5'): siempre el Representante Legal
del cliente (ficha .docx), nunca el nombre reciclado del machote.
Las columnas que aún no tienen fórmula identificada fuera de la oferta
comercial (Q Addon's de datos, V Protección de equipo, X Otros servicios)
se dejan como "N/A"."""

from __future__ import annotations

import datetime as _dt
import re
import shutil
from dataclasses import dataclass, field as dc_field
from pathlib import Path

import win32com.client as win32

from .calculadora_mpe import EquipoNoEncontradoError as MPENoEncontradoError
from .calculadora_mpe import calcular_mpe, mpe_desde_lista
from .catalogo_precios import EquipoNoEncontradoError as PrecioNoEncontradoError
from .catalogo_precios import buscar_precio_lista
from .ladas_lookup import LadaNoEncontradaError, ciudad_dn_por_telefono
from .models import ClienteContrato
from .oferta_comercial_extractor import OfertaComercial, buscar_movimiento_para_linea

_ADDON_CTRL_PRECIO = 50
_HOJA = "OP_OK"
_PRIMERA_FILA_TABLA = 16
# Última fila de la tabla de líneas en el machote. OJO: la fila 31 trae la
# fórmula de "Total a Pagar" y desde la fila 33 empieza contenido FIJO
# (Observaciones y la tabla de abreviaturas Abrev./Addón) que nunca debe
# limpiarse ni tocarse — nunca subir este límite sin revisar el machote.
_ULTIMA_FILA_TABLA = 30

_CHECKBOX_TIPO_VENTA = {
    "NUEVA": "Rectangle 1",
    "ADICION": "Rectangle 2",
    "RENOVACION": "Rectangle 3",
}

_TEXTBOX_FIRMA_SUSCRIPTOR = "TextBox 5"

# 3 shapes (no celdas — mismo caso que los checkboxes de Tipo de Venta) que
# muestran los "Meses Gratis" en J12/K12/L12, de izquierda a derecha.
_SHAPES_MESES_GRATIS = ["Rectangle 4", "Rectangle 9", "Rectangle 10"]

# El número del plan es el monto que sigue a "Negocios" en el nombre del
# plan tarifario, pero a veces trae un "$" u otro separador de por medio
# (ej. "ATT Ármalo Negocios $599 C M"); solo puede ser uno de estos montos
# (regla confirmada por el usuario 2026-09-22), así que se busca ese
# whitelist explícito en vez de "el primer número tras Negocios".
_PATRON_NUMERO_PLAN = re.compile(r"Negocios\D*(299|399|499|599|799|1299|1499)\b", re.I)

# Columnas de la tabla que el machote puede traer con datos de un cliente
# anterior (reciclado) y que todavía no se recalculan automáticamente (no
# se encontró su fórmula/fuente). Se limpian explícitamente en cada fila
# para no dejar cifras de otro cliente mezcladas con los datos nuevos.
# J (Precio de lista), L (Cuotas) y N (MPE) SÍ se calculan (ver más abajo).
_COLUMNAS_PENDIENTES = ["K", "O", "P", "Q", "S", "V", "X", "Y"]
_TODAS_LAS_COLUMNAS_TABLA = [
    "C", "D", "E", "F", "H", "J", "K", "L", "N", "O", "P", "Q", "S", "T", "U", "V", "X", "Y",
]


@dataclass
class ResultadoOP:
    lineas_escritas: int = 0
    alertas: list[str] = dc_field(default_factory=list)


def _set_con_ajuste(ws, rango: str, valor) -> None:
    """Escribe `valor` en `rango` y activa 'Reducir hasta ajustar'
    (ShrinkToFit) para que, si el texto no cabe en el ancho de la celda
    (o del rango combinado), Excel reduzca la letra hasta que quepa en
    una sola línea dentro del recuadro — en vez de recortarse (celdas
    sueltas) o desbordarse verticalmente sobre la fila de abajo (celdas
    combinadas con WrapText), como pasaba antes (regla pedida por el
    usuario 2026-09-22, verificado visualmente con Marca/Modelo/Color y
    Dirección de Entrega). WrapText y ShrinkToFit son excluyentes en
    Excel, por eso se apaga WrapText explícitamente."""
    celda = ws.Range(rango)
    celda.Value = valor
    celda.WrapText = False
    celda.ShrinkToFit = True


def _es_addon_ctrl(plan_tarifario: str) -> bool:
    """Addon CTRL = 'X' (col P) y $50 en Pago mensual de servicios
    adicionales unitario (col S), que luego se suman al número del plan en
    Pago total mensual unitario (col Y). Aplica si el nombre del plan trae
    la palabra 'CTRL', o si trae el sufijo abreviado 'C M'/'CM' (regla
    confirmada por el usuario 2026-09-22: ahí la 'C' ya representa control,
    igual que en _modalidad_mpp_cpp)."""
    texto = plan_tarifario or ""
    return bool(re.search(r"\bCTRL\b", texto, re.I) or re.search(r"\bC\s*M\b", texto, re.I))


def _modalidad_mpp_cpp(plan_tarifario: str, valor_extraido: str | None) -> str:
    """Determina la modalidad MPP/CPP a partir del sufijo del nombre del
    Plan tarifario: 'CPP' tal cual, o las letras sueltas 'C M' (Control +
    MPP) que se interpretan como 'MPP' (regla confirmada por el usuario
    2026-09-22: el sufijo del plan puede venir abreviado así). Si el nombre
    no trae ninguno de los dos, se usa como respaldo el valor ya extraído
    de la columna 'Modalidad MPP/CPP' del SAE."""
    texto = plan_tarifario or ""
    if re.search(r"\bCPP\b", texto, re.I):
        return "CPP"
    if re.search(r"\bC\s*M\b", texto, re.I):
        return "MPP"
    if re.search(r"\bMPP\b", texto, re.I):
        return "MPP"
    return valor_extraido or ""


def _numero_plan(plan_tarifario: str) -> float | None:
    m = _PATRON_NUMERO_PLAN.search(plan_tarifario or "")
    return float(m.group(1)) if m else None


# Folio de Autorización (celda única G12, aplica a toda la OP, no es por
# línea): depende de si el cliente trae equipo o es una renovación de solo
# SIM, y en ese segundo caso del plazo (regla dada por el usuario
# 2026-09-22). Cuando trae equipo casi nunca llega el PDF de "oferta
# comercial" del correo (es la mayoría de los casos) y aun así corresponde
# este folio fijo.
_FOLIO_CON_EQUIPO = "240729MKT0305"
_FOLIO_SIN_EQUIPO_12M = "240729MKT0605"
_FOLIO_SIN_EQUIPO_24M = "240717MKT0205"

# Línea "sin equipo" (solo SIM): primer caso real visto 2026-09-23 (DESECHABLES
# MANOLO, ADICIÓN por portabilidad): SAE/control traen Modelo = "SIM CARD" y
# Color = "NA". Las demás variantes son suposiciones sin verificar.
_MARCADORES_SIN_EQUIPO = {"SIN EQUIPO", "SOLO SIM", "SIM", "SIM CARD", "N/A", "NA"}


def _trae_equipo(linea) -> bool:
    modelo = (linea.modelo or "").strip().upper()
    return bool(modelo) and modelo not in _MARCADORES_SIN_EQUIPO


def _plazo_int(linea) -> int | None:
    try:
        return int(str(linea.plazo_meses).strip())
    except (TypeError, ValueError):
        return None


def _folio_autorizacion(cliente: ClienteContrato) -> tuple[str | None, str | None]:
    """Devuelve (folio, alerta) — ver constantes _FOLIO_* arriba."""
    if not cliente.lineas:
        return None, (
            "Folio de Autorización se dejó en blanco (no hay líneas para "
            "determinar si el cliente trae equipo) — completar manualmente."
        )

    con_equipo = [l for l in cliente.lineas if _trae_equipo(l)]
    sin_equipo = [l for l in cliente.lineas if not _trae_equipo(l)]

    if con_equipo:
        alerta = None
        if sin_equipo:
            alerta = (
                "El cliente tiene líneas con equipo y sin equipo mezcladas; "
                "el Folio de Autorización es un solo campo por toda la OP, "
                f"se usó el de 'trae equipo' ({_FOLIO_CON_EQUIPO}) — "
                "verificar si también aplica a las líneas sin equipo."
            )
        return _FOLIO_CON_EQUIPO, alerta

    plazos = {_plazo_int(l) for l in sin_equipo}
    if plazos == {12}:
        return _FOLIO_SIN_EQUIPO_12M, None
    if plazos == {24}:
        return _FOLIO_SIN_EQUIPO_24M, None
    return None, (
        "Folio de Autorización se dejó en blanco: el cliente no trae "
        "equipo (solo SIM) pero el plazo no es uniformemente 12 o 24 meses "
        f"(plazos encontrados: {sorted(p for p in plazos if p is not None)}) "
        "— completar manualmente."
    )


_FILAS_OBSERVACIONES_REDES = ["C35", "C36", "C37"]


def _limpiar_texto_sae(texto: str) -> str:
    # SAE trae saltos de línea (\r\n) dentro de la celda de REDES/STREAMING;
    # se normalizan a espacio para que no se vea '_x000D_' al reescribirlos
    # con win32com.
    return re.sub(r"\s+", " ", texto).strip()


def _redes_streaming_por_plan(cliente: ClienteContrato) -> list[str]:
    """Las 'REDES SOCIALES'/'STREAMING' de SAE varían por plan tarifario
    (regla confirmada por el usuario: ej. plan 399 trae una combinación
    distinta a la del plan 1299), así que se muestra una línea por cada
    combinación de plan distinta presente en el cliente, no solo la de la
    primera línea."""
    vistos: dict[str, str] = {}
    for linea in cliente.lineas:
        if not linea.redes_sociales and not linea.streaming:
            continue
        numero_plan = _numero_plan(linea.plan_tarifario)
        etiqueta_plan = (
            f"PLAN {int(numero_plan)}" if numero_plan is not None else (linea.plan_tarifario or "")
        )
        if etiqueta_plan in vistos:
            continue
        redes = _limpiar_texto_sae(linea.redes_sociales) if linea.redes_sociales else "N/A"
        streaming = _limpiar_texto_sae(linea.streaming) if linea.streaming else "N/A"
        vistos[etiqueta_plan] = f" {etiqueta_plan}: REDES: {redes}    STREAMING: {streaming}"
    return list(vistos.values())


def _domicilio_entrega(cliente: ClienteContrato) -> str:
    # Si el paquete se envía al domicilio de ENTREGA (regla del usuario
    # 2026-09-23) la OP lleva ese domicilio y ya no el fiscal.
    if cliente.envio_a == "ENTREGA":
        partes = [
            cliente.entrega_calle_numero(),
            cliente.domicilio_entrega_colonia,
            cliente.domicilio_entrega_ciudad,
            cliente.domicilio_entrega_estado,
            cliente.domicilio_entrega_cp,
        ]
        return ", ".join(p for p in partes if p)
    partes = [
        cliente.domicilio_calle_numero(),
        cliente.domicilio_colonia,
        cliente.domicilio_ciudad,
        cliente.domicilio_estado,
        cliente.domicilio_cp,
    ]
    return ", ".join(p for p in partes if p)


def llenar_op(
    machote_xlsx: str | Path,
    salida_xlsx: str | Path,
    cliente: ClienteContrato,
    ladas_csv_path: str | Path | None,
    tipo_venta: str = "RENOVACION",
    lista_precios_xlsx: str | Path | None = None,
    calculo_mpe_xlsx: str | Path | None = None,
    oferta_comercial: OfertaComercial | None = None,
    fecha_contratacion: _dt.date | None = None,
) -> ResultadoOP:
    """Copia machote_xlsx a salida_xlsx y lo llena con los datos del cliente.
    Requiere Excel instalado (usa win32com para preservar las formas de los
    checkboxes de Tipo de Venta)."""
    if tipo_venta not in _CHECKBOX_TIPO_VENTA:
        raise ValueError(
            f"tipo_venta debe ser uno de {list(_CHECKBOX_TIPO_VENTA)}, se recibió {tipo_venta!r}"
        )

    salida_xlsx = Path(salida_xlsx)
    salida_xlsx.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy(machote_xlsx, salida_xlsx)

    resultado = ResultadoOP()

    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        wb = excel.Workbooks.Open(str(salida_xlsx.resolve()))
        try:
            ws = wb.Worksheets(_HOJA)

            # El machote trae la fecha del ejemplo tal cual (valor fijo, no
            # =TODAY()); debe ser la misma fecha de contratación del CONTRATO
            # (verificado con los ejemplos terminados: OP y contrato coinciden),
            # que es la fecha de cotejo del INE (ver cli.py).
            ws.Range("X6").Value = _dt.datetime.combine(
                fecha_contratacion or _dt.date.today(), _dt.time()
            )

            _set_con_ajuste(ws, "C9", cliente.razon_social)
            ws.Range("J9").Value = cliente.numero_cuenta
            ws.Range("N9").Value = cliente.rfc
            _set_con_ajuste(
                ws, "N12", f"{cliente.tipo_identificacion} //{cliente.numero_identificacion}"
            )
            _set_con_ajuste(ws, "S12", _domicilio_entrega(cliente))

            if oferta_comercial and oferta_comercial.folio_producto:
                folio_autorizacion = oferta_comercial.folio_producto
            else:
                folio_autorizacion, alerta_folio = _folio_autorizacion(cliente)
                if alerta_folio:
                    resultado.alertas.append(alerta_folio)
            ws.Range("G12").Value = folio_autorizacion

            for nombre_shape in _SHAPES_MESES_GRATIS:
                ws.Shapes(nombre_shape).TextFrame2.TextRange.Text = ""
            if oferta_comercial and len(oferta_comercial.meses_renta_gratis) == 3:
                for nombre_shape, mes in zip(
                    _SHAPES_MESES_GRATIS, oferta_comercial.meses_renta_gratis
                ):
                    ws.Shapes(nombre_shape).TextFrame2.TextRange.Text = str(mes)

            for nombre_shape in _CHECKBOX_TIPO_VENTA.values():
                ws.Shapes(nombre_shape).TextFrame2.TextRange.Text = ""
            ws.Shapes(_CHECKBOX_TIPO_VENTA[tipo_venta]).TextFrame2.TextRange.Text = "X"

            nombre_firma = cliente.representante_legal or "N/A"
            ws.Shapes(_TEXTBOX_FIRMA_SUSCRIPTOR).TextFrame2.TextRange.Text = nombre_firma
            if not cliente.representante_legal:
                resultado.alertas.append(
                    "No hay Representante Legal del cliente para la firma del "
                    "suscriptor en la OP; se dejó 'N/A'."
                )

            bloques_redes = _redes_streaming_por_plan(cliente)
            if bloques_redes:
                for celda, bloque in zip(_FILAS_OBSERVACIONES_REDES, bloques_redes):
                    _set_con_ajuste(ws, celda, bloque)
                if len(bloques_redes) > len(_FILAS_OBSERVACIONES_REDES):
                    resultado.alertas.append(
                        f"El cliente tiene {len(bloques_redes)} combinaciones de "
                        f"plan con REDES/STREAMING distintas, pero solo hay "
                        f"{len(_FILAS_OBSERVACIONES_REDES)} filas disponibles en "
                        f"Observaciones; se omitieron: "
                        f"{bloques_redes[len(_FILAS_OBSERVACIONES_REDES):]}"
                    )
            else:
                resultado.alertas.append(
                    "No se encontraron 'REDES SOCIALES'/'STREAMING' en SAE para "
                    "completar la fila de Observaciones de la OP."
                )

            fila = _PRIMERA_FILA_TABLA
            for linea in cliente.lineas:
                ciudad_dn = ""
                if ladas_csv_path:
                    try:
                        ciudad_dn = ciudad_dn_por_telefono(linea.telefono, ladas_csv_path)
                    except LadaNoEncontradaError as e:
                        resultado.alertas.append(str(e))

                for col in _COLUMNAS_PENDIENTES:
                    ws.Range(f"{col}{fila}").Value = "N/A"

                ws.Range(f"C{fila}").Value = 1
                # Sin lada determinable (ej. SIM con teléfono "NA" por portabilidad)
                # se pone N/A en vez de dejar la celda vacía.
                _set_con_ajuste(ws, f"D{fila}", ciudad_dn or "N/A")
                _set_con_ajuste(ws, f"E{fila}", linea.plan_tarifario)
                ws.Range(f"F{fila}").Value = linea.plazo_meses
                _set_con_ajuste(ws, f"H{fila}", linea.marca_modelo_color)
                ws.Range(f"T{fila}").Value = linea.telefono
                ws.Range(f"U{fila}").Value = _modalidad_mpp_cpp(
                    linea.plan_tarifario, linea.modalidad_mpp_cpp
                )

                ws.Range(f"L{fila}").Value = linea.plazo_meses  # Cuotas = Plazo

                numero_plan = _numero_plan(linea.plan_tarifario)
                es_ctrl = _es_addon_ctrl(linea.plan_tarifario)

                mov_oferta = None
                if oferta_comercial:
                    mov_oferta = buscar_movimiento_para_linea(
                        linea.modelo, numero_plan, oferta_comercial.movimientos
                    )

                # Descuento multilínea unitario (OP col O) = "%DMR" de la
                # oferta comercial cuando la línea matchea un movimiento de
                # ese PDF; ese mismo % descuenta también el plan base y el
                # Addon CTRL en 'Pago total mensual unitario' (regla
                # descubierta 2026-09-22 comparando EJEMPLO 2 de OP con PDF
                # de oferta: plan*[1-DMR%] + Addon CTRL*[1-DMR%] = Pago
                # total mensual unitario; verificado exacto en 4 líneas).
                factor_dmr = 1.0
                if mov_oferta is not None:
                    ws.Range(f"O{fila}").Value = f"{mov_oferta.dmr_pct:g}%"
                    factor_dmr = 1 - mov_oferta.dmr_pct / 100

                servicio_adicional = 0.0
                if es_ctrl:
                    servicio_adicional = round(_ADDON_CTRL_PRECIO * factor_dmr, 2)
                    ws.Range(f"P{fila}").Value = "X"
                    ws.Range(f"S{fila}").Value = servicio_adicional

                if numero_plan is not None:
                    ws.Range(f"Y{fila}").Value = round(
                        numero_plan * factor_dmr + servicio_adicional, 2
                    )
                else:
                    resultado.alertas.append(
                        f"No se pudo extraer el número de plan de "
                        f"{linea.plan_tarifario!r} para calcular el Pago total "
                        f"mensual unitario; se dejó 'N/A'."
                    )

                diferencial_lista = None
                precio_lista_valor = None
                if lista_precios_xlsx and _trae_equipo(linea):
                    try:
                        precio_lista = buscar_precio_lista(
                            linea.modelo,
                            lista_precios_xlsx,
                            numero_plan=numero_plan,
                            plazo_meses=linea.plazo_meses,
                        )
                        ws.Range(f"J{fila}").Value = precio_lista.precio_lista
                        precio_lista_valor = precio_lista.precio_lista
                        diferencial_lista = precio_lista.diferencial_equipo
                    except PrecioNoEncontradoError as e:
                        ws.Range(f"J{fila}").Value = "N/A"
                        resultado.alertas.append(str(e))
                else:
                    ws.Range(f"J{fila}").Value = "N/A"

                # PIE (Pago inicial/Diferencial de equipo, OP col K):
                # prioridad 1) la oferta comercial (PDF "Formato de
                # Autorizaciones Especiales"), el monto "$... sin IVA c/u"
                # ya negociado — la fuente más autorizada cuando existe
                # (descubierto 2026-09-22: coincide exacto con 'Costo de
                # Equipo' de control de renovación en un caso real, lo que
                # confirma que ese campo se llena desde el mismo documento
                # cuando lo hay); 2) si no hay oferta comercial o la línea
                # no matcheó ningún movimiento, la lista de precios
                # (columna de plan+plazo, ver
                # catalogo_precios._columna_diferencial); 3) 'Costo de
                # Equipo' de control de renovación solo de último respaldo.
                # Si dos fuentes disponibles difieren, se alerta en vez de
                # elegir en silencio.
                pie = 0.0
                if not _trae_equipo(linea):
                    # SIM sin equipo: no tiene costo, el pago inicial no aplica
                    # (regla del usuario 2026-09-23) aunque la oferta traiga $0.00.
                    ws.Range(f"K{fila}").Value = "N/A"
                elif mov_oferta is not None:
                    pie = mov_oferta.precio_unitario_sin_iva
                    ws.Range(f"K{fila}").Value = pie
                    if linea.costo_equipo is not None and abs(linea.costo_equipo - pie) > 1:
                        resultado.alertas.append(
                            f"'Pago inicial/Diferencial de equipo' de la oferta "
                            f"comercial (${pie:,.2f}) no coincide con 'Costo de "
                            f"Equipo' de control de renovación "
                            f"(${linea.costo_equipo:,.2f}) para la línea "
                            f"{linea.telefono}; se usó el valor de la oferta "
                            f"comercial (documento formal, fuente más autorizada). "
                            f"Verificar manualmente."
                        )
                elif diferencial_lista is not None and diferencial_lista > 0:
                    pie = diferencial_lista
                    ws.Range(f"K{fila}").Value = pie
                    if (
                        linea.costo_equipo is not None
                        and linea.costo_equipo > 0
                        and abs(linea.costo_equipo - diferencial_lista) > 1
                    ):
                        resultado.alertas.append(
                            f"'Pago inicial/Diferencial de equipo' de la lista de "
                            f"precios (${diferencial_lista:,.2f}, plan "
                            f"{int(numero_plan) if numero_plan is not None else '?'} a "
                            f"{linea.plazo_meses} meses) no coincide con 'Costo de "
                            f"Equipo' de control de renovación "
                            f"(${linea.costo_equipo:,.2f}) para la línea "
                            f"{linea.telefono}; se usó el valor de la lista de "
                            f"precios (fuente confirmada). Verificar manualmente."
                        )
                elif linea.costo_equipo is not None and linea.costo_equipo > 0:
                    pie = linea.costo_equipo
                    ws.Range(f"K{fila}").Value = pie
                    resultado.alertas.append(
                        f"No se pudo obtener 'Pago inicial/Diferencial de equipo' "
                        f"de la lista de precios para la línea {linea.telefono} "
                        f"(plan/plazo fuera del catálogo conocido o equipo no "
                        f"encontrado); se usó como respaldo 'Costo de Equipo' de "
                        f"control de renovación (${linea.costo_equipo:,.2f}) — "
                        f"verificar manualmente."
                    )

                # MPE (col N). Criterio principal: la calculadora MPE. Criterio
                # alternativo (regla del usuario 2026-09-23, verificado en 37 de 38
                # líneas ya generadas): (Precio de lista - Diferencial de equipo) /
                # plazo con la misma fórmula de la calculadora. Se usa si la
                # calculadora falla, y si ambas difieren gana el precio de la LISTA
                # DE PRECIOS, que es el correcto según el usuario.
                plazo_int = int(linea.plazo_meses) if str(linea.plazo_meses).isdigit() else None
                mpe_lista = (
                    mpe_desde_lista(precio_lista_valor, pie, plazo_int)
                    if precio_lista_valor is not None and plazo_int
                    else None
                )
                mpe_calc = None
                error_calc = None
                if calculo_mpe_xlsx and plazo_int and _trae_equipo(linea):
                    try:
                        mpe_calc = calcular_mpe(linea.modelo, plazo_int, calculo_mpe_xlsx, pie=pie)
                    except (MPENoEncontradoError, ValueError) as e:
                        error_calc = e

                mpe_final = None
                if mpe_calc is not None and (mpe_lista is None or abs(mpe_calc.mpe - mpe_lista) < 0.005):
                    mpe_final = mpe_calc.mpe
                elif mpe_calc is not None:
                    mpe_final = mpe_lista
                    resultado.alertas.append(
                        f"'{linea.modelo}' (línea {linea.telefono}): el Precio de lista de la "
                        f"lista de precios (${precio_lista_valor:,.2f}) no coincide con el "
                        f"Precio Base del cálculo de MPE (${mpe_calc.precio_base:,.2f}); el MPE "
                        f"se calculó con el precio de la LISTA DE PRECIOS ({mpe_lista}) por ser "
                        f"el correcto (con el de la calculadora sería {mpe_calc.mpe}). "
                        f"Consultar por qué difieren."
                    )
                elif mpe_lista is not None:
                    mpe_final = mpe_lista
                    resultado.alertas.append(
                        f"La calculadora MPE no dio resultado para '{linea.modelo}' (línea "
                        f"{linea.telefono}): {error_calc or 'sin archivo de calculadora'}. El MPE "
                        f"se calculó con el criterio alternativo (Precio de lista - Diferencial "
                        f"de equipo) / plazo = {mpe_lista}."
                    )
                elif error_calc is not None:
                    resultado.alertas.append(str(error_calc))

                if mpe_final is not None and mpe_final < 0:
                    resultado.alertas.append(
                        f"MPE calculado negativo ({mpe_final}) para '{linea.modelo}' (línea "
                        f"{linea.telefono}); se dejó N/A, consultar por qué los datos no cuadran."
                    )
                    mpe_final = None
                ws.Range(f"N{fila}").Value = mpe_final if mpe_final is not None else "N/A"

                fila += 1
                resultado.lineas_escritas += 1

            # Limpiar cualquier fila sobrante que el machote reciclado traía
            # con datos de un cliente anterior (más líneas de las que trae
            # este cliente).
            for fila_sobrante in range(fila, _ULTIMA_FILA_TABLA + 1):
                for col in _TODAS_LAS_COLUMNAS_TABLA:
                    ws.Range(f"{col}{fila_sobrante}").Value = None

            if resultado.lineas_escritas:
                if oferta_comercial and oferta_comercial.movimientos:
                    resultado.alertas.append(
                        "Addon's de datos, Protección de equipo y Otros "
                        "servicios se dejaron como 'N/A' (todavía no se "
                        "identificó su fórmula/fuente) — completar manualmente "
                        "si aplican. Precio de lista, Cuotas, MPE, Pago inicial, "
                        "Descuento multilínea, Addon CTRL y Pago total mensual "
                        "sí se calcularon automáticamente (usando la oferta "
                        "comercial donde matcheó): verificar que tengan sentido."
                    )
                else:
                    resultado.alertas.append(
                        "Descuento multilínea, Addon's de datos, Protección de "
                        "equipo y Otros servicios se dejaron como 'N/A' (todavía "
                        "no se identificó su fórmula/fuente, o no llegó oferta "
                        "comercial) — completar manualmente si aplican. Precio "
                        "de lista, Cuotas, MPE, Pago inicial, Addon CTRL y Pago "
                        "total mensual sí se calcularon automáticamente: "
                        "verificar que tengan sentido."
                    )

            wb.Save()
        finally:
            wb.Close(False)
    finally:
        excel.Quit()

    return resultado
