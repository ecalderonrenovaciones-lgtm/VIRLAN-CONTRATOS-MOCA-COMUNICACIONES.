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
  X6  Fecha (trae por defecto =TODAY() en el machote; se sobreescribe con la
      fecha de contratación)

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
  K Pago inicial/Diferencial de equipo (PIE) = 'Costo de Equipo' de control
    de renovacion.xlsx, SOLO si es mayor a 0 (regla confirmada por el
    usuario); si es 0 o no hay dato, se deja "N/A".
  N MPE (calculadora_mpe.calcular_mpe) — usa como PIE el mismo valor de la
    columna K (antes se mandaba PIE=0 siempre, lo cual el usuario confirmó
    que estaba mal: el PIE del MPE debe ser el mismo 'Diferencial de
    equipo unitario' de la columna K).
  P Addon CTRL = "X" y S Pago mensual de servicios adicionales = 50,
    SOLO si el nombre del Plan tarifario trae el sufijo "CTRL" (regla
    confirmada por el usuario; precio fijo de $50).
  Y Pago total mensual unitario = número del plan (299/399/.../1499,
    extraído del nombre del Plan tarifario) + 50 si aplica Addon CTRL.
  Fila 31 (Total a Pagar) ya trae la fórmula =SUM(Y16:Y30) en el machote:
    basta con llenar Y correctamente, no hace falta tocarla.
Observaciones (celdas C35:C37, " PLAN <número>: REDES: ... STREAMING: ..."):
'REDES SOCIALES' y 'STREAMING' de SAE 2.2.xlsx varían por plan tarifario
(regla confirmada por el usuario), así que se escribe una fila por cada
combinación de plan distinta presente en el cliente (máximo 3, las únicas
filas libres que deja el machote antes de la tabla de Abrev./Addón).
Firma del suscriptor (shape 'TextBox 5'): siempre el Representante Legal
del cliente (ficha .docx), nunca el nombre reciclado del machote.
Las columnas que aún no tienen fórmula identificada (O Descuento
multilínea, Q Addon's de datos, V Protección de equipo, X Otros servicios)
y "Meses Gratis" se dejan como "N/A" / en blanco."""

from __future__ import annotations

import re
import shutil
from dataclasses import dataclass, field as dc_field
from pathlib import Path

import win32com.client as win32

from .calculadora_mpe import EquipoNoEncontradoError as MPENoEncontradoError
from .calculadora_mpe import calcular_mpe
from .catalogo_precios import EquipoNoEncontradoError as PrecioNoEncontradoError
from .catalogo_precios import buscar_precio_lista
from .ladas_lookup import LadaNoEncontradaError, ciudad_dn_por_telefono
from .models import ClienteContrato

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

            ws.Range("C9").Value = cliente.razon_social
            ws.Range("J9").Value = cliente.numero_cuenta
            ws.Range("N9").Value = cliente.rfc
            ws.Range("N12").Value = f"{cliente.tipo_identificacion} //{cliente.numero_identificacion}"
            ws.Range("S12").Value = _domicilio_entrega(cliente)

            # El machote es un archivo reciclado de un cliente anterior; el
            # Folio de Autorización que trae por defecto NO corresponde a
            # este cliente y no hay todavía una fuente identificada para el
            # de este cliente, así que se limpia en vez de dejar un folio
            # ajeno que parezca válido.
            ws.Range("G12").Value = None
            resultado.alertas.append(
                "Folio de Autorización se dejó en blanco (el machote traía "
                "el de un cliente anterior y no hay fuente confirmada para "
                "el de este cliente todavía) — completar manualmente."
            )

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
                    ws.Range(celda).Value = bloque
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
                ws.Range(f"D{fila}").Value = ciudad_dn
                ws.Range(f"E{fila}").Value = linea.plan_tarifario
                ws.Range(f"F{fila}").Value = linea.plazo_meses
                ws.Range(f"H{fila}").Value = linea.marca_modelo_color
                ws.Range(f"T{fila}").Value = linea.telefono
                ws.Range(f"U{fila}").Value = _modalidad_mpp_cpp(
                    linea.plan_tarifario, linea.modalidad_mpp_cpp
                )

                ws.Range(f"L{fila}").Value = linea.plazo_meses  # Cuotas = Plazo

                pie = 0.0
                if linea.costo_equipo is not None and linea.costo_equipo > 0:
                    pie = linea.costo_equipo
                    ws.Range(f"K{fila}").Value = pie

                es_ctrl = _es_addon_ctrl(linea.plan_tarifario)
                if es_ctrl:
                    ws.Range(f"P{fila}").Value = "X"
                    ws.Range(f"S{fila}").Value = _ADDON_CTRL_PRECIO

                numero_plan = _numero_plan(linea.plan_tarifario)
                if numero_plan is not None:
                    ws.Range(f"Y{fila}").Value = numero_plan + (
                        _ADDON_CTRL_PRECIO if es_ctrl else 0
                    )
                else:
                    resultado.alertas.append(
                        f"No se pudo extraer el número de plan de "
                        f"{linea.plan_tarifario!r} para calcular el Pago total "
                        f"mensual unitario; se dejó 'N/A'."
                    )

                if lista_precios_xlsx:
                    try:
                        precio_lista = buscar_precio_lista(linea.modelo, lista_precios_xlsx)
                        ws.Range(f"J{fila}").Value = precio_lista.precio_lista
                        if precio_lista.es_sustituto:
                            resultado.alertas.append(
                                f"'{linea.modelo}' no está en la lista de precios "
                                f"vigente; se usó el sustituto más similar con la "
                                f"misma capacidad y conectividad: "
                                f"'{precio_lista.familia}' (${precio_lista.precio_lista}). "
                                f"Verificar manualmente antes de enviar (línea "
                                f"{linea.telefono})."
                            )
                    except PrecioNoEncontradoError as e:
                        ws.Range(f"J{fila}").Value = "N/A"
                        resultado.alertas.append(str(e))
                else:
                    ws.Range(f"J{fila}").Value = "N/A"

                if calculo_mpe_xlsx and linea.plazo_meses:
                    try:
                        mpe = calcular_mpe(
                            linea.modelo, int(linea.plazo_meses), calculo_mpe_xlsx, pie=pie
                        )
                        ws.Range(f"N{fila}").Value = mpe.mpe
                        if mpe.es_sustituto:
                            resultado.alertas.append(
                                f"'{linea.modelo}' no está en el cálculo de MPE; se "
                                f"usó el sustituto más similar con la misma "
                                f"capacidad y conectividad (Precio Base "
                                f"${mpe.precio_base}, fila {mpe.fila} de "
                                f"'{Path(calculo_mpe_xlsx).name}'). Verificar "
                                f"manualmente antes de enviar (línea {linea.telefono})."
                            )
                    except (MPENoEncontradoError, ValueError) as e:
                        ws.Range(f"N{fila}").Value = "N/A"
                        resultado.alertas.append(str(e))
                else:
                    ws.Range(f"N{fila}").Value = "N/A"

                fila += 1
                resultado.lineas_escritas += 1

            # Limpiar cualquier fila sobrante que el machote reciclado traía
            # con datos de un cliente anterior (más líneas de las que trae
            # este cliente).
            for fila_sobrante in range(fila, _ULTIMA_FILA_TABLA + 1):
                for col in _TODAS_LAS_COLUMNAS_TABLA:
                    ws.Range(f"{col}{fila_sobrante}").Value = None

            if resultado.lineas_escritas:
                resultado.alertas.append(
                    "Descuento multilínea, Addon's de datos, Protección de "
                    "equipo y Otros servicios se dejaron como 'N/A' (todavía "
                    "no se identificó su fórmula/fuente) — completar "
                    "manualmente si aplican. Precio de lista, Cuotas, MPE, "
                    "Pago inicial, Addon CTRL y Pago total mensual sí se "
                    "calcularon automáticamente: verificar que tengan sentido."
                )

            wb.Save()
        finally:
            wb.Close(False)
    finally:
        excel.Quit()

    return resultado
