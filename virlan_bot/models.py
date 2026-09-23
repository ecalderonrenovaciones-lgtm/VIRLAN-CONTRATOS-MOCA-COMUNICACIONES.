"""Modelo de datos intermedio: representa toda la información de un cliente
ya normalizada, lista para alimentar tanto el llenado del CONTRATO como el
de la ORDEN DE PROGRAMACIÓN (OP)."""

from __future__ import annotations

from dataclasses import dataclass, field


class DatosIncompletosError(Exception):
    """Se lanza cuando un extractor no puede obtener un dato requerido y no
    debe adivinarse. El pipeline debe detenerse y pedir revisión humana."""


@dataclass
class LineaRenovacion:
    """Una línea/equipo a renovar, tomada de SAE 2.2.xlsx o control de renovacion.xlsx."""

    telefono: str
    plan_tarifario: str
    plazo_meses: str
    modelo: str
    marca_modelo_color: str
    modalidad_mpp_cpp: str | None = None
    addon_extra: str | None = None
    costo_equipo: float | None = None  # 'Costo de Equipo' de control de renovacion.xlsx
    redes_sociales: str | None = None  # 'REDES SOCIALES' de SAE
    streaming: str | None = None  # 'STREAMING' de SAE
    ciudad_dn: str | None = None  # se completa con ladas_lookup.py
    fuente: str = ""  # nombre del archivo del que se obtuvo la fila


@dataclass
class ClienteContrato:
    """Datos normalizados de un cliente, listos para llenar CONTRATO y OP."""

    # De la ficha .docx
    razon_social: str = ""
    numero_cuenta: str = ""
    representante_legal: str = ""
    domicilio_calle: str = ""
    domicilio_numero: str = ""
    domicilio_colonia: str = ""
    domicilio_ciudad: str = ""
    domicilio_municipio: str = ""
    domicilio_estado: str = ""
    domicilio_cp: str = ""
    domicilio_pais: str = ""
    telefono: str = ""
    correo: str = ""

    # A qué domicilio se envía el paquete, según la línea "EL PAQUETE SE ENVIA
    # A ESTE DOMICILIO:<FISCAL|ENTREGA>" de la ficha (regla del usuario
    # 2026-09-23). "ENTREGA" => el CONTRATO llena también la fila de Domicilio
    # de Entrega y la OP usa ese domicilio como Dirección de Entrega.
    envio_a: str = "FISCAL"

    # Domicilio de entrega (bloque "DOMICILIO DE ENTREGA" de la ficha). Mismos
    # 8 componentes que el fiscal: calle, número, colonia, ciudad, municipio,
    # estado, CP, país.
    domicilio_entrega_raw: str = ""
    domicilio_entrega_calle: str = ""
    domicilio_entrega_numero: str = ""
    domicilio_entrega_colonia: str = ""
    domicilio_entrega_ciudad: str = ""
    domicilio_entrega_municipio: str = ""
    domicilio_entrega_estado: str = ""
    domicilio_entrega_cp: str = ""
    domicilio_entrega_pais: str = ""

    # De LAYOUT DE VINCULACION (join por numero_cuenta)
    rfc: str = ""
    tipo_identificacion: str = ""
    numero_identificacion: str = ""

    # De SAE / control de renovacion (join por numero_cuenta)
    lineas: list[LineaRenovacion] = field(default_factory=list)

    # Trazabilidad y control de calidad
    alertas: list[str] = field(default_factory=list)
    campos_faltantes: list[str] = field(default_factory=list)
    origen: dict[str, str] = field(default_factory=dict)

    def domicilio_calle_numero(self) -> str:
        return ", ".join(p for p in (self.domicilio_calle, self.domicilio_numero) if p)

    def domicilio_ciudad_estado_cp(self) -> str:
        partes = [self.domicilio_ciudad, self.domicilio_estado, self.domicilio_cp]
        return ", ".join(p for p in partes if p)

    def entrega_calle_numero(self) -> str:
        return ", ".join(p for p in (self.domicilio_entrega_calle, self.domicilio_entrega_numero) if p)

    def entrega_ciudad_estado_cp(self) -> str:
        partes = [self.domicilio_entrega_ciudad, self.domicilio_entrega_estado, self.domicilio_entrega_cp]
        return ", ".join(p for p in partes if p)

    def agregar_alerta(self, mensaje: str) -> None:
        self.alertas.append(mensaje)
