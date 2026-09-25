import datetime as dt

import pytest

from virlan_bot.contrato_fieldmap import cargar_fieldmap
from virlan_bot.contrato_valores import _fecha_espaciada, abreviar_estado, construir_valores_contrato

FECHA = dt.date(2026, 9, 24)


def valores(cliente, **kw):
    return construir_valores_contrato(cliente, fecha_contratacion=FECHA, **kw)


@pytest.mark.parametrize(
    "estado,esperado",
    [("CIUDAD DE MEXICO", "CDMX"), ("Ciudad de México", "CDMX"), ("ESTADO DE MEXICO", "EDOMEX"),
     ("JALISCO", "JALISCO"), ("", ""), ("SAN LUIS POTOSI", "SAN LUIS POTOSI")],
)
def test_abreviar_estado(estado, esperado):
    assert abreviar_estado(estado) == esperado


def test_fechas_con_el_espaciado_exacto_del_ejemplo():
    assert _fecha_espaciada(dt.date(2026, 9, 23)) == "23" + " " * 10 + "09" + " " * 10 + "2026"
    assert _fecha_espaciada(dt.date(2026, 10, 8), (9, 7)) == "08" + " " * 9 + "10" + " " * 7 + "2026"


def test_valores_basicos(cliente):
    v, alertas = valores(cliente)
    assert v["numero_cuenta"] == "500000001"
    assert v["fecha_contratacion"] == _fecha_espaciada(FECHA)
    # entrega máxima = contratación + 14 días naturales (09-24 -> 10-08)
    assert v["fecha_maxima_entrega"] == _fecha_espaciada(dt.date(2026, 10, 8), (9, 7))
    assert v["plazo_meses"] == "24"
    assert "hora_entrega" not in v            # constante del machote
    assert v["representante_legal_1"] == v["representante_legal_2"] == "JUAN PEREZ LOPEZ"
    # sin persona indicada: el representante legal, con alerta
    assert v["persona_autorizada_recibir_equipos"] == "JUAN PEREZ LOPEZ"
    assert any("persona autorizada" in a for a in alertas)


def test_domicilio_en_casillas_separadas_y_estado_abreviado(cliente):
    v, _ = valores(cliente)
    assert v["domicilio_calle_numero"] == "CALLE UNO, 12"
    assert v["domicilio_colonia"] == "COL CENTRO"
    assert v["domicilio_municipio"] == "CUAUHTEMOC"
    assert v["domicilio_estado"] == "CDMX"
    assert v["domicilio_cp"] == "06000"
    assert "domicilio_ciudad_estado_cp" not in v           # ya no existe en el formato nuevo
    assert not any(k.startswith("domicilio_entrega") for k in v)  # envío FISCAL: fila de entrega vacía


def test_envio_a_entrega_llena_la_fila_de_abajo(cliente):
    cliente.envio_a = "ENTREGA"
    cliente.domicilio_entrega_calle = "AV DOS"
    cliente.domicilio_entrega_numero = "45"
    cliente.domicilio_entrega_colonia = "COL NORTE"
    cliente.domicilio_entrega_ciudad = "GUADALAJARA"
    cliente.domicilio_entrega_estado = "JALISCO"
    cliente.domicilio_entrega_cp = "44100"
    v, _ = valores(cliente)
    assert v["domicilio_entrega_calle_numero"] == "AV DOS, 45"
    assert v["domicilio_entrega_estado"] == "JALISCO" and v["domicilio_entrega_cp"] == "44100"
    assert v["telefono_entrega"] == cliente.telefono        # mismo teléfono del cliente


def test_persona_autorizada_indicada_no_genera_alerta(cliente):
    v, alertas = valores(cliente, persona_autorizada_recibir_equipos="ANA LOPEZ Y LUIS DIAZ")
    assert v["persona_autorizada_recibir_equipos"] == "ANA LOPEZ Y LUIS DIAZ"
    assert not any("persona autorizada" in a for a in alertas)


def test_vendedor_predeterminado_es_prime(cliente):
    v, _ = valores(cliente)
    assert v["nombre_ejecutivo"] == "SERGIO YAHIR VALDERRAMA VELOZ" and v["rfc_ejecutivo"] == "VAVS020813"


def test_vendedor_de_otro_canal_reemplaza_todo(cliente):
    from virlan_bot import config

    v, alertas = valores(cliente, vendedor=dict(config.VENDEDORES_POR_CANAL["ONE STOP"]))
    assert v["nombre_ejecutivo"] == "ALEJANDRO CASTERA ARELLANOS"
    assert v["rfc_ejecutivo"] == "CAAA760226LJ4"
    assert "punto_venta_codigo" not in v                     # "" explícito: en blanco a propósito
    assert not any("Vendedor distinto" in a for a in alertas)


def test_vendedor_incompleto_nunca_reutiliza_el_rfc_de_otro(cliente):
    v, alertas = valores(cliente, vendedor={"nombre_ejecutivo": "OTRO EJECUTIVO"})
    assert v["nombre_ejecutivo"] == "OTRO EJECUTIVO"
    assert "rfc_ejecutivo" not in v and "punto_venta_nombre" not in v   # NO se hereda a Sergio
    assert any("Vendedor distinto" in a for a in alertas)


def test_sin_lineas_o_plazos_distintos_avisa(cliente):
    cliente.lineas = []
    v, alertas = valores(cliente)
    assert "plazo_meses" not in v and any("Sin líneas" in a for a in alertas)


def test_todas_las_claves_de_valores_existen_en_el_mapa_de_campos(cliente):
    """Detecta nombres que el filler ignoraría en silencio (campos huérfanos)."""
    cliente.envio_a = "ENTREGA"
    cliente.domicilio_entrega_calle = "AV DOS"
    v, _ = valores(cliente, persona_autorizada_recibir_equipos="X Y")
    campos = set(cargar_fieldmap("v2")["campos_texto"])
    assert set(v) <= campos, f"claves sin campo en el mapa: {set(v) - campos}"
