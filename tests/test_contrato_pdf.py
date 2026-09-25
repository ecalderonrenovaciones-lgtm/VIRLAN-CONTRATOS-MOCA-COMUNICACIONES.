"""Llenado real del CONTRATO (nueva versión) sobre el machote en blanco."""

import re
from collections import Counter

import fitz
import pytest

from virlan_bot import config
from virlan_bot.contrato_fieldmap import cargar_fieldmap
from virlan_bot.contrato_filler import llenar_contrato
from virlan_bot.contrato_valores import construir_valores_contrato
import datetime as dt

pytestmark = pytest.mark.machote

ETIQUETAS = ("Estado", "Teléfono", "Colonia", "Correspondencia", "C.P.", "Municipio", "Del./Municipio")


def norm(texto):
    return re.sub(r"\s+", " ", texto)


def generar(tmp_path, cliente, checkboxes, persona=None, vendedor=None):
    v, _ = construir_valores_contrato(
        cliente, persona_autorizada_recibir_equipos=persona, fecha_contratacion=dt.date(2026, 9, 24), vendedor=vendedor
    )
    salida = tmp_path / "contrato.pdf"
    fm = cargar_fieldmap(config.CONTRATO_FIELDMAP_VERSION)
    res = llenar_contrato(config.MACHOTE_CONTRATO_PDF, salida, fm, v, cliente.rfc or None, checkboxes)
    return salida, res


def rotulos(pagina):
    return Counter(w[4] for w in pagina.get_text("words") if w[4] in ETIQUETAS)


def test_machote_es_la_nueva_version():
    doc = fitz.open(config.MACHOTE_CONTRATO_PDF)
    assert len(doc) == 3
    assert "PROFECO" in doc[2].get_text() and "426-2026" in doc[2].get_text()


def test_contrato_fiscal_completo(tmp_path, cliente):
    pdf, res = generar(tmp_path, cliente, ["tipo_contratacion_renovacion", "domicilio_fiscal_es_dom_entrega",
                                           "domicilio_fiscal_es_correspondencia"])
    assert res.campos_omitidos == [] and res.desbordes == []
    t = norm(fitz.open(pdf)[0].get_text())
    for esperado in ("EMPRESA PRUEBA SA DE CV", "500000001", "JUAN PEREZ LOPEZ", "CALLE UNO, 12", "COL CENTRO",
                     "CUAUHTEMOC", "CDMX", "06000", "5512345678", "juan@prueba.mx", "24 09 2026", "08 10 2026",
                     "SERGIO YAHIR VALDERRAMA VELOZ", "VAVS020813", "INE", "1234567890"):
        assert esperado in t, esperado
    assert not re.search(r"\bSra?\.", t)                    # sin títulos


def test_no_quedan_datos_de_muestra_del_ejemplo(tmp_path, cliente):
    pdf, _ = generar(tmp_path, cliente, ["tipo_contratacion_renovacion"])
    t = norm(fitz.open(pdf)[0].get_text())
    for muestra in ("RAZO ESTRADA", "PUERTO MANZANILLO", "JUAN CARLOS", "8682031196", "gerardo_auto", "507203414"):
        assert muestra not in t


def test_los_rotulos_del_formato_no_se_borran(tmp_path, cliente):
    """Regresión: el borrado subía ~3 pt y se llevaba 'Estado'/'Teléfono'."""
    machote = fitz.open(config.MACHOTE_CONTRATO_PDF)[0]
    pdf, _ = generar(tmp_path, cliente, ["tipo_contratacion_renovacion", "domicilio_fiscal_es_dom_entrega"])
    assert rotulos(fitz.open(pdf)[0]) == rotulos(machote)


def contar_x(pagina, y0, y1, x0=0, x1=612):
    return [
        w for w in pagina.get_text("words") if w[4] == "X" and y0 <= w[1] <= y1 and x0 <= w[0] <= x1
    ]


def test_una_sola_x_de_tipo_de_contratacion_en_su_casilla(tmp_path, cliente):
    for checkbox, x_esperado in (("tipo_contratacion_renovacion", 416.9), ("tipo_contratacion_adicion", 131.9),
                                 ("tipo_contratacion_suscripcion_nueva", 27.9)):
        pdf, _ = generar(tmp_path, cliente, [checkbox])
        xs = contar_x(fitz.open(pdf)[0], 80, 96)
        assert len(xs) == 1 and abs(xs[0][0] - x_esperado) < 1.5, (checkbox, xs)


def test_envio_fiscal_marca_las_dos_x_de_arriba_y_ninguna_de_abajo(tmp_path, cliente):
    pdf, _ = generar(tmp_path, cliente, ["tipo_contratacion_renovacion", "domicilio_fiscal_es_dom_entrega",
                                         "domicilio_fiscal_es_correspondencia"])
    pg = fitz.open(pdf)[0]
    assert len(contar_x(pg, 230, 246)) == 2
    assert contar_x(pg, 253, 266) == []


def test_envio_a_entrega_quita_las_de_arriba_y_marca_la_de_abajo(tmp_path, cliente):
    cliente.envio_a = "ENTREGA"
    cliente.domicilio_entrega_calle = "AV DOS"
    cliente.domicilio_entrega_numero = "45"
    cliente.domicilio_entrega_colonia = "COL NORTE"
    cliente.domicilio_entrega_ciudad = "GUADALAJARA"
    cliente.domicilio_entrega_estado = "JALISCO"
    cliente.domicilio_entrega_cp = "44100"
    pdf, _ = generar(tmp_path, cliente, ["tipo_contratacion_renovacion", "domicilio_entrega_es_correspondencia"])
    pg = fitz.open(pdf)[0]
    assert contar_x(pg, 230, 246) == []
    assert len(contar_x(pg, 253, 266)) == 1
    t = norm(pg.get_text())
    for esperado in ("AV DOS, 45", "COL NORTE", "GUADALAJARA", "JALISCO", "44100"):
        assert esperado in t


@pytest.mark.parametrize("rfc", ["EPR010101AB1", "PERJ800101AB9"])   # 12 (moral) y 13 (física)
def test_rfc_un_caracter_por_casilla(tmp_path, cliente, rfc):
    cliente.rfc = rfc
    pdf, res = generar(tmp_path, cliente, ["tipo_contratacion_renovacion"])
    letras = [w for w in fitz.open(pdf)[0].get_text("words") if 199 <= w[1] <= 212 and w[0] >= 440 and len(w[4]) == 1]
    assert "".join(w[4] for w in sorted(letras, key=lambda w: w[0])) == rfc
    assert res.desbordes == []


def test_campo_sin_valor_queda_en_blanco_y_avisa(tmp_path, cliente):
    cliente.telefono = ""
    pdf, res = generar(tmp_path, cliente, ["tipo_contratacion_renovacion"])
    assert {"telefono_1", "telefono_2"} <= set(res.campos_omitidos)
    assert "5512345678" not in norm(fitz.open(pdf)[0].get_text())


def test_texto_largo_se_reduce_hasta_caber(tmp_path, cliente):
    cliente.domicilio_calle = "AV REVOLUCION Y BLVD DIAZ ORDAZ LOCAL 25 Y 26 LETRA F ALTOS"
    cliente.domicilio_colonia = "FRACC RESIDENCIAL PASEO DE LAS PALMAS"
    pdf, res = generar(tmp_path, cliente, ["tipo_contratacion_renovacion"])
    assert res.desbordes == []
    t = norm(fitz.open(pdf)[0].get_text())
    assert "LETRA F ALTOS" in t and "FRACC RESIDENCIAL PASEO DE LAS PALMAS" in t


def test_texto_imposible_de_ajustar_se_avisa_y_no_se_recorta(tmp_path, cliente):
    """Nunca se trunca en silencio: si no cabe ni a 4 pt, queda íntegro y con alerta."""
    largo = "FRACCIONAMIENTO RESIDENCIAL PASEO DE LAS PALMAS NORTE"
    cliente.domicilio_colonia = largo
    pdf, res = generar(tmp_path, cliente, ["tipo_contratacion_renovacion"])
    assert any("domicilio_colonia" in d for d in res.desbordes)
    assert largo in norm(fitz.open(pdf)[0].get_text())


def test_persona_autorizada_sin_dos_puntos(tmp_path, cliente):
    pdf, _ = generar(tmp_path, cliente, ["tipo_contratacion_renovacion"], persona="ANA LOPEZ Y LUIS DIAZ")
    lineas = [l for l in fitz.open(pdf)[0].get_text().split("\n") if "ANA" in l and "LOPEZ" in l]
    assert lineas and not any(l.strip().startswith(":") for l in lineas)


def test_vendedor_one_stop_en_el_contrato(tmp_path, cliente):
    pdf, _ = generar(tmp_path, cliente, ["tipo_contratacion_adicion"], vendedor=dict(config.VENDEDORES_POR_CANAL["ONE STOP"]))
    t = norm(fitz.open(pdf)[0].get_text())
    assert "ONE STOP MARKET AC425E" in t and "ALEJANDRO CASTERA ARELLANOS" in t and "CAAA760226LJ4" in t
    assert "SERGIO YAHIR" not in t and "VAVS020813" not in t


def test_paginas_2_y_3_no_se_tocan_y_la_fuente_esta_incrustada(tmp_path, cliente):
    pdf, _ = generar(tmp_path, cliente, ["tipo_contratacion_renovacion"])
    doc, base = fitz.open(pdf), fitz.open(config.MACHOTE_CONTRATO_PDF)
    assert len(doc) == 3
    assert doc[1].get_text() == base[1].get_text() and doc[2].get_text() == base[2].get_text()
    # lo que escribe el bot usa Arial INCRUSTADA (editable en Acrobat), no Helvetica base-14
    incrustadas = {f[3] for f in doc[0].get_fonts() if f[1] == "ttf" and "Arial" in f[3]}
    assert len(incrustadas) >= 2
    assert not any(f[3] in ("ArialInc", "ArialBoldInc") and f[1] == "n/a" for f in doc[0].get_fonts())
