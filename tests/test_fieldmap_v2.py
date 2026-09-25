"""Coherencia geométrica del mapa de campos (detecta errores de calibración sin abrir un PDF lleno)."""

import fitz
import pytest

from virlan_bot import config
from virlan_bot.contrato_fieldmap import cargar_fieldmap

FM = cargar_fieldmap("v2")
ANCHO, ALTO = 612, 792


def test_configuracion_apunta_a_la_version_2():
    assert FM["sin_redaccion"] is True and FM["pagina"] == 0
    assert FM["checkboxes_siempre"] == []


@pytest.mark.parametrize("nombre,spec", list(FM["campos_texto"].items()))
def test_cada_campo_esta_dentro_de_la_pagina_y_bien_formado(nombre, spec):
    x0, y0, x1, y1 = spec["bbox"]
    assert 0 <= x0 < x1 <= ANCHO and 0 <= y0 < y1 <= ALTO, nombre
    assert x1 - x0 >= 20, f"{nombre}: ancho disponible demasiado chico"
    assert spec["font"] in ("ArialMT", "Arial-BoldMT") and 6 <= spec["size"] <= 12


def test_las_lineas_base_de_filas_hermanas_coinciden():
    """Fiscal: mismo renglón; entrega: 21 pt más abajo (divisiones dibujadas)."""
    def base(n):
        b = FM["campos_texto"][n]
        return round(b["bbox"][3] - 0.15 * b["size"], 1)

    fiscal = {base(n) for n in ("domicilio_calle_numero", "domicilio_colonia", "domicilio_municipio",
                                "domicilio_estado", "domicilio_cp")}
    entrega = {base(n) for n in ("domicilio_entrega_calle_numero", "domicilio_entrega_colonia",
                                 "domicilio_entrega_municipio", "domicilio_entrega_estado", "domicilio_entrega_cp")}
    assert fiscal == {251.7} and entrega == {272.7}
    assert base("telefono_entrega") - base("telefono_2") == pytest.approx(21.0)


def test_las_casillas_de_una_fila_no_se_solapan_en_x():
    for fila in (("domicilio_calle_numero", "domicilio_colonia", "domicilio_municipio", "domicilio_estado",
                  "domicilio_cp", "telefono_2"),
                 ("domicilio_entrega_calle_numero", "domicilio_entrega_colonia", "domicilio_entrega_municipio",
                  "domicilio_entrega_estado", "domicilio_entrega_cp", "telefono_entrega")):
        cajas = sorted((FM["campos_texto"][n]["bbox"][0], FM["campos_texto"][n]["bbox"][2], n) for n in fila)
        for (_, x1a, na), (x0b, _, nb) in zip(cajas, cajas[1:]):
            assert x1a <= x0b, f"{na} invade a {nb}"


def test_los_campos_de_entrega_son_opcionales():
    for n, spec in FM["campos_texto"].items():
        assert bool(spec.get("opcional")) == (n.startswith("domicilio_entrega") or n == "telefono_entrega"), n


def test_checkboxes_con_origen_y_marca():
    assert set(FM["checkboxes"]) == {
        "tipo_contratacion_suscripcion_nueva", "tipo_contratacion_adicion", "tipo_contratacion_renovacion",
        "domicilio_fiscal_es_dom_entrega", "domicilio_fiscal_es_correspondencia",
        "domicilio_entrega_es_correspondencia",
    }
    for n, cb in FM["checkboxes"].items():
        x, y = cb["origen"]
        assert 0 < x < ANCHO and 0 < y < ALTO and cb["marca"] == "X", n


@pytest.mark.machote
def test_cada_casilla_de_x_cae_dentro_de_una_casilla_dibujada_del_machote():
    """La X debe quedar dentro de un rectángulo relleno (gris) del formato."""
    pagina = fitz.open(config.MACHOTE_CONTRATO_PDF)[0]
    casillas = [d["rect"] for d in pagina.get_drawings() if d["type"] == "f" and 5 <= d["rect"].width <= 14]
    for n, cb in FM["checkboxes"].items():
        x, y = cb["origen"]
        centro = fitz.Point(x + 3.0, y - 3.2)   # centro aproximado del glifo 'X' (6 pt de ancho)
        assert any(r.contains(centro) for r in casillas), f"{n}: la X no cae en ninguna casilla del machote"
