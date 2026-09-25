"""Lista de precios, calculadora MPE y el criterio alternativo (precio de lista)."""

import datetime as dt
import math

import pytest

from virlan_bot.calculadora_mpe import calcular_mpe, mpe_desde_lista
from virlan_bot.catalogo_precios import EquipoNoEncontradoError, buscar_precio_lista
from virlan_bot.calculadora_mpe import EquipoNoEncontradoError as MPENoEncontrado

HOY = dt.date(2026, 9, 24)


def precio(lista, equipo, plan=599, plazo=24):
    return buscar_precio_lista(equipo, lista, fecha=HOY, numero_plan=plan, plazo_meses=plazo)


# ------------------------------------------------------------- lista de precios
def test_precio_y_diferencial_del_mismo_equipo(lista_precios):
    p = precio(lista_precios, "APPLE IPHONE 17 256GB 5G")
    assert (p.precio_lista, p.diferencial_equipo) == (19999, 13199)


def test_no_confunde_variantes_pro_air_promax(lista_precios):
    assert precio(lista_precios, "APPLE IPHONE 17 PRO 256GB").precio_lista == 28499  # SAE sin '5G'
    assert precio(lista_precios, "APPLE IPHONE 17 PRO MAX 256GB 5G").precio_lista == 30999


def test_guion_en_el_diferencial_es_sin_dato_no_error(lista_precios):
    p = precio(lista_precios, "HONOR X5D 128GB 4G")
    assert p.precio_lista == 3999 and p.diferencial_equipo is None


def test_solo_cuenta_la_vigencia_de_hoy(lista_precios):
    assert precio(lista_precios, "OPPO A80 256GB 5G").diferencial_equipo == 1199


def test_plan_o_plazo_fuera_de_catalogo_no_da_diferencial(lista_precios):
    assert precio(lista_precios, "APPLE IPHONE 17 256GB 5G", plan=123).diferencial_equipo is None
    assert precio(lista_precios, "APPLE IPHONE 17 256GB 5G", plazo=7).diferencial_equipo is None


def test_alias_en_lista(lista_precios):
    assert precio(lista_precios, "HONOR MAGIC8 LITE  BDLCHOICEH5G").precio_lista == 9999


def test_equipo_inexistente_explica_y_no_sustituye(lista_precios):
    with pytest.raises(EquipoNoEncontradoError) as e:
        precio(lista_precios, "SAMSUNG GALAXY A14 128GB 4G")
    assert "MISMO modelo" in str(e.value) and "NO usados" in str(e.value)


# ------------------------------------------------------------- calculadora MPE
def test_mpe_formula_de_la_calculadora(calculo_mpe):
    r = calcular_mpe("APPLE IPHONE 17 256GB 5G", 24, calculo_mpe, pie=13199)
    assert r.precio_base == 19999
    assert r.mpe == math.trunc(((19999 - 13199) / 24 + 0.01) * 100) / 100 == 283.34


def test_mpe_encuentra_nombres_distintos_del_mismo_equipo(calculo_mpe):
    assert calcular_mpe("HONOR X5D 128GB 4G", 24, calculo_mpe).precio_base == 3999
    assert calcular_mpe("SAMSUNG GALAXY ZFOLD 8 ULTRA 512GB 5G", 24, calculo_mpe).precio_base == 48499
    assert calcular_mpe("HONOR MAGIC8 LITE  BDLCHOICEH5G", 24, calculo_mpe).precio_base == 9999


def test_mpe_nunca_usa_otro_equipo(calculo_mpe):
    with pytest.raises(MPENoEncontrado) as e:
        calcular_mpe("OPPO A80 256GB 5G", 24, calculo_mpe)
    assert "MISMO modelo" in str(e.value)


@pytest.mark.parametrize(
    "lista,pie,plazo,esperado",
    [(19999, 13199, 24, 283.34), (37999, 20349, 24, 735.42), (53499, 37759, 24, 655.84), (3999, 1259, 24, 114.17)],
)
def test_criterio_alternativo_coincide_con_la_calculadora(lista, pie, plazo, esperado):
    """Casos reales verificados (37 de 38 líneas ya generadas coincidían exacto)."""
    assert mpe_desde_lista(lista, pie, plazo) == esperado


def test_criterio_alternativo_pie_igual_a_precio_da_un_centavo():
    assert mpe_desde_lista(19999, 19999, 24) == 0.01
