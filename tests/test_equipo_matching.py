"""Regla del usuario: SIEMPRE el mismo modelo; jamás el precio de otro equipo."""

from virlan_bot.equipo_matching import (
    buscar_coincidencias,
    mejor_sustituto,
    sugerencia_para_aviso,
)

CATALOGO = [
    ("SAMSUNG GALAXY Z FOLD8", 1),
    ("SAMSUNG GALAXY Z FOLD8 ULTRA", 2),
    ("APPLE IPHONE 17 PRO 256GB 5G", 3),
    ("APPLE IPHONE 17 PRO MAX 256GB 5G", 4),
    ("APPLE IPHONE 17 256GB 5G", 5),
    ("APPLE IPHONE 17 AIR 256GB 5G", 6),
    ("HONOR X5D", 7),
    ("SAMSUNG GALAXY A13 128GB 4G", 8),
    ("HONOR MAGIC8 LITE BUNDLE", 9),
    ("OPPO FIND X8 PRO 512GB 5G", 10),
]


def buscar(nombre):
    return buscar_coincidencias(nombre, CATALOGO)


def test_prefiere_el_nombre_mas_exacto_entre_variantes():
    assert buscar("APPLE IPHONE 17 256GB 5G") == [5]
    assert buscar("APPLE IPHONE 17 PRO 256GB 5G") == [3]
    assert buscar("APPLE IPHONE 17 PRO MAX 256GB 5G") == [4]


def test_tolera_que_falte_la_conectividad_en_el_nombre_de_SAE():
    # SAE no siempre trae '5G'; el catálogo sí: debe elegir PRO, no PRO MAX
    assert buscar("APPLE IPHONE 17 PRO 256GB") == [3]


def test_tolera_que_el_catalogo_omita_gb_y_conectividad():
    assert buscar("HONOR X5D 128GB 4G") == [7]


def test_mismo_modelo_con_distinto_espaciado():
    assert buscar("SAMSUNG GALAXY ZFOLD 8 ULTRA 512GB 5G") == [2]
    assert buscar("SAMSUNG GALAXY ZFOLD 8 512GB 5G") == [1]  # y NO el 'ULTRA'


def test_alias_confirmado_por_el_usuario():
    assert buscar("HONOR MAGIC8 LITE  BDLCHOICEH5G") == [9]


def test_nunca_usa_un_equipo_distinto():
    # A14 no existe: NO debe devolver el A13 ni el OPPO (misma capacidad/conectividad)
    assert buscar("SAMSUNG GALAXY A14 128GB 4G") == []
    assert buscar("SAMSUNG GALAXY Z FOLD 9 ULTRA 512GB 5G") == []


def test_capacidad_distinta_no_empata():
    assert buscar("HONOR X5C PLUS 128GB 4G") == []


def test_sugerencia_es_solo_texto_de_apoyo():
    texto = sugerencia_para_aviso("SAMSUNG GALAXY A14 128GB 4G", CATALOGO)
    assert "NO usados" in texto
    assert "A13" in texto  # se sugiere, pero solo como pista


def test_mejor_sustituto_exige_misma_capacidad_y_conectividad():
    assert mejor_sustituto("SAMSUNG GALAXY A14 128GB 4G", CATALOGO).familia_original == "SAMSUNG GALAXY A13 128GB 4G"
    assert mejor_sustituto("SAMSUNG GALAXY A14 999GB 4G", CATALOGO) is None
    assert mejor_sustituto("EQUIPO SIN CAPACIDAD", CATALOGO) is None
