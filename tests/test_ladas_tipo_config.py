import pytest

from virlan_bot import config
from virlan_bot.ladas_lookup import LadaNoEncontradaError, ciudad_dn_por_telefono
from virlan_bot.tipo_contratacion import detectar_tipo_desde_asunto


# ------------------------------------------------------------------ Ciudad DN
@pytest.mark.parametrize(
    "telefono,esperado",
    [
        ("5512345678", "CDMX"),     # lada 55: 'Ciudad de Mexico' -> CDMX (aunque también aparezca como Edo. de México)
        ("3312345678", "Guadalajara"),
        ("8112345678", "Monterrey"),
        ("2215699113", "Puebla"),   # San Martin Texmelucan -> el estado
        ("2231234567", "Puebla"),   # lada 223 aparece con Puebla y Tlaxcala: el usuario indicó dejarla como Puebla
        ("7221234567", "EDOMEX"),
        ("9991234567", "Merida"),
        ("4441234567", "San"),      # regla de la primera palabra (se conserva tal cual)
    ],
)
def test_ciudad_dn(ladas_csv, telefono, esperado):
    assert ciudad_dn_por_telefono(telefono, ladas_csv) == esperado


def test_telefono_invalido_o_sin_lada_no_adivina(ladas_csv):
    for malo in ("NA", "123", ""):
        with pytest.raises(LadaNoEncontradaError):
            ciudad_dn_por_telefono(malo, ladas_csv)
    with pytest.raises(LadaNoEncontradaError):
        ciudad_dn_por_telefono("0000000000", ladas_csv)


def test_ladas_reales_si_existen():
    real = config.ladas_csv()
    if not real:
        pytest.skip("no hay catálogo de ladas en DOCUMENTOS CONSULTA")
    assert ciudad_dn_por_telefono("5512345678", real) == "CDMX"
    assert ciudad_dn_por_telefono("5612345678", real) == "CDMX"


# ------------------------------------------------------------------ tipo de contratación
@pytest.mark.parametrize(
    "asunto,tipo",
    [
        ("X // ADICION", "ADICION"), ("X // ADICIÓN", "ADICION"), ("X // ADICIONES", "ADICION"),
        ("X // RENOVACION T2", "RENOVACION"), ("X // RENOVACIONES", "RENOVACION"), ("X // RENOVACIÓN VENCIDA", "RENOVACION"),
        ("X // VENTA NUEVA", "NUEVA"), ("X // NUEVA", "NUEVA"),
    ],
)
def test_tipo_desde_asunto(asunto, tipo):
    t, alertas = detectar_tipo_desde_asunto(asunto)
    assert t == tipo and alertas == []


@pytest.mark.parametrize("asunto", ["X // VENCIDA", "X // VENCIDAS", "sin palabra clave"])
def test_sin_palabra_clave_asume_renovacion_con_alerta(asunto):
    t, alertas = detectar_tipo_desde_asunto(asunto)
    assert t == "RENOVACION" and alertas


# ------------------------------------------------------------------ canal / vendedor
@pytest.mark.parametrize(
    "asunto,canal",
    [("CENTRO // DEALERS // ONE STOP // 6108", "ONE STOP"), ("X // one stock", "ONE STOP"),
     ("X // FES PRIMECOMMS", None), ("X // PRIME", None), ("", None)],
)
def test_detectar_canal(asunto, canal):
    assert config.detectar_canal(asunto) == canal


def test_vendedor_one_stop_tiene_todos_los_datos_del_ejemplo():
    v = config.VENDEDORES_POR_CANAL["ONE STOP"]
    assert v["nombre_ejecutivo"] == "ALEJANDRO CASTERA ARELLANOS"
    assert v["rfc_ejecutivo"] == "CAAA760226LJ4"
    assert v["punto_venta_nombre"] == "ONE STOP MARKET AC425E"
    assert v["punto_venta_codigo"] == ""   # vacío A PROPÓSITO (el código va dentro del nombre)


def test_limites_de_configuracion():
    assert config.LIMITE_PDF_KB == 700
    assert config.ENTREGA_POR_DEFECTO["dias_habiles_entrega"] == 14
    assert config.CONTRATO_FIELDMAP_VERSION == "v2"


def test_cli_ofrece_las_opciones_documentadas(capsys):
    from virlan_bot import cli

    with pytest.raises(SystemExit) as salida:
        cli.main(["procesar", "--help"])
    assert salida.value.code == 0
    ayuda = capsys.readouterr().out
    for opcion in ("--eml", "--tipo-venta", "--persona-autorizada", "--fecha-contratacion", "--sufijo",
                   "--solo-contrato", "--ejecutivo", "--rfc-ejecutivo", "--punto-venta-nombre", "--punto-venta-codigo"):
        assert opcion in ayuda, opcion
