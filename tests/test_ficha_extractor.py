import pytest

from tests.conftest import DOM_ENTREGA, DOM_FISCAL, crear_docx
from virlan_bot.ficha_extractor import extraer_ficha, quitar_titulos
from virlan_bot.models import DatosIncompletosError


@pytest.mark.parametrize(
    "entrada,esperado",
    [
        ("Sr. FRANCISCO OLLIVIER ROMERO", "FRANCISCO OLLIVIER ROMERO"),
        ("Sra. MARTHA RUTH HILTON GONZALEZ", "MARTHA RUTH HILTON GONZALEZ"),
        ("Lic. Juan Pérez", "Juan Pérez"),
        ("Ing. Dr. Ana Ruiz", "Ana Ruiz"),
        ("C.P. Luis Mora", "Luis Mora"),
        ("SEÑOR José Cruz", "José Cruz"),
        ("Sr JOSE", "JOSE"),
        ("MONICA LORENA COMADURAN GARCIA", "MONICA LORENA COMADURAN GARCIA"),
        ("Ingrid Soto", "Ingrid Soto"),   # 'Ing' + 'rid': NO es un título
        ("Drake Lopez", "Drake Lopez"),   # 'Dr' + 'ake': NO es un título
        ("", ""),
    ],
)
def test_quitar_titulos(entrada, esperado):
    assert quitar_titulos(entrada) == esperado


def test_ficha_simple(ficha_simple):
    c = extraer_ficha(ficha_simple)
    assert c.razon_social == "EMPRESA PRUEBA SA DE CV"
    assert c.numero_cuenta == "500000001"
    assert c.representante_legal == "JUAN PEREZ LOPEZ"          # sin "Sr."
    assert (c.domicilio_calle, c.domicilio_numero, c.domicilio_colonia) == ("CALLE UNO", "12", "COL CENTRO")
    assert (c.domicilio_ciudad, c.domicilio_estado, c.domicilio_cp) == ("CUAUHTEMOC", "CIUDAD DE MEXICO", "06000")
    assert (c.rfc, c.telefono, c.correo) == ("EPR010101AB1", "5512345678", "juan@prueba.mx")
    assert c.envio_a == "FISCAL"
    assert any("no indica a qué domicilio" in a for a in c.alertas)  # sin banner: se asume FISCAL + alerta


def test_ficha_con_banner_y_rotulos_de_entrega(ficha_entrega):
    c = extraer_ficha(ficha_entrega)
    assert c.representante_legal == "MARIA GOMEZ RUIZ"
    assert c.envio_a == "ENTREGA"
    assert c.domicilio_calle == "CALLE UNO"                         # el fiscal no se pisa
    assert c.domicilio_entrega_calle == "AV DOS"
    assert c.domicilio_entrega_numero == "45 INT B"
    assert c.domicilio_entrega_estado == "JALISCO"
    assert c.entrega_calle_numero() == "AV DOS, 45 INT B"
    assert c.entrega_ciudad_estado_cp() == "GUADALAJARA, JALISCO, 44100"


def test_banner_fiscal_con_segundo_domicilio_usa_fiscal(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "500000002", "EL PAQUETE SE ENVIA A ESTE DOMICILIO:FISCAL", "JUAN PEREZ",
         "DOMICILIO FISCAL", DOM_FISCAL, "EPR010101AB1", "5512345678", "j@p.mx", "ENTREGA", DOM_ENTREGA],
    )
    c = extraer_ficha(ruta)
    assert c.envio_a == "FISCAL"
    assert c.domicilio_entrega_calle == "AV DOS"  # se guarda, pero no se usa


def test_rotulo_entrega_sin_dos_puntos_ni_acentos_y_variantes(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "500000002", "el paquete se envía a este domicilio: entrega", "JUAN PEREZ",
         "FISCAL", DOM_FISCAL, "EPR010101AB1", "5512345678", "j@p.mx", "DOMICILIO DE ENTREGA", DOM_ENTREGA],
    )
    assert extraer_ficha(ruta).envio_a == "ENTREGA"


def test_envio_a_entrega_sin_domicilio_de_entrega_se_detiene(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "500000002", "EL PAQUETE SE ENVIA A ESTE DOMICILIO:ENTREGA", "JUAN PEREZ",
         "DOMICILIO FISCAL", DOM_FISCAL, "EPR010101AB1", "5512345678", "j@p.mx"],
    )
    with pytest.raises(DatosIncompletosError):
        extraer_ficha(ruta)


def test_ficha_sin_rfc_no_se_detiene(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "500000002", "JUAN PEREZ", DOM_FISCAL, "5512345678", "j@p.mx", "extra"],
    )
    assert extraer_ficha(ruta).rfc == ""  # el RFC oficial sale del LAYOUT DE VINCULACION


def test_ficha_sin_telefono_queda_en_blanco_con_alerta(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "500000002", "JUAN PEREZ", DOM_FISCAL, "EPR010101AB1", "j@p.mx", "extra"],
    )
    c = extraer_ficha(ruta)
    assert c.telefono == "" and "telefono" in c.campos_faltantes


def test_correo_con_etiqueta_y_domicilio_duplicado(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "500000002", "JUAN PEREZ", DOM_FISCAL, "Correo Electrónico : j@p.mx", DOM_FISCAL,
         "EPR010101AB1", "5512345678"],
    )
    c = extraer_ficha(ruta)
    assert c.correo == "j@p.mx"
    assert c.domicilio_entrega_raw == ""  # el duplicado exacto no es un 2º domicilio


def test_domicilio_con_componentes_de_mas_no_se_inventa(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "500000002", "JUAN PEREZ", "A, B, C, D, E, F, G", "EPR010101AB1", "5512345678", "j@p.mx"],
    )
    c = extraer_ficha(ruta)
    assert c.domicilio_calle == "A, B, C, D, E, F, G" and any("componentes" in a for a in c.alertas)


def test_ficha_demasiado_corta_se_detiene(tmp_path):
    with pytest.raises(DatosIncompletosError):
        extraer_ficha(crear_docx(tmp_path / "f.docx", ["SOLO", "DOS"]))


def test_cuenta_no_numerica_se_detiene(tmp_path):
    ruta = crear_docx(
        tmp_path / "f.docx",
        ["EMPRESA PRUEBA", "ABC", "JUAN PEREZ", DOM_FISCAL, "EPR010101AB1", "5512345678", "j@p.mx"],
    )
    with pytest.raises(DatosIncompletosError):
        extraer_ficha(ruta)
