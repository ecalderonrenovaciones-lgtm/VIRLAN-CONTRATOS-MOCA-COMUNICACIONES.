from email.message import EmailMessage

import pytest

from virlan_bot.eml_reader import CorreoIncompletoError, extraer_adjuntos, extraer_personas_autorizadas


@pytest.mark.parametrize(
    "cuerpo,esperado",
    [
        ("PERSONAS AUTORIZADAS A RECIBIR: RICARDO ROCHA, CLAUDIA VILLEGAS\n\nDe: x", ["RICARDO ROCHA", "CLAUDIA VILLEGAS"]),
        ("PERSONAS : JOSE GUSTAVO FLORES, JOSE EMILIANO MONROY", ["JOSE GUSTAVO FLORES", "JOSE EMILIANO MONROY"]),
        ("recibe en el de entrega, favor de cargar por modelo, personas autorizadas para recibir: Carlos Román Gama\n\nNOTA: x",
         ["Carlos Román Gama"]),
        ("Pendiente\n\nReciben: Orlando González Ramírez y Yareli Messina Sanchez\n\nSaludos",
         ["Orlando González Ramírez", "Yareli Messina Sanchez"]),
        ("PERSONAS AUTORIZADAS: MARIA RETIS, ANIBAL AGUILAR", ["MARIA RETIS", "ANIBAL AGUILAR"]),
        ("Recibe en el de entrega por favor\n\nSaludos", []),   # sin nombres
        ("Hola, se envía\n\nDe: alguien\nPERSONAS: NO DEBE LEERSE", []),  # después del reenvío no cuenta
        ("", []),
    ],
)
def test_personas_autorizadas_del_cuerpo(cuerpo, esperado):
    assert extraer_personas_autorizadas(cuerpo) == esperado


def _correo(tmp_path, adjuntos, cuerpo="Hola", asunto="ASUNTO // RENOVACION"):
    msg = EmailMessage()
    msg["Subject"] = asunto
    msg.set_content(cuerpo)
    for nombre in adjuntos:
        msg.add_attachment(b"contenido", maintype="application", subtype="octet-stream", filename=nombre)
    ruta = tmp_path / "correo.eml"
    ruta.write_bytes(bytes(msg))
    return ruta


ADJ_COMPLETOS = [
    "01control de renovacion.xlsx",
    "02Copia de LAYOUT DE VINCULACION EMPRESARIAL CURP.xlsx",
    "03SAE 2.2.xlsx",
    "EMPRESA PRUEBA SA DE CV.docx",
    "INE_Sr. JUAN PEREZ.pdf",
    "Lista Nominal _ INE JUAN.pdf",
    "CURP_XXXX010101HDFXXX00.pdf",
]


def test_clasifica_los_adjuntos_reales(tmp_path):
    a = extraer_adjuntos(_correo(tmp_path, ADJ_COMPLETOS, "PERSONAS: A B, C D"), tmp_path / "out")
    assert a.ficha_docx.name == "EMPRESA PRUEBA SA DE CV.docx"
    assert a.sae_xlsx.name == "03SAE 2.2.xlsx"
    assert a.control_renovacion_xlsx.name == "01control de renovacion.xlsx"
    assert a.vinculacion_xlsx.name.startswith("02Copia")
    assert a.ine_pdf.name == "INE_Sr. JUAN PEREZ.pdf"          # NO la 'Lista Nominal _ INE ...'
    assert a.personas_autorizadas == ["A B", "C D"]
    assert a.asunto.endswith("RENOVACION")


@pytest.mark.parametrize("nombre_ine", ["ine_JUAN PEREZ.pdf", "INE JUAN PEREZ.pdf", "INE_JUAN.pdf"])
def test_ine_se_detecta_en_minusculas_y_con_espacio(tmp_path, nombre_ine):
    adjuntos = [a for a in ADJ_COMPLETOS if not a.startswith("INE_")] + [nombre_ine]
    assert extraer_adjuntos(_correo(tmp_path, adjuntos), tmp_path / "o").ine_pdf.name == nombre_ine


def test_oferta_comercial_por_nombre(tmp_path):
    a = extraer_adjuntos(_correo(tmp_path, ADJ_COMPLETOS + ["PdfPropuestaComercial X.pdf"]), tmp_path / "o")
    assert a.oferta_comercial_pdf.name == "PdfPropuestaComercial X.pdf"


def test_correo_sin_ficha_se_detiene(tmp_path):
    with pytest.raises(CorreoIncompletoError):
        extraer_adjuntos(_correo(tmp_path, [a for a in ADJ_COMPLETOS if not a.endswith(".docx")]), tmp_path / "o")


def test_formato_no_soportado(tmp_path):
    ruta = tmp_path / "correo.txt"
    ruta.write_text("x")
    with pytest.raises(ValueError):
        extraer_adjuntos(ruta, tmp_path / "o")
