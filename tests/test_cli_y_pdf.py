import datetime as dt
import os
from email.message import EmailMessage
from pathlib import Path

import fitz
import pytest

from virlan_bot import cli, config
from virlan_bot.eml_reader import extraer_adjuntos
from virlan_bot.pdf_optimizer import optimizar_pdf


# ------------------------------------------------------------------ utilidades del CLI
@pytest.mark.parametrize("texto", ["24-09-2026", "24/09/2026", "2026-09-24"])
def test_parsear_fecha(texto):
    assert cli._parsear_fecha(texto) == dt.date(2026, 9, 24)


@pytest.mark.parametrize("texto", ["24 sep", "32-13-2026", "", "mañana"])
def test_parsear_fecha_invalida(texto):
    with pytest.raises(ValueError):
        cli._parsear_fecha(texto)


def test_slug():
    assert cli._slug("SERVICIOS PROFESIONALES HERNANDEZ NUÑO Y ASOCIADOS S.C.") == "SERVICIOS_PROFESIONALES_HERNANDEZ_NUÑO_Y_ASOCIADOS_SC"
    assert cli._slug("A / B - C") == "A_B_C"


def _crear(carpeta: Path, *nombres):
    carpeta.mkdir(parents=True, exist_ok=True)
    for n in nombres:
        (carpeta / n).write_bytes(b"x")


def test_limpieza_de_op_obsoletas_conserva_las_vigentes(tmp_path):
    _crear(tmp_path, "op_borrador.pdf", "op_borrador.xlsx", "OP EQUIPOS.pdf", "OP SIM.pdf", "contrato_borrador.pdf",
           "op_p1.png", "op_equipos_p1.png", "op_sim_p1.png")
    avisos = cli._limpiar_op_obsoletas(tmp_path, ["OP EQUIPOS", "OP SIM"])
    assert avisos == []
    assert sorted(p.name for p in tmp_path.iterdir()) == sorted(
        ["OP EQUIPOS.pdf", "OP SIM.pdf", "contrato_borrador.pdf", "op_equipos_p1.png", "op_sim_p1.png"]
    )


def test_limpieza_cuando_ya_no_hay_mezcla_borra_las_dos_separadas(tmp_path):
    _crear(tmp_path, "OP EQUIPOS.pdf", "OP EQUIPOS.xlsx", "OP SIM.pdf", "op_borrador.pdf")
    cli._limpiar_op_obsoletas(tmp_path, ["op_borrador"])
    assert sorted(p.name for p in tmp_path.iterdir()) == ["op_borrador.pdf"]


def test_copia_a_contratos_terminados(tmp_path, monkeypatch):
    monkeypatch.setattr(config, "CONTRATOS_TERMINADOS_DIR", tmp_path / "TERM")
    salida = tmp_path / "salida"
    _crear(salida, "contrato_borrador.pdf", "op_borrador.pdf", "op_borrador.xlsx", "revision.html", "contrato_p1.png")
    assert cli._copiar_a_contratos_terminados(salida, "500_X", ["op_borrador"]) == []
    assert sorted(p.name for p in (tmp_path / "TERM" / "500_X").iterdir()) == [
        "contrato_borrador.pdf", "op_borrador.pdf", "op_borrador.xlsx"]   # sin revision.html ni PNG


def test_archivo_abierto_en_el_visor_no_tumba_la_corrida(tmp_path, monkeypatch):
    """Regresión: PermissionError WinError 32 (Acrobat abierto) cortaba el proceso al final."""
    monkeypatch.setattr(config, "CONTRATOS_TERMINADOS_DIR", tmp_path / "TERM")
    salida = tmp_path / "salida"
    _crear(salida, "contrato_borrador.pdf", "op_borrador.pdf", "op_borrador.xlsx")

    original = cli.shutil.copy2

    def copia_con_bloqueo(origen, destino, *a, **k):
        if str(origen).endswith("contrato_borrador.pdf"):
            raise PermissionError(32, "El proceso no tiene acceso al archivo porque está siendo utilizado por otro proceso")
        return original(origen, destino, *a, **k)

    monkeypatch.setattr(cli.shutil, "copy2", copia_con_bloqueo)
    avisos = cli._copiar_a_contratos_terminados(salida, "500_X", ["op_borrador"])
    assert len(avisos) == 1 and "contrato_borrador.pdf" in avisos[0] and "abierto" in avisos[0].lower()
    assert (tmp_path / "TERM" / "500_X" / "op_borrador.xlsx").exists()     # lo demás sí se copió


def test_main_muestra_mensaje_claro_si_un_archivo_esta_bloqueado(monkeypatch, capsys, tmp_path):
    def falla(*a, **k):
        raise PermissionError(13, "Permission denied", "C:/x/contrato_borrador.pdf")

    monkeypatch.setattr(cli, "procesar", falla)
    correo = tmp_path / "c.eml"
    correo.write_text("x")
    codigo = cli.main(["procesar", "--eml", str(correo)])
    err = capsys.readouterr().err
    assert codigo == 1 and "contrato_borrador.pdf" in err and "cierra" in err.lower()


# ------------------------------------------------------------------ rutas largas de Windows
def test_adjuntos_con_nombre_larguisimo_no_rompen_por_max_path(tmp_path):
    nombre = "CURP_" + "X" * 210 + ".pdf"
    msg = EmailMessage()
    msg["Subject"] = "S"
    msg.set_content("hola")
    for n in ("01control de renovacion.xlsx", "03SAE 2.2.xlsx", "FICHA.docx", nombre):
        msg.add_attachment(b"c", maintype="application", subtype="octet-stream", filename=n)
    eml = tmp_path / "c.eml"
    eml.write_bytes(bytes(msg))
    destino = tmp_path / ("carpeta_" + "d" * 30) / ("sub_" + "s" * 30)
    a = extraer_adjuntos(eml, destino)
    guardados = list(destino.iterdir())
    assert len(guardados) == 4
    assert max(len(str(p)) for p in guardados) < 250       # límite clásico de Windows: 260


# ------------------------------------------------------------------ optimizador de PDF
def _pdf_pesado(ruta: Path) -> Path:
    doc = fitz.open()
    for i in range(3):
        p = doc.new_page()
        p.insert_text((72, 72), f"Página {i} " + "texto repetido " * 200)
    doc.save(str(ruta), garbage=0, deflate=False)   # sin comprimir a propósito
    doc.close()
    return ruta


def test_optimizador_reduce_sin_perder_contenido(tmp_path):
    pdf = _pdf_pesado(tmp_path / "a.pdf")
    antes = fitz.open(pdf)[1].get_text()
    r = optimizar_pdf(pdf, 700)
    assert r.kb_despues <= r.kb_antes and r.alerta is None
    d = fitz.open(pdf)
    assert len(d) == 3 and d[1].get_text() == antes
    assert not list(tmp_path.glob("*._opt.pdf"))          # sin residuos


def test_optimizador_avisa_si_no_llega_al_limite(tmp_path):
    r = optimizar_pdf(_pdf_pesado(tmp_path / "b.pdf"), 0)
    assert r.alerta and "no se logró bajar" in r.alerta


def test_optimizador_no_agranda_un_pdf_ya_optimo(tmp_path):
    pdf = _pdf_pesado(tmp_path / "c.pdf")
    optimizar_pdf(pdf, 700)
    tam = os.path.getsize(pdf)
    optimizar_pdf(pdf, 700)
    assert os.path.getsize(pdf) <= tam
