"""Rutas y valores por defecto del proyecto."""

from __future__ import annotations

from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent

MACHOTE_CONTRATO_PDF = RAIZ / "MACHOTE CONTRATOS" / "Contrato nueva versión NUEVO.pdf"
MACHOTE_OP_XLSX = RAIZ / "MACHOTE CONTRATOS" / "08 Orden Program EBS.xlsx"
DOCUMENTOS_CONSULTA_DIR = RAIZ / "DOCUMENTOS CONSULTA"
SALIDA_DIR = RAIZ / "salida"


def _mas_reciente(patron: str) -> Path | None:
    """Localiza en DOCUMENTOS_CONSULTA_DIR el archivo más reciente que
    coincida con `patron` (glob). La lista de precios y la calculadora MPE
    se actualizan periódicamente por correo, así que siempre se usa la
    última copia que el usuario haya guardado ahí."""
    coincidencias = sorted(
        DOCUMENTOS_CONSULTA_DIR.glob(patron), key=lambda p: p.stat().st_mtime, reverse=True
    )
    return coincidencias[0] if coincidencias else None


def lista_precios_xlsx() -> Path | None:
    return _mas_reciente("Lista de Precios*.xlsx")


def calculo_mpe_xlsx() -> Path | None:
    return _mas_reciente("*Cálculo de MPE*.xlsx") or _mas_reciente("*Calculo de MPE*.xlsx")


def ladas_csv() -> Path | None:
    # El nombre del archivo ha variado entre correos ('ladas_mexico (2).csv',
    # 'ladas_mexico Final.csv'); se localiza por prefijo en vez de nombre exacto.
    return _mas_reciente("ladas_mexico*.csv") or _mas_reciente("*ladas*mexico*.csv")

CONTRATO_FIELDMAP_VERSION = "v1"

# Datos del punto de venta/ejecutivo: en los dos ejemplos reales disponibles
# fueron siempre los mismos (mismo distribuidor/ejecutivo), y no vienen en
# ningún archivo del correo — se usan como valores por defecto configurables
# hasta que se identifique una fuente por cliente.
VENDEDOR_POR_DEFECTO = {
    "plaza_de_venta": "MEXICO",
    "ciudad": "MEXICO",
    "punto_venta_nombre": "PRIME COMMS TDA CDMX CIBELES",
    "punto_venta_codigo": "9109783-5                        SV348C",
    "nombre_ejecutivo": "SERGIO YAHIR VALDERRAMA VELOZ",
    "rfc_ejecutivo": "VAVS020813",
}

# Igual que arriba: en los dos ejemplos reales la fecha/hora máxima de
# entrega fue idéntica (ventana estándar de entrega). Se deja como
# configurable por si cambia.
ENTREGA_POR_DEFECTO = {
    "dias_habiles_entrega": 14,
    "hora_entrega": "9            18",
}
