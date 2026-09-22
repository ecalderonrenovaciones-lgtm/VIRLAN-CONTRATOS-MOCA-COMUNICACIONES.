"""Carga y valida config/contrato_fieldmap.v<N>.json."""

from __future__ import annotations

import json
from pathlib import Path

_CONFIG_DIR = Path(__file__).resolve().parent.parent / "config"


def cargar_fieldmap(version: str = "v1") -> dict:
    ruta = _CONFIG_DIR / f"contrato_fieldmap.{version}.json"
    if not ruta.exists():
        raise FileNotFoundError(f"No existe el mapa de campos {ruta}")
    with open(ruta, encoding="utf-8") as f:
        return json.load(f)
