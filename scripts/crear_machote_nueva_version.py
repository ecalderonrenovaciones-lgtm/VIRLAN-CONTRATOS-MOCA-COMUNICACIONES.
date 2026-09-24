"""Crea el MACHOTE EN BLANCO de la nueva versión del CONTRATO (2026-09-24) a
partir del ejemplo terminado que compartió el usuario
('EJEMPLO DE CONTRATOS/Contrato de servicios nueva versión TERMINADO.pdf').

Borra SOLO los datos de muestra de la página 1 (redacción de texto con
graphics=NONE e images=NONE, para no tocar líneas, casillas ni las marcas del
formato). Se conservan los valores constantes: 'VER ANEXO', 'NA', 'NA', la hora
de entrega y la X de Plazo Mínimo. Uso:  py scripts/crear_machote_nueva_version.py
"""

from pathlib import Path

import fitz

RAIZ = Path(__file__).resolve().parent.parent
EJEMPLO = RAIZ.parent / "ARCHIVOS INTERCAMBIABLES" / "EJEMPLO DE CONTRATOS" / "Contrato de servicios nueva versión TERMINADO.pdf"
SALIDA = RAIZ.parent / "ARCHIVOS INTERCAMBIABLES" / "MACHOTE CONTRATOS" / "Contrato de servicios nueva versión (machote en blanco).pdf"

# Se conservan (constantes del formato): texto -> se identifica por contenido.
CONSERVAR = {"VER ANEXO", "NA", "09:00        18:00"}
# La X de Plazo Mínimo (origen x≈119.1) también se conserva.
X_PLAZO_MINIMO_X = 119.1


def main() -> None:
    doc = fitz.open(str(EJEMPLO))
    page = doc[0]
    borradas = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            for s in l["spans"]:
                texto = s["text"].strip()
                if "Arial" not in s["font"] or not texto or texto in CONSERVAR:
                    continue
                ox, oy = s["origin"]
                if texto == "X" and abs(ox - X_PLAZO_MINIMO_X) < 0.5:
                    continue
                x0, _, x1, _ = s["bbox"]
                tam = s["size"]
                if texto == "X":
                    rect = fitz.Rect(ox - 0.4, oy - 0.8 * tam, ox + 6.4, oy + 0.05 * tam)
                else:
                    rect = fitz.Rect(x0 - 0.3, oy - 0.80 * tam, x1 + 0.3, oy + 0.26 * tam)
                page.add_redact_annot(rect)
                borradas.append((texto[:30], [round(v, 1) for v in rect]))
    page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_NONE, graphics=fitz.PDF_REDACT_LINE_ART_NONE)
    SALIDA.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(SALIDA), garbage=4, deflate=True)
    print(f"{len(borradas)} datos de muestra borrados -> {SALIDA.name}")
    for t, r in borradas:
        print("  ", r, repr(t))


if __name__ == "__main__":
    main()
