# VIRLAN CONTRATOS BOT

Bot en Python que llena el **CONTRATO** (PDF AT&T) y la **OP / Orden de Programación** (Excel→PDF) de VIRLAN a partir del correo `.eml` que llega con los archivos de un cliente a renovar.

> ⚠️ Repositorio **privado**, de uso interno. Contiene documentos y datos reales de clientes (RFC, domicilios, identificaciones oficiales). No hacer público ni compartir fuera de la empresa.

## Requisitos

- **Windows** con **Microsoft Excel** instalado (la OP se llena y exporta a PDF con Excel real vía `pywin32`/COM, no con librerías puras, porque el machote usa *shapes* de Excel que `openpyxl` no puede editar).
- Python 3.12 (probado con esa versión).

## Instalación

```bash
pip install -r requirements.txt
```

## Estructura del proyecto

```
virlan_bot/                  Código del bot
config/contrato_fieldmap.v1.json   Coordenadas/calibración del PDF del CONTRATO
MACHOTE CONTRATOS/           Plantillas en blanco (PDF del contrato, Excel de la OP) — no tocar
DOCUMENTOS CONSULTA/         Lista de precios, cálculo MPE y catálogo de ladas vigentes
CORREO DE INFORMACIÓN/       Correos .eml de clientes ya procesados (ejemplo/histórico)
salida/                      Salida generada por el bot (se crea sola al correrlo)
tests/fixtures/              Ejemplos reales usados para verificar el pipeline contra archivos reales
```

`DOCUMENTOS CONSULTA/` se edita en vivo (el usuario reemplaza la lista de precios, el cálculo MPE y el catálogo de ladas cuando le llegan versiones nuevas por correo). El bot **siempre toma el archivo más reciente** que coincida con el patrón esperado (`Lista de Precios*.xlsx`, `*Cálculo de MPE*.xlsx`, `ladas_mexico*.csv`), no un nombre exacto, así que basta con dejar caer el archivo nuevo en esa carpeta.

## Uso

```bash
py -m virlan_bot.cli procesar --eml "ruta\al\correo.eml" [--tipo-venta RENOVACION|NUEVA|ADICION] [--persona-autorizada "Nombre"]
```

- `--tipo-venta` es opcional: si se omite, se detecta del asunto del correo (sin palabra clave reconocida, se asume RENOVACIÓN y se genera una alerta).
- `--persona-autorizada` es opcional: si se omite, se usa el representante legal del cliente por defecto.

### Salida

El bot crea `salida/<cuenta>_<razón_social>/` con:

- `contrato_borrador.pdf` — CONTRATO llenado
- `op_borrador.xlsx` / `op_borrador.pdf` — Orden de Programación llenada
- `revision.html` — paquete de revisión humana con todos los datos usados, su origen (de qué archivo salió cada dato) y las **alertas** del pipeline

**El pipeline nunca debe considerarse terminado sin que una persona revise `revision.html` antes de enviar el CONTRATO/OP.** Las alertas marcan explícitamente todo lo que no se pudo determinar con certeza (desbordes de texto, campos sin fuente confirmada, discrepancias entre archivos, etc.) — nunca se rellena nada a ciegas.

## Reglas de negocio confirmadas

Las reglas de negocio no evidentes en el código (de dónde sale cada dato, qué hacer ante desbordes, cuándo aplica cada addon, etc.) están documentadas como comentarios junto a la lógica correspondiente en `virlan_bot/`, principalmente en `contrato_valores.py` y `op_filler.py`. Ante cualquier duda sobre una regla, verificar contra un archivo real de ejemplo antes de asumir el formato — los archivos que mandan los clientes no siempre son consistentes entre sí.
