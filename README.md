# VIRLAN CONTRATOS BOT

Bot en Python que llena el **CONTRATO** (PDF AT&T, nueva versión) y la **OP / Orden de Programación** (Excel→PDF) de VIRLAN a partir del correo (`.eml` o `.msg`) que llega con los archivos de un cliente (renovación, adición o venta nueva).

> ⚠️ Repositorio **privado**, de uso interno. El trabajo con clientes maneja datos personales reales (RFC, domicilios, INE, CURP). **No compartas la carpeta completa**: comparte solo el código (sin `salida/`) o los PDF/Excel entregables. Ver "Privacidad".

## Requisitos

- **Windows** con **Microsoft Excel** instalado (la OP se llena y exporta a PDF con Excel real vía `pywin32`/COM, porque el machote usa *shapes* de Excel que `openpyxl` no puede editar).
- Fuentes de Windows en `C:\Windows\Fonts`: **Arial** (`arial.ttf`, `arialbd.ttf`; se incrusta en el CONTRATO para que sea editable) y **Calibri** (`calibri.ttf`; se usa para medir el texto de la OP).
- Python 3.12.

```bash
pip install -r requirements.txt          # ejecución
pip install -r requirements-dev.txt      # pruebas (pytest)
```

## Estructura

```
VIRLAN CONTRATOS BOT/
├── CONTRATOS TERMINADOS/            Entregables por cliente (copia curada desde salida/):
│   └── <cuenta>_<razón_social>[_SUFIJO]/   contrato_borrador.pdf + op_borrador.pdf/.xlsx
│                                           (o "OP EQUIPOS" y "OP SIM", ver abajo)
├── ARCHIVOS INTERCAMBIABLES/        Carpetas que el usuario edita en vivo (FUERA de Git):
│   ├── MACHOTE CONTRATOS/           Machotes en blanco: contrato (nueva versión) y OP
│   ├── DOCUMENTOS CONSULTA/         Lista de precios, cálculo MPE, catálogo de ladas
│   ├── EJEMPLO DE CONTRATOS/        Ejemplos ya llenos, de referencia (incl. ONE STOP)
│   └── CORREO DE INFORMACIÓN/       Correos de clientes
└── ARCHIVOS DEL BOT/                Todo lo demás (este README; aquí vive .git)
    ├── virlan_bot/                  Código
    ├── config/contrato_fieldmap.v2.json   Calibración del CONTRATO nueva versión (v1 = formato anterior, ya no se usa)
    ├── scripts/crear_machote_nueva_version.py   Genera el machote en blanco desde el ejemplo
    ├── tests/                       Pruebas automáticas (datos sintéticos)
    └── salida/                      Carpeta de trabajo (revision.html, PNG, INE, _tmp_adjuntos)
```

`DOCUMENTOS CONSULTA/` se edita en vivo: el bot toma **siempre el archivo más reciente** que coincida con `Lista de Precios*.xlsx`, `*Cálculo de MPE*.xlsx` y `ladas_mexico*.csv`.

## Uso

Desde `ARCHIVOS DEL BOT/`:

```bash
py -m virlan_bot.cli procesar --eml "ruta\correo.eml" --fecha-contratacion 24-09-2026 [opciones]
```

| Opción | Para qué |
|---|---|
| `--fecha-contratacion DD-MM-AAAA` | Fecha de **cotejo manuscrita en el `INE_*.pdf`**: es la fecha de contratación del CONTRATO y la OP (y la entrega máxima = +14 días). Es letra a mano: se lee a ojo (el bot no hace OCR). Sin ella usa hoy y avisa. |
| `--persona-autorizada "Nombre"` (repetible) | Quien recibe los equipos. Normalmente sale del cuerpo del correo (`PERSONAS…:`, `Reciben:`); sin nombres se usa el representante legal. |
| `--tipo-venta RENOVACION\|ADICION\|NUEVA` | Normalmente se detecta del asunto (acepta plurales). |
| `--sufijo ADICIONES` | Carpeta con sufijo, para dos paquetes de una misma cuenta (renovación y adición). |
| `--solo-contrato` | Regenera solo el CONTRATO; no toca las OP existentes. |
| `--ejecutivo "Nombre"` (+ `--rfc-ejecutivo`, `--punto-venta-nombre`, `--punto-venta-codigo`) | Vendedor distinto al predeterminado. El canal **ONE STOP** se detecta solo desde el asunto. |

`salida/<cuenta>_<razón>/` trae `contrato_borrador.pdf`, la(s) OP, y **`revision.html`** (datos, su origen, alertas y el INE al lado de la fecha usada). **Siempre revisar `revision.html` antes de enviar.** Cada corrida copia los entregables a `CONTRATOS TERMINADOS/`.

Si un archivo de `CONTRATOS TERMINADOS/` está abierto en Acrobat/Excel, la copia no corta la corrida: deja un aviso en `revision.html`; cierra el archivo y vuelve a copiarlo desde `salida/`.

## Reglas de negocio vigentes (resumen)

- **SIM y equipos nunca en la misma OP.** Si el cliente trae ambos: `OP EQUIPOS` y `OP SIM` (mismo contrato). En SIM: precio de lista, pago inicial, MPE y Ciudad DN sin lada → `N/A`.
- **Siempre el mismo modelo**: nunca se usa el precio de un equipo distinto (nada de "sustitutos"). Los nombres se reconocen por similitud (`equipo_matching.py`, alias confirmados en `ALIAS_EQUIPOS`); si no hay el mismo modelo → `N/A` + aviso con sugerencias (no usadas).
- **MPE** = calculadora; criterio alternativo `TRUNC((precio de lista − diferencial)/plazo + 0.01, 2)`. Si difieren gana el **precio de la lista de precios**.
- **Domicilio**: la ficha indica a dónde va el paquete (`EL PAQUETE SE ENVIA A ESTE DOMICILIO:FISCAL|ENTREGA`). ENTREGA → se llena también la fila de Domicilio Entrega del contrato y la OP lleva ese domicilio.
- **Nombres**: representante legal y personas autorizadas sin títulos (Sr., Lic., Ing.…). Estado `CDMX`/`EDOMEX` abreviado; Ciudad DN: `Ciudad de Mexico`→CDMX, San Martín Texmelucan→Puebla, primera palabra en los demás; la casilla que se desborde reduce solo su letra.
- **Vendedor por canal** (`config.VENDEDORES_POR_CANAL`): PRIME/FES PRIMECOMMS = predeterminado; ONE STOP con sus datos. Nunca se reutiliza el RFC/punto de venta de otro ejecutivo.
- **Archivos editables**: OP sin "reducir hasta ajustar" (tamaño de letra fijo), PDF con Arial incrustada. Cada PDF ≤ 700 KB (`pdf_optimizer.py`).
- Las reglas finas están comentadas junto al código (`op_filler.py`, `contrato_valores.py`, `equipo_matching.py`, `ficha_extractor.py`).

## Pruebas

```bash
py -m pytest                 # todo (~2 min; incluye Excel real)
py -m pytest -m "not excel"  # sin Excel (segundos)
```

Datos 100 % sintéticos (`tests/conftest.py`). Las marcadas `excel`/`machote` se **omiten solas** si no hay Excel o faltan los machotes de `ARCHIVOS INTERCAMBIABLES`. Cubren: reconocimiento de equipos, precios/MPE, ficha y correo, ladas, valores y llenado real del CONTRATO (rótulos intactos, X en su casilla, RFC de 12/13, fuentes incrustadas), coherencia del mapa de campos, optimizador de PDF, utilidades del CLI (archivo abierto, rutas largas) y punta a punta con la OP en Excel (dos OP, `--solo-contrato`, ONE STOP, precios).

## Mantenimiento

- **Otro formato de contrato**: poner el ejemplo en `EJEMPLO DE CONTRATOS/`, ajustar `scripts/crear_machote_nueva_version.py` y calibrar un `contrato_fieldmap.vN.json` (coordenadas por *origen* de cada dato del ejemplo); `pytest tests/test_fieldmap_v2.py` valida la geometría.
- **Otro canal de venta**: agregar su entrada en `VENDEDORES_POR_CANAL` y su detección en `detectar_canal` (`config.py`).
- **Equipo con nombre distinto entre SAE y catálogos**: agregar el alias en `ALIAS_EQUIPOS` **solo tras confirmación**.

## Límites conocidos

- La **fecha de cotejo** y el **nombre del ejecutivo que coteja** están escritos a mano en el INE: se leen a ojo, no de forma automática.
- El correo se indica manualmente (no hay lectura directa de Outlook).
- La edición del PDF en Acrobat no está probada (solo se verifica que la fuente esté incrustada).
- Windows + Excel obligatorios.

## Privacidad y Git

`.git` vive en `ARCHIVOS DEL BOT/`, así que `CONTRATOS TERMINADOS/` y `ARCHIVOS INTERCAMBIABLES/` no tienen respaldo en Git. `salida/` contiene INE, CURP y datos personales: las carpetas nuevas de cliente están en `.gitignore` (algunas antiguas ya estaban versionadas y siguen así).
