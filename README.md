# tableau_datum

Pipeline para bajar los reportes de venta de **Datum** (POS de Paraiso Country
Club), unificarlos en un solo libro y acomodarlos en las hojas que consume el
dashboard de **Tableau**.

```
Datum (web) --> downloads/*.xlsx --> DATUM_unificado.xlsx --> DATUM.xlsx --> Tableau (.twbx)
              (1 archivo/día)      (consolidator.py)      (sheeter.py)
```

## Requisitos

- Python 3.10+ (probado con 3.14, ver `.venv`)
- Dependencias: `pip install -r requirements.txt`
  (`playwright`, `python-dotenv`, `pandas`, `openpyxl`, `requests`)
- Si vas a usar `--mode ui`, además:
  ```
  playwright install chrome
  ```

## Configuración

Copia `.env.example` a `.env` y llena las credenciales de Datum:

```
DATUM_USERNAME=
DATUM_PASSWORD=
DATUM_BASE=
```

`AREAS_BLACKLIST` (en `config.py`) excluye áreas cuyo nombre contenga alguna
de esas palabras (por defecto `"pruebas"`).

## Uso

```bash
python main.py --mode url --from 2025-01-01 --to 2026-08-26
```

Flags principales de `main.py`:

| Flag | Default | Qué hace |
|---|---|---|
| `--mode` | `url` | `url` = descarga por HTTP puro, sin navegador (sirve en servidor headless). `ui` = abre Chromium real y llena el formulario a clicks. |
| `--from` / `--to` | `2025-01-01` / ayer | Rango de fechas a descargar (un archivo por día). |
| `--exit` | `downloads/` | Carpeta donde caen los `.xlsx` diarios. |
| `--rewrite` | off | Vuelve a bajar días que ya estén en disco. |
| `--retries` | `3` | Reintentos por día antes de darlo por perdido. |
| `--headless` | off | Solo aplica con `--mode ui`. |
| `--slow-mo` | `0` | ms de pausa entre acciones en `--mode ui`, para ver el recorrido. |
| `--no-consolidate` | off | Solo descarga, no arma `DATUM_unificado.xlsx`. |
| `--consolidate-out` | `DATUM_unificado.xlsx` | Salida de `consolidator.py`. |
| `--no-layout` | off | No acomoda el consolidado en hojas (se queda en el paso anterior). |
| `--layout-out` | `DATUM.xlsx` | Libro final por hojas, el que lee Tableau. |

Si el rango queda incompleto (algún día falló), el comando lo dice al final;
se puede volver a correr tal cual: los días que ya están en disco se saltan.

### Correr los pasos sueltos

```bash
python consolidator.py --dir downloads --out DATUM_unificado.xlsx [--dedup]
python sheeter.py --src DATUM_unificado.xlsx --out DATUM.xlsx
```

## Qué descarga cada día

Al llenar el formulario de reporte de Datum se marcan/desmarcan estos
checkboxes (ver `select_all()` en `browser_downloader.py` para `--mode ui`):

- **Tipo, Tipo de producto, Día, Hora**: todos activados.
- **Status**: todos activados **excepto `frStatus2`**, que queda desmarcado.
- **Producto**: `frProductoVendido` activado, `frProductoCancelado`
  desmarcado.

> ⚠️ **Pendiente**: `http_downloader.py` (`--mode url`) todavía usa el patrón
> viejo (`FORCE_CHECK_PATTERN`) que marca *todos* los checkboxes de Status y
> Producto por igual, sin el ajuste de `frStatus2` / `frProductoCancelado`
> descrito arriba. Si vas a correr en modo `url` en producción, hay que
> alinear ese patrón con la lógica de `browser_downloader.py`.

## Consolidación (`consolidator.py`)

- Junta un `.xlsx` por día (`reporte_ventas_YYYY-MM-DD.xlsx`) leyendo
  `A5:W<fin>` de cada uno (el detalle termina donde `Area` se vacía; después
  viene el Total/resumen de Datum, que se descarta).
- Tipa columnas numéricas, porcentajes y fecha; por defecto **no** quita
  duplicados (Datum repite legítimamente la misma línea cuando se piden
  varios productos en una nota) — usa `--dedup` si hace falta.
- Soporta salida `.xlsx`, `.csv` o `.parquet`.

## Acomodo por hojas (`sheeter.py`)

Reproduce la estructura de `DATUM_bueno.xlsx` en 4 hojas:

| Hoja | Contenido |
|---|---|
| `Hoja1` | Alimentos y Bebidas, con `Tipo Conjunto` / `Tipo` / `Subtipo` derivados por reglas de texto (bebidas, alimentos, tabaco, descuentos, eventos, clases golf, modificadores; con/sin alcohol). |
| `Casa club` | Ventas con Área de negocio = Casa club. |
| `Campo de golf` | Ventas con Área de negocio = Campo de golf. |
| `Gastos` | Captura manual: **se hereda** del `DATUM.xlsx` anterior en cada corrida (nunca se pisa; captúrala directo en Excel). |

El Área de negocio se deriva de `Area`/`Grupo` con reglas fijas
(`AREA_OVERRIDE`, `Proshop` se reparte según si el grupo empieza con
`"CAMPO DE GOLF"`, etc. — ver `derive_area_negocio()`).

## Tableau

El dashboard vive en `out/Ventas Paraiso al <fecha>.twbx`, con 4 fuentes de
datos (`Casa club`, `Gastos`, `Campo de golf`, `Hoja1`/Alimentos y bebidas),
cada una como extracto `.hyper` sacado de una hoja de `DATUM.xlsx`.

- La conexión Excel de esas 4 fuentes debe apuntar al `DATUM.xlsx` **de este
  proyecto** (`/Users/eumircamargo/Desktop/tableau_datum/DATUM.xlsx`), no a
  una copia suelta en otra carpeta.
- Flujo normal: corre `main.py`, y luego en Tableau **Datos → Actualizar
  todos los extractos** y vuelve a guardar el `.twbx`.

## Estructura del repo

```
main.py                CLI: orquesta descarga -> consolidar -> acomodar
browser_downloader.py  Descarga por navegador real (Playwright, --mode ui)
http_downloader.py     Descarga por HTTP puro, sin navegador (--mode url)
consolidator.py        Une los .xlsx diarios en un solo libro
sheeter.py             Acomoda el consolidado en las 4 hojas finales
common.py              Tipos y helpers compartidos (rango de días, nombres de archivo)
config.py              Carga credenciales/URL de Datum desde .env
downloads/             Un .xlsx por día, tal como lo entrega Datum
out/                   Dashboard(s) de Tableau (.twbx)
```

## Troubleshooting

- **"El login no pasó"**: revisa `DATUM_USERNAME`/`DATUM_PASSWORD` en `.env`.
- **Archivo descargado pesa poco / no es `.xlsx`**: normalmente es la sesión
  expirada o Datum devolviendo una página de error en vez del reporte; el
  downloader lo detecta y reintenta con sesión fresca.
- **Tableau no refleja los datos nuevos**: confirma que la conexión de las 4
  fuentes de datos apunte al `DATUM.xlsx` de este proyecto (ver sección
  Tableau) y corre "Actualizar todos los extractos".
