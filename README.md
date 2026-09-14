# tableau_datum

Pipeline para bajar los reportes de venta de **Datum** (POS de Paraiso Country
Club), unificarlos en un solo libro y acomodarlos en las hojas que consume el
dashboard de **Tableau**.

```
Datum (web) --> downloads/*.xlsx --> DATUM_unificado.xlsx --> DATUM.xlsx --> Tableau (.twbx)
              (1 archivo/día)      (consolidator.py)      (sheeter.py)
```

Tres pasos, cada uno con su script y su rol:

1. **Descarga** (`browser_downloader.py` / `http_downloader.py`, orquestados
   por `main.py`) — baja un `.xlsx` de Datum por cada día del rango pedido.
2. **Consolidación** (`consolidator.py`) — junta todos esos `.xlsx` diarios en
   un solo libro (`DATUM_unificado.xlsx`), tipando columnas numéricas/fecha.
3. **Acomodo por hojas** (`sheeter.py`) — reparte ese consolidado en las 4
   hojas que Tableau espera (`DATUM.xlsx`), calculando columnas derivadas
   (`Tipo`, `Subtipo`, `Punto de Venta`, `Area de negocio`).

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

## Cómo correrlo: los 2 modos de descarga

Todo se dispara con `main.py --mode <url|ui>`. Los dos modos bajan exactamente
los mismos reportes (un `.xlsx` por día) y después corren los mismos pasos de
consolidación y acomodo — la única diferencia es **cómo** le hablan a Datum.

### Modo `url` (por defecto, recomendado)

Descarga por HTTP puro, sin abrir navegador. Es el más rápido y el único que
sirve en un servidor sin entorno gráfico (headless / CI / cron).

```bash
python main.py --mode url --from 2025-01-01 --to 2026-08-26
```

> ⚠️ **Pendiente conocido**: `http_downloader.py` todavía usa un patrón de
> checkboxes viejo (`FORCE_CHECK_PATTERN`) que marca *todos* los de Status y
> Producto por igual, sin el ajuste de `frStatus2` / `frProductoCancelado`
> que sí tiene `--mode ui` (ver más abajo). Si vas a correr `url` en
> producción, alinea ese patrón con `browser_downloader.py` primero.

### Modo `ui`

Abre un Chromium real (Playwright) y llena el formulario de reporte a clicks,
igual que lo haría una persona. Más lento, pero es la referencia "correcta"
de qué checkboxes van marcados (ver sección siguiente), y sirve para ver con
tus propios ojos qué está pasando si algo falla en `url`.

```bash
python main.py --mode ui --from 2026-08-01 --to 2026-08-26 --headless
```

Quita `--headless` para ver el navegador en pantalla, o usa `--slow-mo 200`
para que vaya despacio y puedas seguir el recorrido paso a paso.

### Flags principales de `main.py`

| Flag | Default | Qué hace |
|---|---|---|
| `--mode` | `url` | `url` = descarga por HTTP puro, sin navegador. `ui` = abre Chromium real y llena el formulario a clicks. |
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

Útil si ya tienes los `.xlsx` diarios y solo quieres re-consolidar o
re-acomodar (por ejemplo, después de un cambio en las reglas de `sheeter.py`):

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
| `Hoja1` | Alimentos y Bebidas, con columnas derivadas (ver tabla abajo). |
| `Casa club` | Ventas con Área de negocio = Casa club. |
| `Campo de golf` | Ventas con Área de negocio = Campo de golf. |
| `Gastos` | Captura manual: **se hereda** del `DATUM.xlsx` anterior en cada corrida (nunca se pisa; captúrala directo en Excel). |

El Área de negocio se deriva de `Area`/`Grupo` con reglas fijas
(`AREA_OVERRIDE`, `Proshop` se reparte según si el grupo empieza con
`"CAMPO DE GOLF"`, etc. — ver `derive_area_negocio()`).

### `Hoja1`: columnas derivadas

| Columna | De dónde sale | Para qué sirve |
|---|---|---|
| `Area` | Tal cual viene de Datum, **sin tocar**. | Evidencia de en qué caja se punchó la operación. |
| `Punto de Venta` | Calculada (`derive_punto_venta`) a partir de `Area` / `Tipo Conjunto` / `Producto`. | A qué punto de venta pertenece **económicamente** la línea. No siempre coincide con `Area`: un platillo del menú de Callos vendido desde la caja de Mulligan, o un "Descuento VL" cobrado en Mulligan, se le atribuyen a Callos/Vista aunque `Area` diga Mulligan. Así Tableau puede sumar por punto de venta sin inflar la caja donde se punchó. |
| `Tipo` / `Subtipo` | Reglas de texto sobre `Tipo Conjunto` (bebidas, alimentos, tabaco, descuentos, eventos, clases golf, modificadores; con/sin alcohol). | Reemplaza las fórmulas de Excel de `DATUM_bueno.xlsx` (aquí se calculan en Python porque openpyxl escribe fórmulas sin valor en caché, y Tableau lee la caché). |

El punto de venta "Restaurante Vista del Lago" es el único cuyo nombre en
`Area`/`Tipo Conjunto` (`"Vista..."`) difiere del nombre completo que se
escribe en `Punto de Venta` — ver `TIPO_CONJUNTO_A_PV` / `AREA_A_PV` /
`PRODUCTO_DESCUENTO_A_PV` en `sheeter.py` si hay que agregar otro punto de
venta o ajustar una regla.

### Validar el resultado

`probar_puntoventa.py` compara el resumen por `Punto de Venta` (tomando la
columna directo, y también re-derivándola desde cero) contra los totales de
agosto 2026 ya auditados a mano, y marca OK/DIFERENCIA por punto de venta:

```bash
python probar_puntoventa.py --src DATUM.xlsx --sheet Hoja1
```

Si el total general no cuadra, algo se está perdiendo o duplicando en el
pipeline; si el total cuadra pero un punto de venta individual no, es que se
movió dinero entre puntos de venta (revisa las reglas de clasificación, no la
suma).

## Qué esperar como resultado final

Una corrida completa y exitosa de `main.py` (sin `--no-consolidate` ni
`--no-layout`) termina con:

1. **`downloads/`** con un `reporte_ventas_YYYY-MM-DD.xlsx` por cada día del
   rango pedido, y en consola `OK: los <N> archivos están en downloads/`
   (si faltó alguno, lo dice explícito y el comando termina con código de
   salida 1 — se puede volver a correr tal cual).
2. **`DATUM_unificado.xlsx`** — todos esos días juntos en un solo libro, sin
   pérdidas ni (por defecto) deduplicación.
3. **`DATUM.xlsx`** — el libro final, con las hojas `Hoja1` (Alimentos y
   Bebidas), `Casa club`, `Campo de golf` y `Gastos`. En consola se ve
   cuántas filas cayó en cada hoja, por ejemplo:
   ```
       Hoja1: 164,290 filas
       Casa club: 0 filas
       Campo de golf: 0 filas
   ```
4. Ese `DATUM.xlsx` es el que Tableau lee vía las 4 fuentes de datos del
   `.twbx` — con "Actualizar todos los extractos" el dashboard queda al día
   (ver sección **Tableau**).

Si algún paso falla (descarga incompleta, consolidación, acomodo), el
comando imprime el motivo y termina con código de salida distinto de 0; no
sobrescribe el `DATUM.xlsx` anterior a medias.

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
probar_puntoventa.py   Valida la columna Punto de Venta contra totales auditados
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
- **Un punto de venta no cuadra contra lo auditado**: corre
  `probar_puntoventa.py` (ver sección **Validar el resultado**) para ver si
  es un problema de suma total o solo de a qué punto de venta se le atribuyó
  la línea.
