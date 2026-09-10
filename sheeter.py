"""
Último paso del pipeline: acomoda el consolidado en hojas, como DATUM_bueno.xlsx.

Entrada : el DataFrame (o el .xlsx/.csv/.parquet) que produce consolidator.py
Salida  : un libro con 4 hojas
            Hoja1          -> Alimentos y Bebidas (+ Tipo Conjunto / Tipo /
                              Subtipo / PuntoVenta)
                              "Area" se conserva TAL CUAL viene de Datum (la
                              caja donde se registró la operación). La
                              atribución económica -a qué punto de venta
                              pertenece realmente la venta- vive en la columna
                              nueva "PuntoVenta" (ver derive_punto_venta):
                              las ventas normales se clasifican por
                              "Tipo Conjunto", el Tabaco por "Area", el
                              Servicio a Domicilio por "Area", y los descuentos
                              por "Producto". Así Tableau puede sumar por
                              PuntoVenta sin inflar la caja donde se punchó
                              (ej. platillos del menú de Callos vendidos desde
                              Mulligan, o un "Descuento VL" cobrado en Mulligan).
            Casa club      -> Area de negocio = Casa club      (+ Union)
            Campo de golf  -> Area de negocio = Campo de golf  (+ Union)
            Gastos         -> captura manual, se conserva del libro anterior

Notas de diseño
---------------
* En DATUM_bueno.xlsx las columnas Tipo y Subtipo son fórmulas de Excel. Aquí se
  calculan en Python con exactamente la misma lógica: openpyxl escribe fórmulas
  SIN valor en caché, y Tableau lee la caché, así que una fórmula recién escrita
  llegaría como NULL. El valor calculado es idéntico al que Excel cachea.
* Fecha se recalcula a partir de la columna "periodo" del consolidado (el día
  del reporte de Datum, no el timestamp original de la venta) y se escribe
  como fecha real (no texto). Los % se escriben como número con formato de
  porcentaje. Así Tableau los tipa solo.
"""

from __future__ import annotations

import re
from pathlib import Path

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

# --------------------------------------------------------------------------- #
# Reglas de negocio
# --------------------------------------------------------------------------- #
AYB = "Alimentos y Bebidas"
CASA_CLUB = "Casa club"
CAMPO_GOLF = "Campo de golf"

# Areas que no son punto de venta de A y B
AREA_OVERRIDE = {
    "Estética": CASA_CLUB,
    "Estetica": CASA_CLUB,
    "Performance Lab": CAMPO_GOLF,
}
# Proshop vende de los dos lados: se decide por el grupo de primer nivel
PROSHOP = "Proshop"
GRUPO_CAMPO_PREFIX = "CAMPO DE GOLF"

# --------------------------------------------------------------------------- #
# PuntoVenta: a qué punto de venta pertenece ECONÓMICAMENTE la operación
# --------------------------------------------------------------------------- #
# "Area" (la caja donde Datum registró la línea) NO siempre es el punto de
# venta que recibe el dinero: platillos del menú de Callos se venden desde la
# caja de Mulligan, un "Descuento VL" se cobra desde Mulligan, etc. Si Tableau
# suma por "Area", esas cajas quedan infladas. "PuntoVenta" resuelve la
# atribución económica sin tocar "Area" (que se conserva como evidencia de
# origen). Orden de prioridad -primer criterio que aplica gana-:
#   1. Descuentos      -> el nombre del Producto identifica el punto de venta.
#   2. Servicio a Domicilio (por Area) -> siempre Servicio a Domicilio.
#   3. Tabaco          -> el Tipo no dice el punto; manda "Area".
#   4. Venta normal    -> manda "Tipo Conjunto" (contiene el punto de venta).
#   5. No reconocido   -> respaldo: "Area" normalizada.
# Las siglas de Producto (MN/VL/SU/CS) NO son regla normativa: solo aparecen
# en ~80-90% de los casos, así que se usan como validación auxiliar, nunca
# como filtro final.

# Prefijos de "Tipo Conjunto" que nombran el punto de venta (venta normal).
# "Vista ..." en Tipo Conjunto == "Restaurante Vista del Lago" en Area.
TIPO_CONJUNTO_A_PV = {
    "Mulligan": "Mulligan",
    "Sushi": "Sushi",
    "Vista": "Vista",
    "Hoyo 10": "Hoyo 10",
    "Carrito 1": "Carrito 1",
    "Carrito 2": "Carrito 2",
    "Callos": "Callos",
}

# "Area" -> PuntoVenta. Se usa para Tabaco, Servicio a Domicilio y el respaldo.
AREA_A_PV = {
    "Mulligan": "Mulligan",
    "Sushi": "Sushi",
    "Restaurante Vista del Lago": "Vista",
    "Hoyo 10": "Hoyo 10",
    "Carrito 1": "Carrito 1",
    "Carrito 2": "Carrito 2",
    "Callos de cortes": "Callos",
    "Servicio a Domicilio": "Servicio a Domicilio",
}

# Descuentos: el "Producto" dice a qué punto de venta pertenece el descuento
# (se cobra desde la caja de otro punto muy seguido). Nombre completo o sigla.
PRODUCTO_DESCUENTO_A_PV = {
    "Descuento Mulligan": "Mulligan",
    "Descuento MN": "Mulligan",
    "Descuento Sushi": "Sushi",
    "Descuento SU": "Sushi",
    "Descuento Vista lago": "Vista",
    "Descuento VL": "Vista",
    "Descuento Callos": "Callos",
    "Descuento CS": "Callos",
    "Descuento Hoyo 10": "Hoyo 10",
    "Descuento Carrito 1": "Carrito 1",
    "Descuento Carrito 2": "Carrito 2",
    "Descuento Servicio": "Servicio a Domicilio",
}

AREA_SERVICIO = "Servicio a Domicilio"

# Basura que Datum mete cuando un día no trae movimientos
AREAS_BASURA = {"", "no existen registros para mostrar."}

SIN_SUBSUBGRUPO_AYB = "-"       # Hoja1 usa "-"; las otras hojas lo dejan vacío

# Tipo: se evalúa en este orden, primer match gana (igual que el IF anidado)
TIPO_RULES = [
    ("bebidas", "Bebidas"),
    ("alimentos", "Alimentos"),
    ("tabaco", "Tabaco"),
    ("descuento", "Descuentos"),
    ("evento", "Eventos"),
    ("clases golf", "Clases golf"),
    ("modificador", "Modificadores"),
]
SUBTIPO_RULES = [
    ("con alcohol", "con alcohol"),
    ("sin alcohol", "sin alcohol"),
]
SUBTIPO_DEFAULT = "-"

HOJA_AYB = "Hoja1"
HOJA_GASTOS = "Gastos"

COLS_AYB = [
    "Area", "PuntoVenta", "Grupo", "Subgrupo", "Sub subgrupo", "Producto",
    "Tipo Conjunto", "Tipo", "Subtipo", "Usuario",
    "Cantidad", "Precio", "Impuesto", "Total", "Costo", "Margen",
    "% Utilidad", "% Margen", "Fecha", "Area de negocio",
]
COLS_NEGOCIO = [
    "Area", "Grupo", "Subgrupo", "Sub subgrupo", "Producto",
    "Tipo", "Usuario",
    "Cantidad", "Precio", "Impuesto", "Total", "Costo", "Margen",
    "% Utilidad", "% Margen", "Fecha", "Area de negocio", "Union",
]

GASTOS_HEADER = ["Administración", "Gastos AyB", "Gastos campo de golf",
                 "Gastos casa club", "Fecha"]

# Formatos de celda
FMT = {
    "Cantidad": "#,##0",
    "Precio": "#,##0.00",
    "Impuesto": "#,##0.00",
    "Total": "#,##0.00",
    "Costo": "#,##0.00",
    "Margen": "#,##0.00",
    "% Utilidad": "0%",
    "% Margen": "0%",
    "Fecha": "yyyy-mm-dd",
}
FUENTE = "Arial"


# --------------------------------------------------------------------------- #
# Derivaciones
# --------------------------------------------------------------------------- #
def split_grupos(serie: pd.Series) -> pd.DataFrame:
    """
    'CAMPO DE GOLF > Rentas CG > Carritos' -> tres columnas.

    Ojo: NO se hace strip. DATUM_bueno.xlsx conserva los espacios alrededor del
    '>' ('CAMPO DE GOLF ', ' Rentas CG ', ' Carritos') y los filtros de Tableau
    ya están hechos sobre esos valores.
    """
    partes = serie.fillna("").astype(str).str.split(">", n=2, expand=True)
    for i in range(3):
        if i not in partes.columns:
            partes[i] = None
    partes = partes[[0, 1, 2]]
    partes.columns = ["Grupo", "Subgrupo", "Sub subgrupo"]
    return partes.replace({"": None})


def _match(texto: str | None, reglas, default=None):
    if not texto:
        return default
    bajo = str(texto).lower()
    for aguja, valor in reglas:
        if aguja in bajo:
            return valor
    return default


def derive_tipo(tipo_conjunto: pd.Series) -> pd.Series:
    """Equivale al IF(ISNUMBER(SEARCH(...))) anidado de la columna G."""
    return tipo_conjunto.map(lambda v: _match(v, TIPO_RULES, ""))


def derive_subtipo(tipo_conjunto: pd.Series) -> pd.Series:
    """Equivale al IF anidado de la columna H."""
    return tipo_conjunto.map(lambda v: _match(v, SUBTIPO_RULES, SUBTIPO_DEFAULT))


def _pv_por_tipo_conjunto(tc: str) -> str | None:
    """Venta normal: el prefijo de "Tipo Conjunto" nombra el punto de venta."""
    for prefijo, pv in TIPO_CONJUNTO_A_PV.items():
        if tc.startswith(prefijo):
            return pv
    return None


def _pv_fila(area: str, tipo_conjunto: str, producto: str) -> str:
    """
    Punto de venta al que pertenece ECONÓMICAMENTE una línea. No toca "Area".
    Orden de prioridad (ver comentario de TIPO_CONJUNTO_A_PV):
      1. descuento -> Producto ; 2. Servicio a Domicilio -> Area ;
      3. Tabaco -> Area ; 4. venta normal -> Tipo Conjunto ; 5. respaldo -> Area.
    """
    tc = tipo_conjunto.strip()
    tc_bajo = tc.lower()

    # 1. Descuentos: el Producto dice el punto de venta.
    if "descuento" in tc_bajo:
        pv = PRODUCTO_DESCUENTO_A_PV.get(producto.strip())
        if pv:
            return pv
        return AREA_A_PV.get(area, area)

    # 2. Servicio a Domicilio: manda el Area aunque el Tipo Conjunto diga otra cosa.
    if area == AREA_SERVICIO:
        return AREA_SERVICIO

    # 3. Tabaco: el Tipo no indica el punto; manda el Area.
    if tc_bajo == "tabaco":
        return AREA_A_PV.get(area, area)

    # 4. Venta normal: manda el Tipo Conjunto.
    pv = _pv_por_tipo_conjunto(tc)
    if pv:
        return pv

    # 5. No reconocido: respaldo controlado sobre el Area normalizada.
    return AREA_A_PV.get(area, area)


def derive_punto_venta(area: pd.Series, tipo_conjunto: pd.Series,
                       producto: pd.Series) -> pd.Series:
    """
    Columna "PuntoVenta": a qué punto de venta se atribuye económicamente cada
    línea. "Area" (la caja donde se punchó) se conserva intacta; esta función
    NO la modifica y NO borra filas. Es una pura reasignación: el dinero solo
    se mueve entre puntos de venta, nunca se crea ni se elimina.
    """
    a = area.map(lambda v: "" if pd.isna(v) else str(v).strip())
    tc = tipo_conjunto.map(lambda v: "" if pd.isna(v) else str(v))
    p = producto.map(lambda v: "" if pd.isna(v) else str(v))
    return pd.Series(
        [_pv_fila(ai, tci, pi) for ai, tci, pi in zip(a, tc, p)],
        index=area.index,
    )


def derive_area_negocio(area: pd.Series, grupo: pd.Series) -> pd.Series:
    def resolver(a, g):
        a = (a or "").strip()
        if a in AREA_OVERRIDE:
            return AREA_OVERRIDE[a]
        if a == PROSHOP:
            gu = (g or "").strip().upper()
            return CAMPO_GOLF if gu.startswith(GRUPO_CAMPO_PREFIX) else CASA_CLUB
        return AYB

    return pd.Series([resolver(a, g) for a, g in zip(area, grupo)], index=area.index)


# --------------------------------------------------------------------------- #
# Armado de las hojas
# --------------------------------------------------------------------------- #
def build_sheets(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    df = df.copy()

    # 1. tirar la basura de días sin movimientos
    area_norm = df["Area"].fillna("").astype(str).str.strip()
    basura = area_norm.str.lower().isin(AREAS_BASURA)
    if basura.any():
        print(f"    descarto {basura.sum():,} filas sin Area válida")
    df = df[~basura].copy()
    df["Area"] = area_norm[~basura]

    # 2. Grupos -> Grupo / Subgrupo / Sub subgrupo
    col_grupos = "Grupos" if "Grupos" in df.columns else "Grupo"
    df[["Grupo", "Subgrupo", "Sub subgrupo"]] = split_grupos(df[col_grupos])

    # 2.5 Fecha se toma de "periodo" (el día del reporte de Datum), no del
    #     timestamp original de la venta: Datum a veces marca una venta con
    #     hora de después de medianoche pero el reporte -y el corte de caja-
    #     al que pertenece es el que indica "periodo". El resultado es solo
    #     la fecha (sin hora).
    col_periodo = next((c for c in df.columns if c.lower() == "periodo"), None)
    fuente_fecha = col_periodo if col_periodo else "Fecha"
    df["Fecha"] = pd.to_datetime(df[fuente_fecha], errors="coerce").dt.date

    # 3. columnas derivadas
    df["Usuario"] = df["Usuario"].fillna("").astype(str).str.strip().str.upper()
    # "Area" se deja TAL CUAL vino de Datum (evidencia de origen). El Area de
    # negocio -en qué hoja cae la fila- se decide sobre ese Area original.
    df["Area de negocio"] = derive_area_negocio(df["Area"], df["Grupo"])
    df["Tipo Conjunto"] = df["Tipo"]
    # "PuntoVenta": atribución económica real (no toca Area, no borra filas).
    df["PuntoVenta"] = derive_punto_venta(
        df["Area"], df["Tipo Conjunto"], df["Producto"]
    )
    df["Union"] = 1

    if "Fecha" in df:
        df = df.sort_values("Fecha", kind="stable")

    hojas: dict[str, pd.DataFrame] = {}

    # --- Hoja1: Alimentos y Bebidas. "Area" = caja de origen (sin tocar);
    #     "PuntoVenta" = punto de venta al que pertenece el dinero. ---
    ayb = df[df["Area de negocio"] == AYB].copy()
    ayb["Tipo"] = derive_tipo(ayb["Tipo Conjunto"])
    ayb["Subtipo"] = derive_subtipo(ayb["Tipo Conjunto"])
    ayb["Sub subgrupo"] = ayb["Sub subgrupo"].fillna(SIN_SUBSUBGRUPO_AYB)
    hojas[HOJA_AYB] = ayb[COLS_AYB]

    # --- Casa club / Campo de golf: conservan el Tipo original ---
    for hoja in (CASA_CLUB, CAMPO_GOLF):
        hojas[hoja] = df[df["Area de negocio"] == hoja][COLS_NEGOCIO].copy()

    for nombre, hoja in hojas.items():
        print(f"    {nombre}: {len(hoja):,} filas")
    return hojas


# --------------------------------------------------------------------------- #
# Gastos (captura manual: se hereda, nunca se pisa)
# --------------------------------------------------------------------------- #
def read_gastos(fuente: Path | None) -> list[tuple]:
    if not fuente or not Path(fuente).exists():
        return []
    try:
        wb = load_workbook(fuente, read_only=True, data_only=True)
    except Exception as e:
        print(f"    no pude abrir {fuente} para heredar Gastos: {e}")
        return []
    if HOJA_GASTOS not in wb.sheetnames:
        wb.close()
        return []
    filas = [r for r in wb[HOJA_GASTOS].iter_rows(min_row=2, max_col=5, values_only=True)
             if any(v is not None for v in r)]
    wb.close()
    print(f"    Gastos: heredo {len(filas)} filas de {Path(fuente).name}")
    return filas


# --------------------------------------------------------------------------- #
# Escritura
# --------------------------------------------------------------------------- #
def _ancho(serie: pd.Series, nombre: str) -> int:
    crudo = serie.astype(str).str.len().head(2000).max()
    contenido = int(crudo) if pd.notna(crudo) else 10
    return max(10, min(38, contenido + 2), len(nombre) + 2)


def _escribe_hoja(wb: Workbook, nombre: str, df: pd.DataFrame) -> None:
    ws = wb.create_sheet(nombre)
    ws.freeze_panes = "A2"

    negritas = Font(name=FUENTE, bold=True)
    normal = Font(name=FUENTE)

    cabecera = []
    for col in df.columns:
        c = WriteOnlyCell(ws, value=col)
        c.font = negritas
        cabecera.append(c)
    ws.append(cabecera)

    for i, col in enumerate(df.columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = _ancho(df[col], col)

    formatos = [FMT.get(c) for c in df.columns]
    for fila in df.itertuples(index=False, name=None):
        celdas = []
        for valor, fmt in zip(fila, formatos):
            if valor is None or valor is pd.NaT or (isinstance(valor, float) and pd.isna(valor)):
                celdas.append(None)
                continue
            if isinstance(valor, pd.Timestamp):
                valor = valor.to_pydatetime()
            if fmt:
                c = WriteOnlyCell(ws, value=valor)
                c.number_format = fmt
                c.font = normal
                celdas.append(c)
            else:
                celdas.append(valor)
        ws.append(celdas)


def _escribe_gastos(wb: Workbook, filas: list[tuple]) -> None:
    ws = wb.create_sheet(HOJA_GASTOS)
    negritas = Font(name=FUENTE, bold=True)

    cabecera = []
    for nombre in GASTOS_HEADER:
        c = WriteOnlyCell(ws, value=nombre)
        c.font = negritas
        cabecera.append(c)
    ws.append(cabecera)

    for i, nombre in enumerate(GASTOS_HEADER, 1):
        ws.column_dimensions[get_column_letter(i)].width = max(14, len(nombre) + 2)

    for fila in filas:
        celdas = []
        for j, valor in enumerate(fila):
            c = WriteOnlyCell(ws, value=valor)
            c.number_format = "yyyy-mm-dd" if j == 4 else "#,##0.00"
            c.font = Font(name=FUENTE)
            celdas.append(c)
        ws.append(celdas)


def write_workbook(hojas: dict[str, pd.DataFrame], out: Path,
                   gastos: list[tuple]) -> Path:
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)

    for nombre, df in hojas.items():
        if len(df) > 1_048_575:
            raise RuntimeError(
                f"La hoja '{nombre}' trae {len(df):,} filas y no cabe en Excel "
                "(límite 1,048,575)."
            )

    wb = Workbook(write_only=True)
    for nombre in (HOJA_AYB, CASA_CLUB, CAMPO_GOLF):
        _escribe_hoja(wb, nombre, hojas[nombre])
    _escribe_gastos(wb, gastos)
    wb.save(out)
    return out


# --------------------------------------------------------------------------- #
# API pública
# --------------------------------------------------------------------------- #
def _load(src) -> pd.DataFrame:
    if isinstance(src, pd.DataFrame):
        return src
    src = Path(src)
    suffix = src.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(src, parse_dates=["Fecha"])
    if suffix == ".parquet":
        return pd.read_parquet(src)
    if suffix == ".xlsx":
        return pd.read_excel(src, sheet_name=0)
    raise ValueError(f"Formato no soportado: {suffix}")


def layout(src, out: Path = Path("DATUM_final.xlsx"),
           gastos_from: Path | None = None) -> Path:
    """
    Acomoda el consolidado en hojas y escribe el libro final.

    src         : DataFrame o ruta al consolidado (.xlsx/.csv/.parquet)
    out         : libro final
    gastos_from : de dónde heredar la hoja Gastos (por defecto, el propio `out`
                  si ya existe: así la captura manual sobrevive cada corrida)
    """
    df = _load(src)
    hojas = build_sheets(df)

    fuente_gastos = gastos_from if gastos_from is not None else out
    gastos = read_gastos(fuente_gastos)
    if not gastos:
        print("    Gastos: sin datos previos, dejo solo los encabezados "
              "(captúralos en Excel y la próxima corrida los conserva)")

    destino = write_workbook(hojas, Path(out), gastos)
    total = sum(len(h) for h in hojas.values())
    print(f"\n{total:,} filas repartidas en {len(hojas) + 1} hojas -> {destino.resolve()}")
    return destino


if __name__ == "__main__":
    import argparse

    ap = argparse.ArgumentParser(description="Acomoda el consolidado en hojas")
    ap.add_argument("--src", type=Path, default=Path("DATUM.xlsx"))
    ap.add_argument("--out", type=Path, default=Path("DATUM_final.xlsx"))
    ap.add_argument("--gastos-from", type=Path, default=None,
                    help="libro del que se hereda la hoja Gastos (default: --out)")
    a = ap.parse_args()
    layout(a.src, a.out, a.gastos_from)