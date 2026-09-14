"""
Prueba de la clasificacion de PuntoVenta contra la tabla auditada de agosto 2026.

Uso:
    python probar_puntoventa.py                         # prueba DATUM.xlsx / Hoja1
    python probar_puntoventa.py --src DATUM_unificado.xlsx --sheet ventas
    python probar_puntoventa.py --src "Reporte mes agosto.xlsx" --header 4 --sin-filtro-fecha

Que hace:
  1. Carga la fuente y (por defecto) filtra agosto 2026 por la columna de fecha.
  2. Calcula el resumen por Punto de Venta de dos formas:
       A) confiando en la columna 'Punto de Venta' si el archivo ya la trae;
       B) re-derivandola con las reglas de negocio (descuento->Producto,
          Servicio a Domicilio->Area, Tabaco->Area, normal->Tipo Conjunto).
  3. Compara ambas contra los totales validados del reporte detallado y
     marca OK / DIFERENCIA por punto de venta.
"""

from __future__ import annotations

import argparse
import sys

import pandas as pd

# Totales validados de agosto 2026 (reporte detallado, reglas de negocio).
OBJETIVO = {
    "Callos": 31129.48,
    "Carrito 1": 269600.28,
    "Carrito 2": 232850.99,
    "Hoyo 10": 267879.98,
    "Mulligan": 982906.62,
    "Servicio a Domicilio": 11958.59,
    "Sushi": 359647.98,
    "Restaurante Vista del Lago": 768418.47,
    "Total": 2924392.39,
}

PRODUCTO_DESCUENTO_A_PV = {
    "Descuento": "Servicio a Domicilio",
    "Descuento Servicio": "Servicio a Domicilio",
    "Descuento Carrito 1": "Carrito 1",
    "Descuento Carrito 2": "Carrito 2",
    "Descuento Hoyo 10": "Hoyo 10",
    "Descuento MN": "Mulligan",
    "Descuento Mulligan": "Mulligan",
    "Descuento VL": "Restaurante Vista del Lago",
    "Descuento VL (Mulligan)": "Restaurante Vista del Lago",
    "Descuento Vista lago": "Restaurante Vista del Lago",
    "Descuento CS": "Callos",
    "Descuento Callos": "Callos",
    "Descuento SU": "Sushi",
    "Descuento Sushi": "Sushi",
}
TIPO_CONJUNTO_PREFIJOS = {
    "Mulligan": "Mulligan",
    "Sushi": "Sushi",
    "Vista": "Restaurante Vista del Lago",
    "Hoyo 10": "Hoyo 10",
    "Carrito 1": "Carrito 1",
    "Carrito 2": "Carrito 2",
    "Callos": "Callos",
}
AREA_A_PV = {
    "Restaurante Vista del Lago": "Restaurante Vista del Lago",
    "Callos de cortes": "Callos",
    "Mulligan": "Mulligan",
    "Sushi": "Sushi",
    "Hoyo 10": "Hoyo 10",
    "Carrito 1": "Carrito 1",
    "Carrito 2": "Carrito 2",
    "Servicio a Domicilio": "Servicio a Domicilio",
}
AREAS_EXCLUIDAS = {"Proshop", "Estética", "Estetica", "Performance Lab"}


def _norm(serie: pd.Series) -> pd.Series:
    return (serie.fillna("").astype(str).str.strip()
            .str.replace(r"\s+", " ", regex=True))


def clasificar_fila(area: str, tipo_conjunto: str, producto: str) -> str:
    tc = tipo_conjunto.strip()
    if producto.lower().startswith("descuento"):
        return PRODUCTO_DESCUENTO_A_PV.get(producto, AREA_A_PV.get(area, area))
    if area == "Servicio a Domicilio":
        return "Servicio a Domicilio"
    if tc.lower() == "tabaco":
        return AREA_A_PV.get(area, area)
    for prefijo, pv in TIPO_CONJUNTO_PREFIJOS.items():
        if tc.startswith(prefijo):
            return pv
    return AREA_A_PV.get(area, area)


def resumen(df: pd.DataFrame, usar_columna_puntoventa: bool) -> pd.DataFrame:
    df = df.copy()
    for c in ("Area", "Producto", "Tipo"):
        if c in df:
            df[c] = _norm(df[c])
    df["Precio"] = pd.to_numeric(df["Precio"], errors="coerce").fillna(0)

    # tipo COMPLETO: en el reporte detallado es 'Tipo'; en DATUM.xlsx es 'Tipo Conjunto'
    tcol = "Tipo Conjunto" if "Tipo Conjunto" in df.columns else "Tipo"
    df[tcol] = _norm(df[tcol])

    df = df[df["Area"].ne("") & ~df["Area"].isin(AREAS_EXCLUIDAS)].copy()

    if usar_columna_puntoventa and "Punto de Venta" in df.columns:
        df["PV"] = _norm(df["Punto de Venta"])
        modo = "columna Punto de Venta"
    else:
        df["PV"] = [clasificar_fila(a, t, p)
                    for a, t, p in zip(df["Area"], df[tcol], df["Producto"])]
        modo = f"re-derivado desde '{tcol}'"

    df["EsDescuento"] = df["Producto"].str.lower().str.startswith("descuento")
    df["Ventas"] = df["Precio"].where(~df["EsDescuento"], 0)
    df["Descuento"] = df["Precio"].where(df["EsDescuento"], 0)

    r = (df.groupby("PV")
           .agg(Ventas=("Ventas", "sum"), Descuento=("Descuento", "sum"),
                Total=("Precio", "sum"))
           .round(2).reset_index().rename(columns={"PV": "Punto de Venta"}))
    r.loc[len(r)] = ["Total", r["Ventas"].sum(), r["Descuento"].sum(),
                     r["Total"].sum()]
    r.attrs["modo"] = modo
    return r


def comparar_con_objetivo(r: pd.DataFrame) -> bool:
    got = dict(zip(r["Punto de Venta"], r["Total"]))
    ok = True
    ancho = max(22, max(len(pv) for pv in set(OBJETIVO) | set(got)))
    print(f"    {'Punto de Venta':{ancho}} {'obtenido':>15} {'objetivo':>15}   estado")
    print(f"    {'-'*ancho} {'-'*15} {'-'*15}   ------")
    for pv, exp in OBJETIVO.items():
        g = got.get(pv)
        if g is None:
            print(f"    {pv:{ancho}} {'(ausente)':>15} {exp:>15,.2f}   FALTA")
            ok = False
            continue
        estado = "OK" if abs(g - exp) < 0.005 else "DIFERENCIA"
        if estado != "OK":
            ok = False
        print(f"    {pv:{ancho}} {g:>15,.2f} {exp:>15,.2f}   {estado}")
    extra = set(got) - set(OBJETIVO)
    for pv in sorted(extra):
        print(f"    {pv:{ancho}} {got[pv]:>15,.2f} {'(no esperado)':>15}   REVISAR")
        ok = False
    return ok


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--src", default="DATUM.xlsx")
    ap.add_argument("--sheet", default="Hoja1",
                    help="hoja a leer (Hoja1 para DATUM.xlsx, 'ventas' para el unificado)")
    ap.add_argument("--header", type=int, default=0,
                    help="fila de encabezados (4 para los reportes crudos de Datum)")
    ap.add_argument("--sin-filtro-fecha", action="store_true",
                    help="no filtrar agosto (util si la fuente ya es solo agosto)")
    ap.add_argument("--inicio", default="2026-08-01")
    ap.add_argument("--fin", default="2026-09-01")
    a = ap.parse_args()

    try:
        df = pd.read_excel(a.src, sheet_name=a.sheet, header=a.header)
    except Exception as e:
        print(f"No pude leer {a.src} (hoja {a.sheet!r}): {e}")
        return 2

    if "Precio" not in df.columns:
        print(f"La fuente no tiene columna 'Precio'. Columnas: {list(df.columns)}")
        return 2

    col_fecha = "periodo" if "periodo" in df.columns else (
        "Fecha" if "Fecha" in df.columns else None)
    if not a.sin_filtro_fecha and col_fecha:
        f = pd.to_datetime(df[col_fecha], errors="coerce")
        df = df[(f >= pd.Timestamp(a.inicio)) & (f < pd.Timestamp(a.fin))].copy()
        print(f"Fuente : {a.src}  (hoja {a.sheet!r})")
        print(f"Filtro : {col_fecha} en [{a.inicio}, {a.fin})  ->  {len(df):,} filas")
    else:
        print(f"Fuente : {a.src}  (hoja {a.sheet!r})  ->  {len(df):,} filas (sin filtro de fecha)")
    print(f"SUM(Precio) crudo : {pd.to_numeric(df['Precio'], errors='coerce').sum():,.2f}")

    todo_ok = True
    for usar_col in (True, False):
        r = resumen(df, usar_columna_puntoventa=usar_col)
        print(f"\n=== Resumen por Punto de Venta [{r.attrs['modo']}] ===")
        print(r.to_string(index=False))
        print("\n  Comparacion contra los totales validados de agosto:")
        todo_ok &= comparar_con_objetivo(r)
        if usar_col and "Punto de Venta" not in df.columns:
            print("  (el archivo no trae columna Punto de Venta; este bloque = el siguiente)")
            break

    print()
    if todo_ok:
        print("RESULTADO: OK  -  todos los puntos de venta cuadran al centavo.")
        return 0
    print("RESULTADO: hay diferencias (ver arriba).")
    return 1


if __name__ == "__main__":
    sys.exit(main())
