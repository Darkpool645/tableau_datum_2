import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

from common import MODE_UI, MODE_URL, DOWNLOAD_DIR, day_chunks
from consolidator import consolidate
from sheeter import layout

INICIO_HISTORICO = date(2025, 1, 1)

def parse_fecha(txt):
    return date.fromisoformat(txt)

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--mode", choices=[MODE_URL, MODE_UI], default=MODE_URL,
                    help="url = fetch puro por HTTP (rápido, sin navegador,"
                        "sirve en un servidor sin gráficos); "
                        "ui = abre un Chromium real y llena el formulario a clicks")
    ap.add_argument("--from", dest="fecha_inicio", type=parse_fecha, default=INICIO_HISTORICO)
    ap.add_argument("--to", type=parse_fecha, default=None,
                    help="por defecto, ayer")
    ap.add_argument("--headless", action="store_true",
                    help="solo aplica con --mode ui")
    ap.add_argument("--exit", type=Path, default=DOWNLOAD_DIR)
    ap.add_argument("--rewrite", action="store_true",
                    help="rebaja días que ya estén en disco")
    ap.add_argument("--retries", type=int, default=3)
    ap.add_argument("--slow-mo", type=int, default=0,
                    help="ms de pausa entre acciones, para ver el recorrido"
                        "(solo aplica con --mode ui)")
    ap.add_argument("--no-consolidate", action="store_true",
                    help="solo descargar, sin armar el archivo unificado")
    ap.add_argument("--consolidate-out", type=Path, default=Path("DATUM_unificado.xlsx"),
                    help="archivo unificado de salida (.xlsx por defecto)")
    ap.add_argument("--no-layout", action="store_true",
                    help="no acomodar el consolidado en hojas")
    ap.add_argument("--layout-out", type=Path, default=Path("DATUM.xlsx"),
                    help="libro final por hojas (Hoja1/Casa club/Campo de golf/Gastos)")
    args = ap.parse_args()

    end = args.to or (date.today() - timedelta(days=1))
    if args.fecha_inicio > end:
        print(f"Rango vacío: {args.fecha_inicio} > {end}")
        return 1

    coverage = len(day_chunks(args.fecha_inicio, end))
    print(f"Rango: {args.fecha_inicio} -> {end} (ayer) | {coverage} archivos esperados\n")

    if args.mode == MODE_UI:
        from browser_downloader import range_download
        results = range_download(
            start=args.fecha_inicio, end=end, headless=args.headless,
            download_dir=args.exit, retries=args.retries,
            rewrite=args.rewrite, slow_mo=args.slow_mo,
        )
    else:
        # import perezoso: en modo url no queremos depender de Playwright
        # para nada, así corre en un servidor sin entorno gráfico.
        from http_downloader import range_download
        results = range_download(
            start=args.fecha_inicio, end=end, download_dir=args.exit,
            retries=args.retries, rewrite=args.rewrite,
        )

    newones = sum(1 for r in results if not r.already_exists)
    print(f"\n{'='*56}")
    print(f"Esperados   : {coverage}")
    print(f"En disco    : {len(results)} ({newones} nuevos,)"
          f"{len(results) - newones} ya estaban)")

    complete = len(results) == coverage
    if not complete:
        missing = {m[0].strftime("%Y-%m-%d") for m in day_chunks(args.fecha_inicio, end)}
        missing -= {r.tag for r in results}
        print(f"FALTAN {coverage - len(results)}: {', '.join(sorted(missing))}")
        print("Vuelve a correr el comando: los días ya bajados se saltan.")
    else:
        print(f"OK: los {coverage} archivos están en {args.exit.resolve()}")

    if args.no_consolidate:
        return 0 if complete else 1

    print("\nConsolidando...")
    try:
        consolidate(args.exit, args.consolidate_out)
    except Exception as e:
        print(f"La consolidación falló: {e}")
        return 1

    if args.no_layout:
        return 0 if complete else 1

    print("\nAcomodando en hojas...")
    try:
        layout(args.consolidate_out, args.layout_out)
    except Exception as e:
        print(f"El acomodo por hojas falló: {e}")
        return 1
    return 0 if complete else 1

if __name__ == "__main__":
    sys.exit(main())