import argparse
import sys
from datetime import date, timedelta
from pathlib import Path

from browser_downloader import (
    MODE_UI, MODE_URL, DOWNLOAD_DIR, month_chunks, range_download
)

INICIO_HISTORICO = date(2025, 1, 1)

def parse_fecha(txt):
    return date.fromisoformat(txt)

def consolidate(results):
    return

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--mode", choices=[MODE_URL, MODE_UI], default=MODE_URL,
                    help="url = navega directo al GET (rápido);"
                        "ui = llena el formulario a clicks")
    ap.add_argument("--from", dest="fecha_inicio", type=parse_fecha, default=INICIO_HISTORICO)
    ap.add_argument("--to", type=parse_fecha, default=None,
                    help="por defecto, ayer")
    ap.add_argument("--headless", action="store_true")
    ap.add_argument("--exit", type=Path, default=DOWNLOAD_DIR)
    ap.add_argument("--rewrite", action="store_true", 
                    help="rebaja meses que ya estén en disco")
    ap.add_argument("--retries", type=int, default=3)
    ap.add_argument("--slow-mo", type=int, default=0,
                    help="ms de pausa entre acciones, para ver el recorrido")
    ap.add_argument("--consolidate", action="store_true",
                    help="al terminar, arma DATUM.xlsx con los .xlsx bajados")
    args = ap.parse_args()

    end = args.to or (date.today() - timedelta(days=1))
    if args.fecha_inicio > end:
        print(f"Rango vacío: {args.fecha_inicio} > {end}")
        return 1

    coverage = len(month_chunks(args.fecha_inicio, end))
    print(f"Rango: {args.fecha_inicio} -> {end} (ayer) | {coverage} archivos esperados\n")

    results = range_download(
        start=args.fecha_inicio, end=end, mode=args.mode, headless=args.headless,
        download_dir=args.exit, retries=args.retries,
        rewrite=args.rewrite, slow_mo=args.slow_mo
    )

    newones = sum(1 for r in results if not r.already_exists)
    print(f"\n{'='*56}")
    print(f"Esperados   : {coverage}")
    print(f"En disco    : {len(results)} ({newones} nuevos,)"
          f"{len(results) - newones} ya estaban)")

    if len(results) != newones:
        missing = {m[0].strftime("%Y-%m") for m in month_chunks(args.fecha_inicio, end)}
        missing -= {r.tag for r in results}
        print(f"FALTAN {coverage - len(results)}: {', '.join(sorted(missing))}")
        print("Vuelve a correr el comando: los meses ya bajados se saltan.")
        return 1

    print(f"OK: los {coverage} archivos están en {args.exit.resolve()}")
    if args.consolidate:
        print("\nConsolidando...")
        consolidate(results)
    return 0

if __name__ == "__main__":
    sys.exit(main())