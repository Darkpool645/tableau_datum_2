from __future__ import annotations
import re

from pathlib import Path
from dataclasses import dataclass
from datetime import date, timedelta
from playwright.sync_api import sync_playwright
from playwright.sync_api import TimeoutError as PWTimeout

from config import DATUM_BASE_URL, DATUM_PASSWORD, DATUM_USER, AREAS_BLACKLIST

DOWNLOAD_DIR = Path("downloads")
MODE_UI = "ui"
MODE_URL = "url"
NAV_TIMEOUT_MS = 90 * 1000
POS_PATH = "/venta.php"
REPORT_PATH = "/venta_reporte_productos.php"
DOWNLOAD_TIMEOUT_MS = 15 * 60 * 1000

@dataclass
class DownloadedDay:
    start: date
    end: date
    file: Path
    already_exists: bool

    @property
    def tag(self) -> str:
        return self.start.strftime("%Y-%m-%d")

def day_chunks(start: date, end: date) -> list[tuple[date, date]]:
    cursor = start
    sections = []
    while cursor <= end:
        sections.append((cursor, cursor))
        cursor += timedelta(days=1)

    return sections

def file_name(start: date, end: date) -> str:
    if start == end:
        return f"reporte_ventas_{start.strftime('%Y-%m-%d')}.xlsx"
    return (f"reporte_ventas_{start.strftime('%Y-%m-%d')}"
            f"_a_{end.strftime('%Y-%m-%d')}.xlsx")

class DatumBrowser:

    def __init__(self, headless:bool = False, slow_mo: int = 0,
                 download_dir: Path = DOWNLOAD_DIR, channel: str = "chrome"):
        self.headless = headless
        self.slow_mo = slow_mo
        self.download_dir = Path(download_dir)
        self.channel = channel
        self._pw = None
        self._pw_manager = None
        self.browser = None
        self.context = None
        self.page = None
        self.areas: list[str] = []

    def __enter__(self):
        self.download_dir.mkdir(parents=True, exist_ok=True)
        self._pw_manager = sync_playwright()
        self._pw = self._pw_manager.start()
        self.browser = self._pw.chromium.launch(
            channel=self.channel, headless=self.headless, slow_mo=self.slow_mo,
        )
        self.context = self.browser.new_context(accept_downloads=True)
        self.context.set_default_timeout(NAV_TIMEOUT_MS)
        self.page = self.context.new_page()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        for obj in (self.context, self.browser):
            try:
                obj and obj.close()
            except Exception:
                pass
        if self._pw_manager:
            self._pw_manager.__exit__(exc_type, exc_val, exc_tb)
        return False


    def login(self):
        if not (DATUM_USER and DATUM_PASSWORD and DATUM_BASE_URL):
            raise RuntimeError(
                "Faltan DATUM_USERNAME / DATUM_PASSWORD / DATUM_BASE en .env"
            )
        self.page.goto(DATUM_BASE_URL, wait_until="domcontentloaded")

        if self.page.locator('input[name="frUsuario"]').count() == 0:
            print("     ya había una sesión activa, la reutilizo")
            return

        self.page.fill('input[name="frUsuario"]', DATUM_USER)
        self.page.fill('input[name="frContrasena"]', DATUM_PASSWORD)
        self.page.click('button[value="Ingresar"]')
        self.page.wait_for_load_state("domcontentloaded")

        if self.page.locator('input[name="frContrasena"]').count():
            raise RuntimeError(
                "El login no pasó: sigue mostrándose el formulario."
                "Revisa usuario/contraseña."
            )
        print("     login OK")

    def go_to_pos(self):
        try:
            trigger = self.page.locator('ul.sf-menu > li > a[href="#"]').first
            trigger.hover(timeout=5000)
            trigger.click(timeout=5000)
            self.page.click('ul.sf-menu a[href$="venta.php"]', timeout=5000)
            self.page.wait_for_load_state("domcontentloaded")
        except PWTimeout:
            print("     menú no respondió, navego directo a venta.php")
            self.page.goto(f"{DATUM_BASE_URL}{POS_PATH}",
                           wait_until="domcontentloaded")

    def go_to_sales_report(self):
        try:
            trigger = self.page.locator('ul.sf-menu > li:has(cufon[alt="Reportes"]) > a').first
            trigger.hover(timeout=5000)
            trigger.click(timeout=5000)
            self.page.click('ul.sf-menu a[href*="venta_reporte_productos.php"]', timeout=5000)
            self.page.wait_for_load_state("domcontentloaded")
        except PWTimeout:
            print("     menú no respondió, navego directo al reporte")
            self.page.goto(f"{DATUM_BASE_URL}{REPORT_PATH}",
                           wait_until="domcontentloaded")
        self.page.wait_for_selector('select[name="frArea[]"]')

    def read_areas(self) -> list[str]:
            options = self.page.eval_on_selector_all(
                'select[name="frArea[]"] option',
                "els => els.map(e => ({value: e.value, label: e.textContent.trim()}))",
            )
            selecction, out = [], []
            for o in options:
                if not o["value"]:
                    continue
                if any(x in o["label"].lower() for x in AREAS_BLACKLIST):
                    out.append(o["label"])
                    continue
                selecction.append(o["value"])
            if not selecction:
                raise RuntimeError("El select de Área vino vacío.")
            print(f"        áreas detectadas: {len(selecction)}"
                  + (f"     (excluidas: {', '.join(out)})" if out else ""))
            self.areas = selecction
            return selecction

    def select_areas(self):
        self.page.select_option('select[name="frArea[]"]',
                                value=self.areas or self.read_areas())

    def add_period(self, start: date, end: date):
        js = """
        ([sel, val]) => {
            const el = document.querySelector(sel);
            el.removeAttribute('readonly');
            el.value = val;
            el.dispatchEvent(new Event('input', {bubbles: true}));
            el.dispatchEvent(new Event('change', {bubbles: true}));
            el.blur();
        }
        """
        self.page.evaluate(js, ['input[name="frInicio"]',
                                start.strftime("%d/%m/%Y")])
        self.page.evaluate(js, ['input[name="frFinal"]',
                                end.strftime("%d/%m/%Y")])
        self.page.keyboard.press("Escape")

        real_ini = self.page.input_value('input[name="frInicio"]')
        real_end = self.page.input_value('input[name="frFinal"]')
        expected = (start.strftime("%d/%m/%Y"), end.strftime("%d/%m/%Y"))
        if (real_ini, real_end) != expected:
            raise RuntimeError(
                f"El periodo no quedó quería {expected},"
                f"el form dice ({real_ini}, {real_end})"
            )

    def select_all(self):
        n = self.page.evaluate("""
        () => {
            const pat = /^(frTipo\\d|frStatus\\d|frTipoProducto\\d+|frDia\\d|frHora\\d+|frProducto(Vendido|Cancelado))$/;
            let n = 0;
            document.querySelectorAll('input[type=checkbox]').forEach(c => {
                if (pat.test(c.name) && !c.checked) { c.click(); n++; }
            });
            return n;
        }
        """)
        if n:
            print(f"        {n} checkbox(es) estaban apagados, los encendí")
        try:
            self.page.select_option('select[name="frReporte"]', "Totales")
        except Exception:
            pass

    def _save(self, download, destiny: Path) -> Path:
        name = download.suggested_filename or ""
        if not re.search(r"\.xlsx?$", name, re.I):
            print(f"    ojo: Datum mandó '{name}', no parece Excel")

        download.save_as(destiny)
        if destiny.stat().st_size < 1024:
            raise RuntimeError(f"{destiny.name} pesa {destiny.stat().st_size} B;"
                               "probablemente es una página de error.")
        with open(destiny, "rb") as f:
            if f.read(2) != b"PK":
                raise RuntimeError(
                    f"{destiny.name} no es un .xlsx (¿sesión expirada?)."
                )
        return destiny

    def download_day(self, start: date, end:date, destiny: Path) -> Path:
        self.go_to_sales_report()
        self.read_areas()
        self.select_areas()
        self.add_period(start, end)
        self.select_all()
        with self.page.expect_download(timeout=DOWNLOAD_TIMEOUT_MS) as dl:
            self.page.click('input[value*="Descargar a excel"], '
                            'button:has-text("Descargar a excel")',
                            timeout=DOWNLOAD_TIMEOUT_MS)

        return self._save(dl.value, destiny)

    

def range_download(start:date, end: date, mode: str = MODE_URL,
                   headless: bool = False, download_dir: Path = DOWNLOAD_DIR,
                   retries: int = 3, rewrite: bool = False,
                   slow_mo: int = 0) -> list[DownloadedDay]:

    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    sections = day_chunks(start, end)
    print(f"{len(sections)} día(s) a descargar: {start} -> {end} (modo {mode})")

    results: list[DownloadedDay] = []
    with DatumBrowser(headless=headless, download_dir=download_dir,
                      slow_mo=slow_mo) as nav:
        nav.login()
        nav.go_to_pos()
        nav.go_to_sales_report()
        nav.read_areas()

        for d_ini, d_end in sections:
            destiny = download_dir / file_name(d_ini, d_end)
            tag = d_ini.strftime("%Y-%m-%d")

            if destiny.exists() and not rewrite:
                print(f"[{tag}] ya está en disco, lo salto")
                results.append(DownloadedDay(d_ini, d_end, destiny, True))
                continue

            print(f"[{tag}]", flush=True)
            for atempt in range(1, retries + 1):
                try:
                    fn = (nav.download_day if mode == MODE_UI
                          else nav.download_day_url)
                    route = fn(d_ini, d_end, destiny)
                    kb = route.stat().st_size / 1024
                    print(f"        -> {route.name} ({kb:,.0f} KB)")
                    results.append(DownloadedDay(d_ini, d_end, route, False))
                    break
                except Exception as e:
                    destiny.unlink(missing_ok=True)
                    if atempt == retries:
                        print(f"    FALLÓ tras {retries} intentos: {e}")
                    else:
                        print(f"        intento {atempt}/{retries} falló ({e});"
                              "reintento con sesión fresca")
                        nav.login()
                        nav.go_to_pos()
                        nav.go_to_sales_report()
    return results 