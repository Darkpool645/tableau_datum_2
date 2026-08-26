from __future__ import annotations
import re

from pathlib import Path
from datetime import date
from html.parser import HTMLParser
from urllib.parse import urlencode, urljoin

import requests

from config import DATUM_BASE_URL, DATUM_PASSWORD, DATUM_USER, AREAS_BLACKLIST
from common import DOWNLOAD_DIR, MODE_URL, DownloadedDay, day_chunks, file_name

REPORT_PATH = "/venta_reporte_productos.php"
REQUEST_TIMEOUT = 90
DOWNLOAD_TIMEOUT = 15 * 60
DOWNLOAD_BTN_NAME = "frBoton"
DOWNLOAD_BTN_VALUE = "Descargar a excel"

# mismo patrón que usa el modo ui para "marcar todo" (ver browser_downloader.select_all)
FORCE_CHECK_PATTERN = re.compile(
    r"^(frTipo\d|frStatus\d|frTipoProducto\d+|frDia\d|frHora\d+|frProducto(Vendido|Cancelado))$"
)


class _FormParser(HTMLParser):
    """Junta cada <form> del HTML con sus <input>/<select><option>, para
    poder serializarlo como lo haría un submit real, sin abrir navegador."""

    def __init__(self):
        super().__init__()
        self.forms: list[dict] = []
        self._form = None
        self._select = None

    def handle_starttag(self, tag, attrs):
        d = dict(attrs)
        if tag == "form":
            self._form = {"action": d.get("action"), "inputs": [], "selects": []}
            self.forms.append(self._form)
        elif tag == "input" and self._form is not None:
            self._form["inputs"].append(d)
        elif tag == "select" and self._form is not None:
            self._select = {"name": d.get("name"), "multiple": "multiple" in d,
                            "options": []}
            self._form["selects"].append(self._select)
        elif tag == "option" and self._select is not None:
            self._select["_pending"] = d

    def handle_data(self, data):
        if self._select is not None and "_pending" in self._select:
            attrs = self._select.pop("_pending")
            self._select["options"].append({
                "value": attrs.get("value", ""),
                "label": data.strip(),
                "selected": "selected" in attrs,
            })

    def handle_endtag(self, tag):
        if tag == "select":
            self._select = None
        elif tag == "form":
            self._form = None


def _find_form(forms: list[dict], *, has_input: str | None = None,
               has_select: str | None = None) -> dict | None:
    for f in forms:
        if has_input and has_input not in {i.get("name") for i in f["inputs"]}:
            continue
        if has_select and has_select not in {s.get("name") for s in f["selects"]}:
            continue
        return f
    return None


class DatumHttpClient:
    """Descarga los reportes de Datum por HTTP puro (login + GET directo),
    sin Playwright ni navegador: sirve para correr en un servidor sin
    entorno gráfico."""

    def __init__(self):
        if not (DATUM_USER and DATUM_PASSWORD and DATUM_BASE_URL):
            raise RuntimeError(
                "Faltan DATUM_USERNAME / DATUM_PASSWORD / DATUM_BASE en .env"
            )
        self.session = requests.Session()
        self.session.headers.update({"User-Agent": "Mozilla/5.0 (datum-downloader)"})
        self._report_action: str | None = None
        self._report_params: list[tuple[str, str]] | None = None

    def login(self):
        r = self.session.get(DATUM_BASE_URL, timeout=REQUEST_TIMEOUT)
        r.raise_for_status()

        fp = _FormParser()
        fp.feed(r.text)
        form = _find_form(fp.forms, has_input="frUsuario")
        if form is None:
            raise RuntimeError("No encontré el formulario de login en la página.")

        payload = {}
        for i in form["inputs"]:
            name = i.get("name")
            if not name:
                continue
            if name == "frUsuario":
                payload[name] = DATUM_USER
            elif name == "frContrasena":
                payload[name] = DATUM_PASSWORD
            elif i.get("type") == "checkbox":
                if "checked" in i:
                    payload[name] = i.get("value", "on")
            else:
                payload[name] = i.get("value", "")

        action = urljoin(r.url, form["action"] or "")
        r2 = self.session.post(action, data=payload, timeout=REQUEST_TIMEOUT)
        r2.raise_for_status()
        if "frContrasena" in r2.text:
            raise RuntimeError(
                "El login no pasó: sigue mostrándose el formulario."
                "Revisa usuario/contraseña."
            )
        print("     login OK")

    def _capture_report_params(self):
        """Lee el formulario de reporte una sola vez y lo serializa completo
        (áreas, checkboxes de tipo/status/producto/día/hora, grupos, botón)
        tal como lo mandaría un submit real, para reconstruir la URL GET en
        cada día sin volver a pedir la página."""
        r = self.session.get(f"{DATUM_BASE_URL}{REPORT_PATH}", timeout=REQUEST_TIMEOUT)
        r.raise_for_status()

        fp = _FormParser()
        fp.feed(r.text)
        form = _find_form(fp.forms, has_select="frArea[]")
        if form is None:
            raise RuntimeError("No encontré el formulario de reporte (¿sesión expiró?).")

        params: list[tuple[str, str]] = []
        for i in form["inputs"]:
            name = i.get("name")
            if not name:
                continue
            itype = i.get("type", "text")
            if itype == "checkbox":
                # igual que select_all() en modo ui: se marcan todos los que
                # matcheen el patrón, sin importar su estado por default.
                if FORCE_CHECK_PATTERN.match(name) or "checked" in i:
                    params.append((name, i.get("value", "1")))
            elif itype == "radio":
                if "checked" in i:
                    params.append((name, i.get("value", "")))
            elif itype in ("submit", "button"):
                continue
            else:
                params.append((name, i.get("value", "")))

        n_forced = sum(1 for k, _ in params if FORCE_CHECK_PATTERN.match(k))
        if n_forced == 0:
            raise RuntimeError(
                "El reporte no trajo checkboxes de tipo/status/producto;"
                "revisa si cambió el formulario."
            )

        areas_out = []
        for sel in form["selects"]:
            name = sel.get("name")
            if not name:
                continue
            if name == "frArea[]":
                for o in sel["options"]:
                    if not o["value"]:
                        continue
                    if any(x in o["label"].lower() for x in AREAS_BLACKLIST):
                        continue
                    params.append((name, o["value"]))
                    areas_out.append(o["value"])
            elif sel.get("multiple"):
                for o in sel["options"]:
                    if o["value"]:
                        params.append((name, o["value"]))
            else:
                chosen = next((o for o in sel["options"] if o["selected"]), None)
                if chosen is None and sel["options"]:
                    chosen = sel["options"][0]
                if chosen is not None:
                    params.append((name, chosen["value"]))

        if not areas_out:
            raise RuntimeError("El select de Área vino vacío.")
        print(f"        áreas detectadas: {len(areas_out)}")

        params.append((DOWNLOAD_BTN_NAME, DOWNLOAD_BTN_VALUE))

        self._report_action = urljoin(r.url, form["action"] or REPORT_PATH)
        self._report_params = params

    def download_day(self, start: date, end: date, destiny: Path) -> Path:
        if self._report_params is None:
            self._capture_report_params()

        params = [(k, v) for k, v in self._report_params
                  if k not in ("frInicio", "frFinal")]
        params += [
            ("frInicio", start.strftime("%d/%m/%Y")),
            ("frFinal", end.strftime("%d/%m/%Y")),
        ]
        url = f"{self._report_action}?{urlencode(params)}"

        r = self.session.get(url, timeout=DOWNLOAD_TIMEOUT, stream=True)
        r.raise_for_status()
        if "text/html" in r.headers.get("Content-Type", ""):
            raise RuntimeError(
                "Datum regresó HTML en vez de un excel (¿sesión expirada?)."
            )

        destiny.parent.mkdir(parents=True, exist_ok=True)
        with open(destiny, "wb") as f:
            for chunk in r.iter_content(chunk_size=64 * 1024):
                f.write(chunk)

        if destiny.stat().st_size < 1024:
            raise RuntimeError(f"{destiny.name} pesa {destiny.stat().st_size} B;"
                               "probablemente es una página de error.")
        with open(destiny, "rb") as f:
            if f.read(2) != b"PK":
                raise RuntimeError(
                    f"{destiny.name} no es un .xlsx (¿sesión expirada?)."
                )
        return destiny


def range_download(start: date, end: date, download_dir: Path = DOWNLOAD_DIR,
                   retries: int = 3, rewrite: bool = False) -> list[DownloadedDay]:
    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    sections = day_chunks(start, end)
    print(f"{len(sections)} día(s) a descargar: {start} -> {end}"
          f" (modo {MODE_URL}, fetch puro, sin navegador)")

    results: list[DownloadedDay] = []
    client = DatumHttpClient()
    client.login()

    for d_ini, d_end in sections:
        destiny = download_dir / file_name(d_ini, d_end)
        tag = d_ini.strftime("%Y-%m-%d")

        if destiny.exists() and not rewrite:
            print(f"[{tag}] ya está en disco, lo salto")
            results.append(DownloadedDay(d_ini, d_end, destiny, True))
            continue

        print(f"[{tag}]", flush=True)
        for attempt in range(1, retries + 1):
            try:
                route = client.download_day(d_ini, d_end, destiny)
                kb = route.stat().st_size / 1024
                print(f"        -> {route.name} ({kb:,.0f} KB)")
                results.append(DownloadedDay(d_ini, d_end, route, False))
                break
            except Exception as e:
                destiny.unlink(missing_ok=True)
                if attempt == retries:
                    print(f"    FALLÓ tras {retries} intentos: {e}")
                else:
                    print(f"        intento {attempt}/{retries} falló ({e});"
                          "reintento con sesión fresca")
                    client = DatumHttpClient()
                    client.login()
    return results
