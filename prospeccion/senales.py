"""Funciones puras para analizar el HTML de una página (sin red, fáciles de testear)."""

import re
from datetime import date
from urllib.parse import urljoin, urlparse

from bs4 import BeautifulSoup

from config import PALABRAS_SUBPAGINAS, PREFIJOS_EMAIL_GENERICOS, SENALES

RE_EMAIL = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
RE_COPYRIGHT = re.compile(r"(?:©|&copy;|copyright)\s*(?:\d{4}\s*[-–]\s*)?(\d{4})", re.I)

_SENALES_COMPILADAS = {
    nombre: [re.compile(p) for p in patrones] for nombre, patrones in SENALES.items()
}


def emails_genericos(texto: str) -> list[str]:
    """Devuelve solo emails corporativos genéricos (info@, ventas@, ...)."""
    encontrados = {e.lower().rstrip(".") for e in RE_EMAIL.findall(texto)}
    return sorted(e for e in encontrados if e.split("@")[0] in PREFIJOS_EMAIL_GENERICOS)


def analizar_pagina(html: str, url: str) -> dict:
    """Extrae señales de una página. Devuelve un diccionario serializable."""
    soup = BeautifulSoup(html, "lxml")
    texto = soup.get_text(" ", strip=True).lower()

    hrefs = [a.get("href", "") for a in soup.find_all("a")]
    mailtos = " ".join(h[7:] for h in hrefs if h.lower().startswith("mailto:"))

    anios = [int(a) for a in RE_COPYRIGHT.findall(html)]
    anios = [a for a in anios if 1995 <= a <= date.today().year]

    return {
        "url": url,
        "https": url.startswith("https://"),
        "senales": [n for n, regs in _SENALES_COMPILADAS.items()
                    if any(r.search(texto) for r in regs)],
        "emails": emails_genericos(texto + " " + mailtos),
        "responsive": soup.find("meta", attrs={"name": "viewport"}) is not None,
        "formulario": soup.find("form") is not None,
        "whatsapp": any("wa.me" in h or "whatsapp" in h for h in hrefs),
        "anio_copyright": max(anios) if anios else None,
        "plataforma": _detectar_plataforma(html),
    }


def links_interesantes(html: str, url_base: str, maximo: int) -> list[str]:
    """Links del mismo dominio que parecen 'contacto', 'nosotros', 'calidad', etc."""
    soup = BeautifulSoup(html, "lxml")
    dominio = urlparse(url_base).netloc.removeprefix("www.")
    resultado: list[str] = []
    for a in soup.find_all("a", href=True):
        destino = urljoin(url_base, a["href"]).split("#")[0]
        p = urlparse(destino)
        if p.scheme not in ("http", "https") or p.netloc.removeprefix("www.") != dominio:
            continue
        pista = (p.path + " " + a.get_text(" ")).lower()
        if any(k in pista for k in PALABRAS_SUBPAGINAS) and destino not in resultado:
            resultado.append(destino)
        if len(resultado) >= maximo:
            break
    return resultado


def _detectar_plataforma(html: str) -> str | None:
    h = html.lower()
    for marca, pista in [("wordpress", "wp-content"), ("wix", "wix.com"),
                         ("shopify", "cdn.shopify"), ("jimdo", "jimdo"),
                         ("squarespace", "squarespace")]:
        if pista in h:
            return marca
    return None
