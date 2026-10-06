"""Paso 2: visitar la web de cada empresa con Crawlee y extraer señales.

Lee data/empresas.csv, visita la página de inicio y hasta MAX_SUBPAGINAS internas
(contacto, nosotros, calidad, trabaja con nosotros...).

Uso:
    python paso2_enriquecer.py             # todas las empresas con web
    python paso2_enriquecer.py --limite 20 # prueba rápida

Resultado: data/paginas.json (una fila por página visitada).
"""

import argparse
import asyncio
import csv
import json
from datetime import timedelta
from pathlib import Path

from crawlee import ConcurrencySettings, Request
from crawlee.crawlers import BeautifulSoupCrawler, BeautifulSoupCrawlingContext

from config import DOMINIOS_NO_PROPIOS, MAX_SUBPAGINAS
from senales import analizar_pagina, links_interesantes

ENTRADA = Path("data/empresas.csv")
SALIDA = Path("data/paginas.json")


def web_propia(url: str) -> bool:
    return bool(url) and not any(d in url.lower() for d in DOMINIOS_NO_PROPIOS)


async def main(limite: int | None) -> None:
    with ENTRADA.open(encoding="utf-8-sig") as f:
        empresas = [e for e in csv.DictReader(f, delimiter=";") if web_propia(e["web"])]
    if limite:
        empresas = empresas[:limite]
    print(f"Visitando {len(empresas)} sitios web...")

    crawler = BeautifulSoupCrawler(
        respect_robots_txt_file=True,   # respetar lo que cada sitio permite
        max_request_retries=1,
        request_handler_timeout=timedelta(seconds=30),
        concurrency_settings=ConcurrencySettings(
            desired_concurrency=5, max_concurrency=5, max_tasks_per_minute=90),
    )

    @crawler.router.handler("INICIO")
    async def inicio(ctx: BeautifulSoupCrawlingContext) -> None:
        html = await _html(ctx)
        url = ctx.request.loaded_url or ctx.request.url
        place_id = ctx.request.user_data["place_id"]
        await ctx.push_data({"place_id": place_id, "tipo": "inicio", **analizar_pagina(html, url)})
        await ctx.add_requests([
            Request.from_url(link, label="SUB", user_data={"place_id": place_id})
            for link in links_interesantes(html, url, MAX_SUBPAGINAS)
        ])

    @crawler.router.handler("SUB")
    async def subpagina(ctx: BeautifulSoupCrawlingContext) -> None:
        html = await _html(ctx)
        url = ctx.request.loaded_url or ctx.request.url
        await ctx.push_data({"place_id": ctx.request.user_data["place_id"], "tipo": "sub",
                             **analizar_pagina(html, url)})

    @crawler.failed_request_handler
    async def fallo(ctx, error: Exception) -> None:
        if ctx.request.label == "INICIO":
            await ctx.push_data({"place_id": ctx.request.user_data["place_id"], "tipo": "error",
                                 "url": ctx.request.url, "error": str(error)[:200]})

    await crawler.run([
        Request.from_url(e["web"], label="INICIO", user_data={"place_id": e["place_id"]})
        for e in empresas
    ])

    filas = (await crawler.get_data(limit=1_000_000)).items
    SALIDA.write_text(json.dumps(filas, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"\nListo: {len(filas)} páginas analizadas → {SALIDA}")


async def _html(ctx: BeautifulSoupCrawlingContext) -> str:
    return (await ctx.http_response.read()).decode("utf-8", errors="replace")


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--limite", type=int, help="procesar solo las primeras N empresas")
    asyncio.run(main(ap.parse_args().limite))
