"""Paso 1: buscar empresas en Google Maps usando la Places API (New).

No se hace scraping de Google Maps (lo prohíben sus términos); se usa la API oficial.

Uso:
    python paso1_descubrir.py                     # todos los rubros x comunas (pide confirmación)
    python paso1_descubrir.py --rubros maestranza --comunas Quilpué "Villa Alemana"
    python paso1_descubrir.py --max-consultas 20  # tope de llamadas a la API
    python paso1_descubrir.py --sectores          # usar las búsquedas de SECTORES en config.py

Resultado: data/empresas.csv (se acumula entre ejecuciones, sin duplicados).
"""

import argparse
import csv
import os
import sys
import time
from pathlib import Path

import httpx
from dotenv import load_dotenv

from config import CATEGORIAS_EXCLUIDAS, NOMBRES_EXCLUIDOS, COMUNAS, FILTRO_UBICACION, RUBROS, SECTORES, ZONA

URL = "https://places.googleapis.com/v1/places:searchText"
CAMPOS = ",".join([
    "places.id", "places.displayName", "places.formattedAddress",
    "places.websiteUri", "places.nationalPhoneNumber", "places.rating",
    "places.userRatingCount", "places.primaryTypeDisplayName",
    "places.businessStatus", "places.googleMapsUri", "nextPageToken",
])
SALIDA = Path("data/empresas.csv")
HECHAS = Path("data/busquedas_hechas.txt")  # para no repetir (ni pagar) la misma búsqueda
COLUMNAS = ["place_id", "nombre", "sector", "rubro_busqueda", "categoria_maps", "direccion",
            "comuna_busqueda", "telefono", "web", "rating", "n_resenas", "maps_url"]


def buscar(cliente: httpx.Client, texto: str, paginas: int) -> tuple[list[dict], int]:
    """Devuelve (lugares, n_llamadas). Cada página trae hasta 20 resultados."""
    lugares, token, llamadas = [], None, 0
    for _ in range(paginas):
        cuerpo = {"textQuery": texto, "languageCode": "es", "regionCode": "CL",
                  "pageSize": 20, "locationRestriction": {"rectangle": ZONA}}
        if token:
            cuerpo["pageToken"] = token
        r = cliente.post(URL, json=cuerpo)
        llamadas += 1
        r.raise_for_status()
        datos = r.json()
        lugares += datos.get("places", [])
        token = datos.get("nextPageToken")
        if not token:
            break
        time.sleep(1)
    return lugares, llamadas


def excluida(categoria: str, nombre: str = "") -> bool:
    c, n = categoria.lower(), nombre.lower()
    return any(x in c for x in CATEGORIAS_EXCLUIDAS) or any(x in n for x in NOMBRES_EXCLUIDOS)


def cargar_existentes() -> dict[str, dict]:
    if not SALIDA.exists():
        return {}
    with SALIDA.open(encoding="utf-8-sig") as f:
        return {fila["place_id"]: fila for fila in csv.DictReader(f, delimiter=";")}


def guardar(empresas: dict[str, dict]) -> None:
    SALIDA.parent.mkdir(exist_ok=True)
    # utf-8-sig y ";" para que Excel en español lo abra bien
    with SALIDA.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNAS, delimiter=";")
        w.writeheader()
        w.writerows(empresas.values())


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--rubros", nargs="+", default=RUBROS)
    ap.add_argument("--comunas", nargs="+", default=COMUNAS)
    ap.add_argument("--paginas", type=int, default=1, choices=[1, 2, 3],
                    help="páginas de 20 resultados por búsqueda (más = más costo)")
    ap.add_argument("--max-consultas", type=int, default=100,
                    help="tope de llamadas a la API en esta ejecución")
    ap.add_argument("--sectores", action="store_true",
                    help="usar las búsquedas por sector definidas en config.py")
    ap.add_argument("--repetir", action="store_true",
                    help="volver a ejecutar búsquedas ya hechas")
    ap.add_argument("-y", "--si", action="store_true", help="no pedir confirmación")
    args = ap.parse_args()

    load_dotenv()
    clave = os.getenv("GOOGLE_PLACES_API_KEY")
    if not clave:
        sys.exit("Falta GOOGLE_PLACES_API_KEY en el archivo .env (ver README).")

    if args.sectores:
        busquedas = [(r, c, sector) for sector, d in SECTORES.items()
                     for r, lugares in d["busquedas"] for c in lugares]
    else:
        busquedas = [(r, c, "") for r in args.rubros for c in args.comunas]
    hechas = set(HECHAS.read_text(encoding="utf-8").splitlines()) if HECHAS.exists() else set()
    if not args.repetir:
        busquedas = [b for b in busquedas if f"{b[0]} | {b[1]}" not in hechas]
    print(f"{len(busquedas)} búsquedas nuevas x hasta {args.paginas} página(s). "
          f"Tope: {args.max_consultas} llamadas a la API.")
    if not args.si and input("¿Continuar? [s/N] ").strip().lower() != "s":
        return

    empresas = cargar_existentes()
    antes, llamadas = len(empresas), 0
    headers = {"X-Goog-Api-Key": clave, "X-Goog-FieldMask": CAMPOS}

    with httpx.Client(headers=headers, timeout=30) as cliente:
        for rubro, comuna, sector in busquedas:
            if llamadas >= args.max_consultas:
                print("Se alcanzó el tope de consultas.")
                break
            paginas = min(args.paginas, args.max_consultas - llamadas)
            try:
                lugares, n = buscar(cliente, f"{rubro} en {comuna}, Chile", paginas)
            except httpx.HTTPStatusError as e:
                if e.response.status_code == 429:
                    print("\nSe alcanzó el límite diario de Google (100 búsquedas/día por defecto). "
                          "Lo avanzado quedó guardado: vuelve a correr el mismo comando mañana "
                          "y seguirá donde quedó.")
                    break
                sys.exit(f"Error de la API ({e.response.status_code}): {e.response.text}")
            llamadas += n
            nuevos = 0
            for p in lugares:
                direccion = p.get("formattedAddress", "")
                categoria = p.get("primaryTypeDisplayName", {}).get("text", "")
                if (p.get("businessStatus", "OPERATIONAL") != "OPERATIONAL"
                        or not any(u in direccion for u in FILTRO_UBICACION) or p["id"] in empresas
                        or excluida(categoria, p.get("displayName", {}).get("text", ""))):
                    continue
                empresas[p["id"]] = {
                    "place_id": p["id"],
                    "nombre": p.get("displayName", {}).get("text", ""),
                    "sector": sector,
                    "rubro_busqueda": rubro,
                    "categoria_maps": categoria,
                    "direccion": direccion,
                    "comuna_busqueda": comuna,
                    "telefono": p.get("nationalPhoneNumber", ""),
                    "web": p.get("websiteUri", ""),
                    "rating": p.get("rating", ""),
                    "n_resenas": p.get("userRatingCount", 0),
                    "maps_url": p.get("googleMapsUri", ""),
                }
                nuevos += 1
            print(f"  {rubro} / {comuna}: {len(lugares)} resultados, {nuevos} nuevos")
            with HECHAS.open("a", encoding="utf-8") as f:
                f.write(f"{rubro} | {comuna}\n")
            guardar(empresas)  # guarda en cada paso: si se corta, no se pierde nada

    print(f"\nListo: {len(empresas) - antes} empresas nuevas, {len(empresas)} en total, "
          f"{llamadas} llamadas a la API → {SALIDA}")


if __name__ == "__main__":
    main()
