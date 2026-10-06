"""Paso 1: buscar empresas en Google Maps usando la Places API (New).

No se hace scraping de Google Maps (lo prohíben sus términos); se usa la API oficial.

Uso:
    python paso1_descubrir.py                     # todos los rubros x comunas (pide confirmación)
    python paso1_descubrir.py --rubros maestranza --comunas Quilpué "Villa Alemana"
    python paso1_descubrir.py --max-consultas 20  # tope de llamadas a la API

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

from config import COMUNAS, FILTRO_REGION, RUBROS, ZONA

URL = "https://places.googleapis.com/v1/places:searchText"
CAMPOS = ",".join([
    "places.id", "places.displayName", "places.formattedAddress",
    "places.websiteUri", "places.nationalPhoneNumber", "places.rating",
    "places.userRatingCount", "places.primaryTypeDisplayName",
    "places.businessStatus", "places.googleMapsUri", "nextPageToken",
])
SALIDA = Path("data/empresas.csv")
COLUMNAS = ["place_id", "nombre", "rubro_busqueda", "categoria_maps", "direccion",
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
    ap.add_argument("-y", "--si", action="store_true", help="no pedir confirmación")
    args = ap.parse_args()

    load_dotenv()
    clave = os.getenv("GOOGLE_PLACES_API_KEY")
    if not clave:
        sys.exit("Falta GOOGLE_PLACES_API_KEY en el archivo .env (ver README).")

    busquedas = [(r, c) for r in args.rubros for c in args.comunas]
    print(f"{len(busquedas)} búsquedas x hasta {args.paginas} página(s). "
          f"Tope: {args.max_consultas} llamadas a la API.")
    if not args.si and input("¿Continuar? [s/N] ").strip().lower() != "s":
        return

    empresas = cargar_existentes()
    antes, llamadas = len(empresas), 0
    headers = {"X-Goog-Api-Key": clave, "X-Goog-FieldMask": CAMPOS}

    with httpx.Client(headers=headers, timeout=30) as cliente:
        for rubro, comuna in busquedas:
            if llamadas >= args.max_consultas:
                print("Se alcanzó el tope de consultas.")
                break
            paginas = min(args.paginas, args.max_consultas - llamadas)
            try:
                lugares, n = buscar(cliente, f"{rubro} en {comuna}, Chile", paginas)
            except httpx.HTTPStatusError as e:
                sys.exit(f"Error de la API ({e.response.status_code}): {e.response.text}")
            llamadas += n
            nuevos = 0
            for p in lugares:
                direccion = p.get("formattedAddress", "")
                if (p.get("businessStatus", "OPERATIONAL") != "OPERATIONAL"
                        or FILTRO_REGION not in direccion or p["id"] in empresas):
                    continue
                empresas[p["id"]] = {
                    "place_id": p["id"],
                    "nombre": p.get("displayName", {}).get("text", ""),
                    "rubro_busqueda": rubro,
                    "categoria_maps": p.get("primaryTypeDisplayName", {}).get("text", ""),
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
            guardar(empresas)  # guarda en cada paso: si se corta, no se pierde nada

    print(f"\nListo: {len(empresas) - antes} empresas nuevas, {len(empresas)} en total, "
          f"{llamadas} llamadas a la API → {SALIDA}")


if __name__ == "__main__":
    main()
