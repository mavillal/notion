"""Paso 3: unir datos de Maps + web, puntuar oportunidades y sugerir el servicio.

Uso:
    python paso3_puntuar.py

Resultado: data/prospectos.csv ordenado de mayor a menor puntaje (abre en Excel).

Ajusta los puntos en PESOS según lo que aprendas de las primeras reuniones.
"""

import csv
import json
from collections import defaultdict
from pathlib import Path

EMPRESAS = Path("data/empresas.csv")
PAGINAS = Path("data/paginas.json")
SALIDA = Path("data/prospectos.csv")

PESOS = {
    # Tamaño / encaje (proxy: reseñas en Maps)
    "resenas_20": 5, "resenas_100": 5,
    # Contexto comercial
    "contratando": 15, "cliente_exigente": 10,
    # SST / cumplimiento
    "sin_ley_karin": 10, "sin_gestion_sst": 10,
    # Mejora continua
    "sin_calidad": 10,
}

LINEA = {  # a qué servicio apunta cada brecha
    "sin_ley_karin": "Cumplimiento SST", "sin_gestion_sst": "Cumplimiento SST",
    "sin_calidad": "Mejora continua",
}

GANCHOS = {  # preguntas de apertura (la web no prueba que no lo tengan: pregunta, no afirmes)
    "sin_ley_karin": "¿Tienen implementado y difundido el protocolo de prevención "
                     "de acoso de la Ley Karin (obligatorio desde ago-2024)?",
    "sin_gestion_sst": "¿Cómo gestionan hoy el cumplimiento del DS 44 en SST "
                       "(matriz de riesgos, programa preventivo)?",
    "sin_calidad": "¿Tienen indicadores de productividad o pérdidas en planta? "
                   "Un diagnóstico lean de 2 semanas suele encontrar ahorros rápidos.",
}

COLUMNAS = ["puntaje", "servicio_sugerido", "gancho", "nombre", "categoria_maps",
            "comuna_busqueda", "direccion", "telefono", "emails", "web", "n_resenas",
            "rating", "brechas", "senales", "paginas_visitadas", "maps_url"]


def agrupar_paginas() -> dict[str, list[dict]]:
    por_empresa: dict[str, list[dict]] = defaultdict(list)
    if PAGINAS.exists():
        for p in json.loads(PAGINAS.read_text(encoding="utf-8")):
            por_empresa[p["place_id"]].append(p)
    return por_empresa


def evaluar(empresa: dict, paginas: list[dict]) -> dict:
    ok = [p for p in paginas if p.get("tipo") != "error"]
    inicio = next((p for p in ok if p["tipo"] == "inicio"), None)
    senales = sorted({s for p in ok for s in p["senales"]})
    emails = sorted({e for p in ok for e in p["emails"]})
    tiene = set(senales).__contains__

    puntos: dict[str, int] = {}
    n = int(empresa.get("n_resenas") or 0)
    if n >= 20:
        puntos["resenas_20"] = PESOS["resenas_20"]
    if n >= 100:
        puntos["resenas_100"] = PESOS["resenas_100"]
    if tiene("crecimiento_empleo"):
        puntos["contratando"] = PESOS["contratando"]
    if tiene("cliente_mineria") or tiene("cliente_portuario"):
        puntos["cliente_exigente"] = PESOS["cliente_exigente"]

    brechas = []
    if inicio:  # sin web visitada no hay evidencia de brechas
        if not tiene("sst_ley_karin"):
            brechas.append("sin_ley_karin")
        if not (tiene("sst_iso45001") or tiene("sst_prevencion") or tiene("sst_comite_paritario")):
            brechas.append("sin_gestion_sst")
        if not (tiene("calidad_iso9001") or tiene("mejora_continua")):
            brechas.append("sin_calidad")
    for b in brechas:
        puntos[b] = PESOS[b]

    por_linea: dict[str, int] = defaultdict(int)
    for b in brechas:
        por_linea[LINEA[b]] += PESOS[b]
    servicio = max(por_linea, key=por_linea.get) if por_linea else ""
    gancho = next((GANCHOS[b] for b in sorted(brechas, key=lambda b: -PESOS[b])
                   if LINEA[b] == servicio and b in GANCHOS), "")

    return {
        "puntaje": sum(puntos.values()),
        "servicio_sugerido": servicio,
        "gancho": gancho,
        "emails": ", ".join(emails),
        "brechas": ", ".join(brechas),
        "senales": ", ".join(senales),
        "paginas_visitadas": len(ok),
    }


def main() -> None:
    with EMPRESAS.open(encoding="utf-8-sig") as f:
        empresas = list(csv.DictReader(f, delimiter=";"))
    paginas = agrupar_paginas()

    filas = [{**e, **evaluar(e, paginas.get(e["place_id"], []))} for e in empresas]
    filas.sort(key=lambda f: f["puntaje"], reverse=True)

    with SALIDA.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=COLUMNAS, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(filas)

    print(f"Listo: {len(filas)} prospectos → {SALIDA}")
    for f in filas[:10]:
        print(f"  {f['puntaje']:>3}  {f['servicio_sugerido']:<17} {f['nombre']}")


if __name__ == "__main__":
    main()
