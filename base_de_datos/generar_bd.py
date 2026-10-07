"""Genera la base de datos de la comunidad LegadoSostenible a partir del export de WhatsApp.

Uso:  python3 generar_bd.py <chat.txt> [--completo <carpeta_salida>]
Salida: comunidad.db (SQLite) + un CSV por tabla en ./csv (importables en Notion).

Los teléfonos se enmascaran (código de país + últimos 3 dígitos) para no publicar datos personales.
--completo conserva los números completos; úsalo con una carpeta fuera del repo.
"""
import csv
import re
import sqlite3
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

OUT = Path(__file__).parent
COMPLETO = False
ADMIN = "@jjosecornejo"

PAISES = {
    "+51": "Perú", "+52": "México", "+56": "Chile", "+57": "Colombia", "+58": "Venezuela",
    "+39": "Italia", "+503": "El Salvador", "+504": "Honduras", "+507": "Panamá",
    "+591": "Bolivia", "+593": "Ecuador",
}

LINEA = re.compile(r"^\[(\d{1,2}/\d{1,2}/\d{2}), (\d{1,2}:\d{2}:\d{2} [AP]M)\] (.*)$", re.S)
TEL = re.compile(r"\+\d[\d \-]{6,}\d")
URL = re.compile(r"https?://\S+")


def mask(tel):
    if COMPLETO:
        return tel
    digits = re.sub(r"\D", "", tel)
    code = next((c for c in sorted(PAISES, key=len, reverse=True) if digits.startswith(c[1:])), "+" + digits[:2])
    return f"{code} ···{digits[-3:]}"


def pais(autor):
    if not autor.startswith("+"):
        return None
    digits = re.sub(r"\D", "", autor)
    return next((v for c, v in sorted(PAISES.items(), key=lambda x: -len(x[0])) if digits.startswith(c[1:])), None)


def anon(texto):
    return TEL.sub(lambda m: mask(m.group()), texto)


def plataforma(url):
    for k, v in [("tiktok", "TikTok"), ("linkedin", "LinkedIn"), ("instagram", "Instagram"),
                 ("gemini", "Gemini"), ("perusostenible", "Perú Sostenible")]:
        if k in url:
            return v
    return "Web"


def parse(path):
    registros = []
    for raw in Path(path).read_text(encoding="utf-8").splitlines():
        m = LINEA.match(raw)
        if m:
            registros.append(list(m.groups()))
        elif registros:
            registros[-1][2] += "\n" + raw
    mensajes, eventos = [], []
    for fecha, hora, cuerpo in registros:
        ts = datetime.strptime(f"{fecha} {hora}", "%m/%d/%y %I:%M:%S %p")
        if cuerpo.startswith("- "):
            eventos.append((ts, cuerpo[2:].strip()))
            continue
        autor, _, texto = cuerpo.partition(": ")
        mensajes.append((ts, autor.strip(), texto.strip()))
    return mensajes, eventos


def tipo_mensaje(texto):
    if "<imagen omitida>" in texto:
        return "Imagen"
    if "<sticker omitido>" in texto:
        return "Sticker"
    if "<documento omitido>" in texto:
        return "Documento"
    if "eliminó este mensaje" in texto or texto.startswith("Se eliminó"):
        return "Eliminado"
    if URL.search(texto):
        return "Enlace"
    return "Texto"


def main(src):
    mensajes, eventos = parse(src)

    # ---------- mensajes ----------
    filas_msg = []
    for i, (ts, autor, texto) in enumerate(mensajes, 1):
        t = tipo_mensaje(texto)
        limpio = re.sub(r"<(imagen|sticker|documento) omitid[oa]>", "", texto).strip()
        url = (URL.search(texto) or [None])[0] if URL.search(texto) else None
        filas_msg.append({
            "id": i, "fecha": ts.date().isoformat(), "hora": ts.strftime("%H:%M"),
            "autor": anon(autor), "es_admin": int(autor == ADMIN), "tipo": t,
            "texto": anon(limpio), "url": url, "plataforma": plataforma(url) if url else None,
            "es_pregunta": int("?" in limpio or bool(re.match(r"(?i)(alguna|alguien|buenas|buena noche)", limpio))),
        })

    # ---------- eventos (ingresos, altas, bajas) ----------
    filas_ev = []
    for i, (ts, txt) in enumerate(eventos, 1):
        if "se unieron" in txt:
            tipo, quien = "Ingreso por enlace", txt.split(" se unieron")[0]
        elif "añadió a" in txt:
            tipo, quien = "Añadido por miembro", txt.split("añadió a ")[1]
        elif "eliminó a" in txt:
            tipo, quien = "Eliminado del grupo", txt.split("eliminó a ")[1]
        else:
            tipo, quien = "Sistema", None
        filas_ev.append({"id": i, "fecha": ts.date().isoformat(), "hora": ts.strftime("%H:%M"),
                         "tipo": tipo, "miembro": anon(quien) if quien else None, "detalle": anon(txt)})

    # ---------- miembros ----------
    conteo = Counter(m["autor"] for m in filas_msg if m["autor"] != "Tú")
    ingreso = {e["miembro"]: (e["fecha"], e["tipo"]) for e in filas_ev if e["miembro"] and e["tipo"] != "Eliminado del grupo"}
    nombres = set(conteo) | set(ingreso)
    nombres.discard("Tú")
    filas_mb = []
    for n in sorted(nombres, key=lambda x: (-conteo.get(x, 0), x)):
        fechas = [m["fecha"] for m in filas_msg if m["autor"] == n]
        filas_mb.append({
            "miembro": n, "rol": "Administrador" if n == ADMIN else "Miembro",
            "pais": pais(n), "fecha_ingreso": ingreso.get(n, (None,))[0],
            "via_ingreso": ingreso.get(n, (None, "Anterior al export"))[1],
            "mensajes": conteo.get(n, 0),
            "primer_mensaje": min(fechas) if fechas else None,
            "ultimo_mensaje": max(fechas) if fechas else None,
            "estado": "Activo" if conteo.get(n, 0) else "Silencioso",
        })

    # ---------- tablas curadas (conocimiento útil del chat) ----------
    herramientas = [
        ("Adobe Podcast", "Audio", "Mejorar el audio de capacitaciones y videos", "+51 ···869; @jjosecornejo", "2026-09-23", "Gratis/Freemium"),
        ("Riverside Transcription", "Transcripción", "Transcribir video/audio para usarlo como conocimiento de una IA", "@jjosecornejo", "2026-09-24", "Freemium"),
        ("NotebookLM (Gemini)", "Presentaciones / estudio", "Generar diapositivas y resúmenes a partir de un PDF", "+51 ···146", "2026-09-24", "Gratis"),
        ("Genially", "Presentaciones", "Presentaciones interactivas", "@jjosecornejo", "2026-09-24", "Freemium"),
        ("Floorplanner", "Planos / layout", "Plano arquitectónico para mapa de riesgos (señalización manual)", "+51 ···571", "2026-09-24", "Freemium"),
        ("Planner 5D", "Planos / layout", "Plano arquitectónico para mapa de riesgos", "+51 ···571", "2026-09-24", "Freemium"),
        ("Canva", "Diseño", "Añadir señalización y acabado gráfico al plano", "+51 ···571", "2026-09-24", "Freemium"),
        ("AutoCAD", "Planos / layout", "Alternativa técnica para el plano del mapa de riesgos", "@RonaldAngo", "2026-09-24", "Pago"),
        ("SideMe", "Aprender a programar", "Lógica de programación básica", "+51 ···571", "2026-09-24", "Por confirmar"),
        ("Gemini", "IA generativa (imagen/video)", "Mapa de riesgos con prompt; infografías; video ~10 s (Pro)", "+51 ···445; +51 ···217; @fredysavath", "2026-09-24", "Freemium"),
        ("ChatGPT", "IA generativa", "Iterar ideas de layout pidiéndole mejoras", "Americo", "2026-09-25", "Freemium"),
        ("Kimi K3", "Presentaciones (PPT)", "Generar presentaciones; reportan anuncio/bloqueo en versión free", "+51 ···445", "2026-09-26", "Por confirmar"),
    ]
    prompts = [
        ("Mapa de riesgos industrial (vista superior)", "Gemini", "+51 ···445", "2026-09-24", "Mapa de riesgos",
         next(m["texto"] for m in filas_msg if m["texto"].startswith("Diseña un mapa de riesgos"))),
        ("Infografía dibujada a mano en cuaderno", "Gemini / generador de imágenes", "@jjosecornejo", "2026-10-04", "Infografía",
         next(m["texto"] for m in filas_msg if m["texto"].startswith("Aquí va: «Genera"))),
    ]
    consultas = [
        ("2026-09-24", "+51 ···871", "Plataforma interactiva que genere diapositivas desde un PDF", "NotebookLM; Genially (si debe ser interactiva)", "Resuelta"),
        ("2026-09-24", "+51 ···775", "IA para generar layout de mapa de riesgos (plano arquitectónico con señalización)", "Floorplanner/Planner 5D + Canva; AutoCAD; prompt en Gemini; ChatGPT", "Resuelta"),
        ("2026-09-24", "+52 ···093", "App para aprender a programar", "SideMe", "Resuelta"),
        ("2026-09-25", "+593 ···950", "¿Cómo ubicar riesgos psicosociales en un mapa de riesgos?", None, "Sin respuesta"),
        ("2026-09-26", "@LuchitoVillalobos", "App de IA para hacer PPT y prompt sugerido", "Kimi K3 (prompt según tema)", "Parcial"),
        ("2026-09-26", "@BryanMR18", "¿Kimi K3 es gratis? Aparece anuncio y no deja usarlo", None, "Sin respuesta"),
        ("2026-09-28", "+51 ···692", "Diplomados gratuitos en SST", "No se ofrecen diplomados; a veces talleres gratuitos", "Resuelta"),
        ("2026-09-28", "+51 ···436", "Expertos titulados con experiencia en IA para validar prototipo de tesis (juicio de experto)", "Admin lo contacta por privado", "En curso"),
        ("2026-10-03", "+58 ···721", "¿Cómo empezar en tecnología aplicada a la seguridad laboral sin experiencia?", "Admin respondió con imagen (contenido no exportado)", "Parcial"),
        ("2026-10-04", "+57 ···197", "Prompt para infografías", "Prompt de infografía a mano + video tutorial", "Resuelta"),
    ]
    anuncios = [
        ("2026-09-21", "Idea", "Asistentes IA para tiempos de espera en SST (ingreso de contratistas, EMOs)", "@jjosecornejo"),
        ("2026-09-23", "Evento", "Cumbre Perú Sostenible (Magdalena, Lima)", "@jjosecornejo"),
        ("2026-09-26", "Curso", "Curso intensivo Creación de Apps para la SST con IA (mié, vie y sáb presencial)", "@jjosecornejo"),
        ("2026-09-28", "Investigación", "Solicitud de juicio de experto para tesis sobre app de registros SST", "+51 ···436"),
        ("2026-09-30", "Curso", "Inicio del curso 'Creación de Apps para la SST con IA' y nueva edición del workshop 'IA generativa para tomadores de decisiones'", "@jjosecornejo"),
        ("2026-10-02", "Norma del grupo", "Prohibido promover migración a otras comunidades y publicar cursos similares sin autorización del admin", "@jjosecornejo"),
        ("2026-10-03", "Empleo", "Manpower: Coordinador y Supervisor SSOMA, construcción, Callao (+3 años, certificaciones)", "+51 ···973"),
        ("2026-10-03", "Curso", "Sesión presencial del curso con Stracon, Grupo Efe, DJI (QTC), EO RS", "@jjosecornejo"),
    ]

    tablas = {
        "mensajes": filas_msg,
        "eventos": filas_ev,
        "miembros": filas_mb,
        "herramientas": [dict(zip(["herramienta", "categoria", "caso_de_uso", "recomendada_por", "fecha", "costo"], h)) for h in herramientas],
        "prompts": [dict(zip(["titulo", "herramienta", "autor", "fecha", "categoria", "prompt"], p)) for p in prompts],
        "consultas": [dict(zip(["fecha", "autor", "pregunta", "respuesta", "estado"], c)) for c in consultas],
        "anuncios": [dict(zip(["fecha", "tipo", "descripcion", "autor"], a)) for a in anuncios],
        "contenidos": [{"fecha": m["fecha"], "autor": m["autor"], "plataforma": m["plataforma"], "url": m["url"]}
                       for m in filas_msg if m["url"]],
    }

    if COMPLETO:  # las tablas curadas usan alias enmascarados: se restauran al número completo
        texto = Path(src).read_text(encoding="utf-8")
        completos = {}
        for t in TEL.findall(texto):
            globals()["COMPLETO"] = False
            completos[mask(t)] = t
            globals()["COMPLETO"] = True
        patron = re.compile("|".join(re.escape(k) for k in completos))
        for filas in tablas.values():
            for f in filas:
                for k, v in f.items():
                    if isinstance(v, str):
                        f[k] = patron.sub(lambda m: completos[m.group()], v)

    db_path = OUT / "comunidad.db"
    db_path.unlink(missing_ok=True)
    con = sqlite3.connect(db_path)
    (OUT / "csv").mkdir(exist_ok=True)
    for nombre, filas in tablas.items():
        cols = list(filas[0].keys())
        con.execute(f"CREATE TABLE {nombre} ({', '.join(cols)})")
        con.executemany(f"INSERT INTO {nombre} VALUES ({', '.join('?' * len(cols))})", [tuple(f.values()) for f in filas])
        with open(OUT / "csv" / f"{nombre}.csv", "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=cols)
            w.writeheader()
            w.writerows(filas)
        print(f"{nombre:13} {len(filas):4} filas")
    con.commit()
    con.close()


if __name__ == "__main__":
    if "--completo" in sys.argv:
        COMPLETO = True
        OUT = Path(sys.argv[sys.argv.index("--completo") + 1])
        OUT.mkdir(parents=True, exist_ok=True)
    main(sys.argv[1])
