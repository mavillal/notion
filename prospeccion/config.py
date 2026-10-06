"""Configuración editable: qué buscar, dónde y qué señales detectar.

Modifica estas listas para ajustar la prospección. No necesitas tocar los
otros archivos para cambiar rubros, comunas o señales.
"""

# --- Paso 1: búsqueda en Google Maps (Places API) ---------------------------

COMUNAS = [
    "Valparaíso", "Viña del Mar", "Concón", "Quilpué", "Villa Alemana",
    "Quintero", "Puchuncaví", "Casablanca", "San Antonio", "Cartagena",
    "Quillota", "La Calera", "La Cruz", "Nogales", "Hijuelas", "Limache",
    "Los Andes", "San Felipe", "Llay-Llay", "Catemu", "Panquehue",
]

RUBROS = [
    "maestranza",
    "empresa metalmecánica",
    "planta procesadora de alimentos",
    "packing frutícola",
    "agroindustria",
    "fábrica",
    "planta industrial",
    "fundición",
    "empresa de ingeniería y montaje industrial",
    "servicios portuarios",
    "centro de distribución logística",
    "transporte de carga",
    "planta química",
    "empresa de mantención industrial",
]

# Búsquedas por sector (se usan con: python paso1_descubrir.py --sectores).
# Cada búsqueda es (rubro, [lugares]): se busca el rubro en cada lugar. La "oportunidad" se copia al Excel final.
SECTORES = {
    "Logística": {
        "oportunidad": "Gestión de turnos, digitalización de hojas de ruta, SST por fatiga "
                       "de conductores y riesgos en bodegaje.",
        "busquedas": [
            ("bodegas y almacenaje", ["Valparaíso", "Placilla, Valparaíso", "Curacaví"]),
            ("centro de distribución", ["Curauma, Valparaíso", "Curacaví"]),
            ("servicios logísticos", ["San Antonio", "Valparaíso", "Los Andes"]),
            ("operador logístico", ["Valparaíso", "San Antonio", "Los Andes"]),
            ("almacén extraportuario", ["Valparaíso", "San Antonio", "Placilla, Valparaíso"]),
            ("depósito de contenedores", ["San Antonio", "Valparaíso"]),
            ("transporte de carga", ["San Antonio", "Concón", "Los Andes", "Quillota",
                                     "San Felipe", "Curacaví", "Casablanca"]),
            ("transporte internacional de carga", ["Los Andes"]),
        ],
    },
    "Manufactura": {
        "oportunidad": "Automatización de reportes de planta, digitalización del layout, "
                       "auditoría de matrices de riesgo e implementación DS 44 / Ley Karin.",
        "busquedas": [
            ("maquinaria industrial", ["Viña del Mar"]),
            ("alimentos procesos", ["Quilpué", "Villa Alemana"]),
            ("maestranza", ["Concón", "Valparaíso", "Viña del Mar", "San Antonio",
                            "Quillota", "La Calera", "Los Andes", "San Felipe", "Curacaví"]),
            ("montajes industriales", ["Región de Valparaíso"]),
            ("empresa metalmecánica", ["Concón", "Quilpué", "Viña del Mar", "Valparaíso",
                                       "Quillota", "Los Andes", "San Antonio"]),
            ("estructuras metálicas", ["Quilpué", "Villa Alemana", "Quillota", "Curacaví",
                                       "San Felipe"]),
            ("planta industrial", ["Concón", "Quintero", "Puchuncaví", "La Calera", "Curacaví"]),
            ("industria química", ["Concón", "Quintero"]),
            ("fábrica de envases", ["Región de Valparaíso"]),
            ("industria plástica", ["Región de Valparaíso"]),
            ("planta de hormigón", ["La Calera", "Quillota", "Viña del Mar", "San Antonio"]),
            ("fábrica de alimentos", ["Viña del Mar", "Valparaíso", "Quillota", "Curacaví",
                                      "San Antonio"]),
            ("planta procesadora de alimentos", ["Curacaví", "Casablanca", "Quillota"]),
        ],
    },
    "Agroindustria": {
        "oportunidad": "Control de contratistas, cumplimiento sanitario/SST en temporada "
                       "alta, digitalización de check-lists de calidad.",
        "busquedas": [
            ("frigoríficos", ["Quillota"]),
            ("empacadoras", ["San Felipe", "Los Andes"]),
            ("viñas", ["Casablanca"]),
            ("packing frutícola", ["Quillota", "La Cruz", "Hijuelas", "Nogales", "San Felipe",
                                   "Los Andes", "Llay-Llay", "Curacaví", "Casablanca"]),
            ("exportadora de fruta", ["San Felipe", "Los Andes", "Quillota", "Curacaví"]),
            ("planta de frío", ["Quillota", "San Felipe", "Los Andes", "Curacaví"]),
            ("procesadora de nueces", ["San Felipe", "Los Andes", "Curacaví"]),
            ("planta deshidratadora", ["Región de Valparaíso", "Curacaví"]),
            ("agroindustria", ["Curacaví", "Quillota", "San Felipe", "Los Andes"]),
            ("planta avícola", ["Curacaví", "Región de Valparaíso"]),
            ("viña", ["Valle de Leyda, San Antonio", "Curacaví", "Valle de Aconcagua"]),
            ("vivero frutal", ["Quillota", "Curacaví"]),
        ],
    },
    "Servicios industriales": {
        "oportunidad": "Contratistas de minería, puertos y plantas: exigencias SST de "
                       "mandantes (DS 44, Ley Karin, ISO 45001) y control de documentación.",
        "busquedas": [
            ("contratista minero", ["Los Andes", "San Felipe"]),
            ("servicios a la minería", ["Los Andes", "Viña del Mar"]),
            ("mantención industrial", ["Concón", "Quintero", "San Antonio", "Los Andes",
                                       "Quillota"]),
            ("servicios portuarios", ["San Antonio", "Valparaíso"]),
            ("aseo industrial", ["Viña del Mar", "San Antonio", "Quillota"]),
            ("montaje eléctrico industrial", ["Viña del Mar", "Quilpué", "Los Andes"]),
            ("arriendo de maquinaria pesada", ["Quillota", "Los Andes", "San Antonio", "Curacaví"]),
            ("empresa constructora industrial", ["Viña del Mar", "Concón"]),
        ],
    },
}

# Categorías de Google Maps que se descartan (no son el cliente objetivo).
# Se compara en minúsculas y basta con que el texto esté contenido.
CATEGORIAS_EXCLUIDAS = [
    "mudanza", "tienda", "supermercado", "restaurante", "cafetería", "panadería",
    "carnicería", "hotel", "camping", "centro comercial", "gran superficie",
    "ferretería", "atracción turística", "parque", "jardín", "jardiner", "condominio",
    "institución educativa", "recinto para eventos", "catering", "agencia de colocación",
    "bancos y finanzas", "oficina de la administración", "oficina de gobierno",
    "taller de automóviles", "concesionario", "taxi", "laboratorio de análisis", "hospital",
    "mercado", "parada de autobús", "estación de carga", "farmacia", "agencia de viajes",
    "gasolinera", "aparcamiento", "coworking", "mirador", "contabilidad", "gimnasio",
    "iglesia", "asociación u organización",
]

# Nombres que se descartan aunque la categoría pase el filtro.
NOMBRES_EXCLUIDOS = ["aduana", "servicio nacional", "municipalidad"]

# Solo se guardan resultados cuya dirección contenga alguno de estos textos.
# (Curacaví es de la Región Metropolitana, pero es parte de la zona objetivo.)
FILTRO_UBICACION = ["Valparaíso", "Curacaví"]

# Rectángulo aproximado de la V Región continental (limita la búsqueda).
ZONA = {
    "low": {"latitude": -34.0, "longitude": -72.0},
    "high": {"latitude": -32.0, "longitude": -70.0},
}

# --- Paso 2: crawler de sitios web -----------------------------------------

# Subpáginas que vale la pena visitar (se busca el texto en la URL o el link).
PALABRAS_SUBPAGINAS = [
    "contact", "nosotros", "empresa", "quienes", "servicio", "calidad",
    "seguridad", "sst", "trabaja", "empleo", "certific", "sostenib", "politica",
]
MAX_SUBPAGINAS = 6

# Dominios que no son una web propia (se tratan como "sin web").
DOMINIOS_NO_PROPIOS = [
    "facebook.com", "instagram.com", "linkedin.com", "wa.me", "whatsapp.com",
    "google.com", "business.site", "linktr.ee",
]

# Solo se guardan emails corporativos genéricos (Ley 21.719: evitar datos
# personales sin base legal). Los emails personales se descartan.
PREFIJOS_EMAIL_GENERICOS = {
    "info", "contacto", "contact", "ventas", "comercial", "administracion",
    "gerencia", "rrhh", "recepcion", "oficina", "cotizaciones", "hola",
}

# Señales: nombre -> lista de expresiones regulares (en minúsculas).
SENALES = {
    # SST / cumplimiento
    "sst_ley_karin": [r"ley karin", r"21\.643", r"acoso laboral"],
    "sst_iso45001": [r"iso\s*45001", r"ohsas\s*18001"],
    "sst_comite_paritario": [r"comit[eé] paritario"],
    "sst_prevencion": [r"prevenci[oó]n de riesgos", r"seguridad y salud"],
    # Calidad / mejora continua
    "calidad_iso9001": [r"iso\s*9001"],
    "ambiente_iso14001": [r"iso\s*14001"],
    "mejora_continua": [r"\blean\b", r"mejora continua", r"\b5s\b", r"kaizen",
                        r"six sigma"],
    # Contexto comercial
    "crecimiento_empleo": [r"trabaja con nosotros", r"[uú]nete a nuestro equipo",
                           r"ofertas? (de )?(empleo|trabajo)",
                           r"estamos contratando", r"postula"],
    "cliente_mineria": [r"miner[ií]a", r"codelco", r"anglo american"],
    "cliente_portuario": [r"portuari", r"terminal pacífico", r"\btps\b"],
}
