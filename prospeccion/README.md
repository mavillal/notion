# Prospección industrial – V Región (Chile)

Encuentra empresas industriales medianas de la V Región en Google Maps, analiza su web
con [Crawlee](https://crawlee.dev/python/) y genera un Excel de prospectos con puntaje,
servicio sugerido (Digitalización / Cumplimiento SST / Mejora continua) y una pregunta
de apertura para el primer contacto.

```
paso1_descubrir.py  →  data/empresas.csv     (Google Maps vía Places API)
paso2_enriquecer.py →  data/paginas.json     (Crawlee visita cada web)
paso3_puntuar.py    →  data/prospectos.csv   (puntaje + servicio + gancho)  ← abrir en Excel
```

## 1. Instalación (una vez)

```bash
cd prospeccion
python3 -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
cp .env.example .env             # luego pega tu clave en .env
```

### Clave de Google Places API
1. Entra a <https://console.cloud.google.com/> y crea un proyecto.
2. Activa la facturación (exige una tarjeta, aunque uses el tramo gratuito).
3. En **APIs y servicios → Biblioteca**, activa **Places API (New)**.
4. En **Credenciales → Crear credencial → Clave de API**, crea la clave y restríngela a *Places API (New)*.
5. Pégala en `.env`: `GOOGLE_PLACES_API_KEY=...`
6. **Recomendado:** en **Facturación → Presupuestos**, crea una alerta (por ejemplo, de USD 20).

**Costo:** cada llamada trae hasta 20 empresas. Como pedimos web y teléfono, se cobra la tarifa
"Text Search Enterprise", hoy alrededor de USD 35 por cada 1.000 llamadas, con un cupo gratis mensual.
Revisa los precios vigentes en la consola. El script tiene un tope (`--max-consultas`, 100 por defecto).

## 2. Uso

```bash
# Prueba chica: 1 rubro en 2 comunas
python paso1_descubrir.py --rubros maestranza --comunas Quilpué "Villa Alemana"

# Barrido completo (14 rubros x 21 comunas ≈ 294 llamadas)
python paso1_descubrir.py --max-consultas 300

python paso2_enriquecer.py --limite 20   # prueba con 20 webs
python paso2_enriquecer.py               # todas
python paso3_puntuar.py
```

Abre `data/prospectos.csv` en Excel. Las empresas vienen ordenadas por puntaje.

**Sin clave de Google** también puedes probar: crea `data/empresas.csv` a mano con las columnas
`place_id;nombre;web` (más las otras columnas vacías) y corre los pasos 2 y 3.

## 3. Cómo se puntúa

| Bloque | Qué detecta | Puntos |
|---|---|---|
| Tamaño | 20+ y 100+ reseñas en Maps (indicio de tamaño) | 5 + 5 |
| Momento | "Trabaja con nosotros" / está contratando | 15 |
| Momento | Atiende minería o puertos (clientes exigentes en SST) | 10 |
| Digitalización | Sin web propia / sin https / no responsive / web de hace 3+ años / sin formulario / sin ERP ni portal | 25 / 10 / 10 / 10 / 5 / 5 |
| Cumplimiento SST | La web no menciona Ley Karin / no menciona gestión SST (ISO 45001, prevención, comité paritario) | 10 / 10 |
| Mejora continua | No menciona ISO 9001, lean, 5S ni kaizen | 10 |

El servicio sugerido es la línea que acumula más puntos de brecha. Edita `PESOS` y `GANCHOS` en
`paso3_puntuar.py`, y los rubros, comunas y señales en `config.py`.

> ⚠️ Que algo no aparezca en la web **no prueba** que la empresa no lo tenga. Por eso los ganchos
> son preguntas, no afirmaciones.

## 4. Reglas legales y de buena práctica

- **Google Maps:** se usa solo la API oficial. No hagas scraping de maps.google.com (lo prohíben sus términos).
- **Webs:** el crawler respeta `robots.txt`, va lento (máximo 90 páginas por minuto, 5 en paralelo) y no entra a zonas con login.
- **Ley 21.719 de datos personales** (vigente desde el 1 de diciembre de 2026): el sistema guarda solo
  datos de la empresa (teléfono comercial, info@, ventas@…). **Descarta los emails de personas.**
  Si después agregas contactos con nombre (por ejemplo, el gerente de planta), registra de dónde
  salió el dato y la base legal, y atiende las solicitudes de eliminación.
- **Correos comerciales:** identifícate claramente, entrega una forma simple de darse de baja y
  respétala. Usa un dominio secundario para los envíos masivos y parte con pocos correos al día.

## 5. Próximas mejoras sugeridas

1. **Filtro de tamaño real:** cruzar por RUT con la nómina de empresas del SII (datos abiertos), que trae el tramo de ventas y el número de trabajadores.
2. **Gancho con IA:** pasar el texto de la web a Claude para redactar un primer mensaje personalizado por empresa.
3. **CRM en Notion:** subir `prospectos.csv` a una base de Notion con estados (contactado / reunión / propuesta).
4. **Webs con JavaScript:** usar `PlaywrightCrawler` en lugar de `BeautifulSoupCrawler` para los sitios que lo requieran.
5. **Re-crawl mensual:** alertar cuando una empresa publique ofertas de empleo o cambie su web.

## Tests

```bash
pytest -q
```
