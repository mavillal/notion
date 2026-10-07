# Base de datos — Comunidad LegadoSostenible (WhatsApp, 20 sep – 6 oct 2026)

`comunidad.db` (SQLite) y `csv/` (un CSV por tabla, listo para **Notion → Import → CSV**).
Teléfonos enmascarados (`+51 ···869`: código de país + últimos 3 dígitos).

| Tabla | Filas | Qué contiene |
|---|---|---|
| `mensajes` | 166 | fecha, hora, autor, tipo (Texto/Imagen/Sticker/Enlace/Documento/Eliminado), texto, url, plataforma, es_pregunta |
| `miembros` | 65 | rol, país, fecha y vía de ingreso, nº de mensajes, primer/último mensaje, estado |
| `eventos` | 42 | ingresos por enlace, añadidos, eliminaciones |
| `herramientas` | 12 | herramientas de IA recomendadas: categoría, caso de uso, quién la recomendó, costo |
| `prompts` | 2 | prompts completos compartidos (mapa de riesgos, infografía a mano) |
| `consultas` | 10 | preguntas de la comunidad, respuesta y estado (Resuelta/Parcial/Sin respuesta) |
| `anuncios` | 8 | cursos, eventos, normas del grupo, empleo, investigación |
| `contenidos` | 14 | enlaces compartidos (TikTok, LinkedIn, Instagram…) |

Regenerar: `python3 generar_bd.py <export_whatsapp.txt>`
