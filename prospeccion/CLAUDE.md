# Contexto para Claude Code

Proyecto de prospección B2B. El usuario sabe Python básico: explica los cambios en
español simple, mantén el código plano (scripts numerados, sin frameworks extra) y los
comentarios en español.

- Objetivo: empresas industriales medianas de la V Región de Chile, con presencia en Google Maps.
- Servicios a ofrecer: cumplimiento legal SST (Ley 16.744, DS 44, Ley Karin 21.643) y mejora
  continua. La digitalización está pausada a pedido del usuario: no agregarla sin que lo pida.
- Pipeline: `paso1_descubrir.py` (Places API New) → `paso2_enriquecer.py` (Crawlee
  BeautifulSoupCrawler) → `paso3_puntuar.py` (puntaje). La configuración está en `config.py`.
  La lógica pura de análisis HTML está en `senales.py` (testeada en `test_senales.py`).
- Los CSV usan `;` y `utf-8-sig` (Excel en español).
- Reglas fijas: no hacer scraping de Google Maps ni de LinkedIn (solo APIs oficiales); respetar
  robots.txt; guardar solo emails corporativos genéricos (Ley 21.719); nunca commitear `.env`
  ni `data/`.
- Antes de commitear, correr `pytest -q`.
- Resultados en Drive: hoja "Prospectos V Región + Curacaví" (carpeta Scraper). Actualizar
  siempre ese mismo archivo, nunca crear otro. El usuario la edita a mano: leerla completa antes
  de escribir y agregar solo filas nuevas (deduplicar por el cid de `maps_url`), respetando sus
  columnas, notas y filas borradas.
