# CRM – Directorio SSO 2026 (Cero Accidentes)

CRM generado a partir del *Directorio de Seguridad y Salud Ocupacional 2026* (22.ª ed., ceroaccidentes.pe), con la misma estructura que `CRM_Eventos_SST_HSE_SSO_Peru_Colombia.xlsx` (hojas Contactos, Oportunidades, Resumen y Eventos).

- `CRM_Directorio_SSO_2026.xlsx`: 674 empresas deduplicadas (1.093 fichas en el PDF).
- Copia viva en Google Drive (carpeta CRM): `CRM_Directorio_SSO_2026_CeroAccidentes`.

## Columnas de Contactos
Las 12 columnas del Excel original (ID … Notas), más: Rubro(s), Segmento SafetyMind, Teléfono(s), Web, Dirección, Ciudad/Región y Anunciante.

## Regenerar
```
python3 scripts/parse.py Directorio.pdf raw.json
python3 scripts/build.py raw.json CRM_Directorio_SSO_2026.xlsx scripts/fixes.json
```
