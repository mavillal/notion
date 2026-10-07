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

# CRM – ExpoMina Perú 2026

`CRM_ExpoMina_2026.xlsx`: 979 empresas únicas a partir de los 1.209 stands de la lista oficial de expositores (expominaperu.com, exportada a CSV desde el navegador). Tiene la misma estructura de hojas y añade las columnas Segmento SafetyMind, Stand(s), N° stands, Web, Teléfono(s), Dirección y En Directorio SSO.

- Prioridad: Alta 83 · Media 51 · Baja 845 (según el segmento SafetyMind).
- 74 empresas también figuran en el CRM del Directorio SSO; de ahí se heredan su email, teléfono y dirección.
- Copia viva en Google Drive (carpeta CRM): `CRM_ExpoMina_2026`.

```
python3 scripts/build_expo.py expomina_exhibidores.csv CRM_Directorio_SSO_2026.xlsx CRM_ExpoMina_2026.xlsx
```

# CRM – Exponor 2026 (Antofagasta, Chile)

`CRM_Exponor_2026.xlsx`: 1.375 empresas únicas (1.382 registros de stands del listado de expositores, exportado a CSV). Mismas hojas; columnas extra: Segmento SafetyMind, Stand(s) (pabellón-número), N° stands, Pabellón país / zona (pabellón país, Lanza tu Innovación, AIA-Pyme, Energía), Web, Teléfono(s), Cruce otros CRM.

- Prioridad: Alta 214 · Media 134 · Baja 1.027.
- País: tomado del pabellón país cuando existe; si no, Chile (operación local).
- Cruce: 96 también en ExpoMina 2026 y 21 en el Directorio SSO (contacto heredado de la filial Perú, marcado en Notas).
- Copia viva en Google Drive (carpeta CRM): `CRM_Exponor_2026`.

```
python3 scripts/build_exponor.py exponor_2026.csv CRM_Directorio_SSO_2026.xlsx CRM_ExpoMina_2026.xlsx CRM_Exponor_2026.xlsx
```
