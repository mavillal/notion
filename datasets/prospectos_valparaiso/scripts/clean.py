import pandas as pd, glob, sys, os, re, unicodedata
from urllib.parse import urlparse, parse_qs
RAW, OUT = sys.argv[1], sys.argv[2]
log=[]
def L(step, n, detail): log.append({'Paso':step,'Filas afectadas':n,'Detalle':detail})

# 1. Carga
frames=[]
files=sorted(glob.glob(RAW+'/*.csv'))
for f in files:
    df=pd.read_csv(f,dtype=str); df['_archivo']=os.path.basename(f); frames.append(df)
a=pd.concat(frames,ignore_index=True)
L('Carga',len(a),f'{len(files)} CSV (se ignoran 22 archivos basura __MACOSX/._*)')
foot=a['Name'].str.contains('Map Lead Scraper|exported from',na=False)
a=a[~foot].copy(); L('Eliminar pie de página del scraper',int(foot.sum()),'Filas "Map Lead Scraper exports only 15 records..." / "This data was exported from..."')
a=a.apply(lambda s: s.str.strip() if s.dtype==object else s)

# 2. Búsqueda de origen
FILEQ={'Logística':'Logística','Bodegas,':'Bodegas','Transporte':'Transporte','«Servicios':'Servicios logísticos','Servicios-':'Servicios logísticos',
 'Agricola':'Agrícola','Viña_':'Viña/Casablanca','Frigorífic':'Frigoríficos','Maquinaria':'Maquinaria','Procesos-Q':'Procesos químicos','Empresa,':'Empresa','Manufactur':'Manufactura','Calca':'Calca'}
def query(r):
    u=r['Review URL']
    if isinstance(u,str):
        q=parse_qs(urlparse(u).query).get('q',[''])[0]
        q=re.sub(r'\s+',' ',q.replace('«','').replace('»','')).strip(' -')
        if q: return q
    for k,v in FILEQ.items():
        if k in r['_archivo']: return v+' (archivo)'
    return r['_archivo']
a['_q']=a.apply(query,axis=1)

# 3. Deduplicación por Place Id
agg=a.groupby('Place Id').agg(_busq=('_q',lambda s:'; '.join(sorted(set(s)))),_n=('_q','size')).reset_index()
before=len(a)
u=a.drop_duplicates('Place Id').merge(agg,on='Place Id')
L('Deduplicar por Place Id',before-len(u),f'{before} filas → {len(u)} empresas únicas (las búsquedas de origen se conservan agregadas)')

# 4. Columnas inútiles
dead=[c for c in ['Email','Social Medias','Claimed','Price'] if c in u]
L('Eliminar columnas sin información',len(u),'Email y Social Medias = "visible after upgrade" en 100%; Claimed = false en 100%; Price con 1 solo valor')
L('Eliminar columnas redundantes',len(u),'Phones (≈Phone), Review URL (derivable del Place Id), Featured Image (URL temporal), Cid (incluido en Google Maps URL)')

# 5. Nombre
def clean_name(n):
    n=re.sub(r'\s+',' ',n).strip()
    for sep in [' I ',' | ',' / ']:
        if sep in n and len(n)>30: n=n.split(sep)[0].strip()
    return n
OVR_NAME={'.':'STI – San Antonio Terminal Internacional'}
u['Empresa']=u['Name'].map(lambda n: OVR_NAME.get(n,clean_name(n)))
ch=(u['Empresa']!=u['Name']).sum(); L('Normalizar nombre',int(ch),'Recorte de nombres con relleno SEO ("X I HERRAMIENTAS I ...", "X | Arriendo ..."); "." → STI (según dominio stiport.com). Original en Nombre_original')

# 6. Teléfono
def e164(r):
    ps=str(r['Phones']).split(',') if isinstance(r['Phones'],str) else []
    p=next((x for x in ps if x.strip().startswith('+')),None)
    if not p: return None
    d=re.sub(r'\D','',p)
    if d.startswith('56'):
        n=d[2:]
        if n.startswith('9') and len(n)==9: return f'+56 9 {n[1:5]} {n[5:]}'
        if n.startswith('600'): return f'+56 {n[:3]} {n[3:6]} {n[6:]}'
        if n.startswith('2') and len(n)==9: return f'+56 2 {n[1:5]} {n[5:]}'
        return f'+56 {n[:2]} {n[2:5]} {n[5:]}'
    return '+'+d
u['Telefono']=u.apply(e164,axis=1)
def ptype(t):
    if not isinstance(t,str): return 'Sin teléfono'
    if t.startswith('+56 9'): return 'Móvil'
    if t.startswith('+56 600'): return '600 (call center)'
    if t.startswith('+56'): return 'Fijo'
    return 'Extranjero'
u['Tipo_telefono']=u['Telefono'].map(ptype)
L('Normalizar teléfono a formato internacional',int(u['Telefono'].notna().sum()),'Se toma la versión +56 de "Phones"; se clasifica móvil/fijo/600')

# 7. Web
NOWEB={'instagram.com','wa.me','g.page','es.m.wikipedia.org','negocio.site','facebook.com'}
u['Dominio']=u['Domain']
u['Web_propia']=u['Dominio'].map(lambda d: 'No' if (not isinstance(d,str)) or d in NOWEB or 'ueniweb' in d else 'Sí')
u['Sitio_web']=u['Website']
L('Marcar sitios no propios',int(u['Dominio'].isin(NOWEB).sum()+u['Dominio'].fillna('').str.contains('ueniweb').sum()),'Instagram, WhatsApp, g.page, Wikipedia, negocio.site, ueniweb no cuentan como web corporativa')

# 8. Dirección
REG={'Valparaíso':'Valparaíso','Región Metropolitana':'Metropolitana','Uruguay':'Uruguay (fuera de Chile)'}
COMUNAS=['San Antonio','Cartagena','Valparaíso','Viña del Mar','Quilpué','Villa Alemana','Concón','Quintero','Casablanca','Quillota','Algarrobo',
 'Curacaví','San Bernardo','La Cisterna','Lampa','Buin','Pudahuel','Las Condes','Pedro Aguirre Cerda','Santiago']
ALIAS={'Llolleo':'San Antonio','Barrancas':'San Antonio','Placilla':'Valparaíso','Lo Vásquez':'Casablanca','Curacavi':'Curacaví'}
def norm(s): return unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode().lower()
def parse_addr(r):
    fa=r['Fulladdress']
    if not isinstance(fa,str): return pd.Series([None,None,None,'Sin dirección'])
    parts=[p.strip() for p in fa.split(',')]
    cp=re.search(r'\b(\d{7})\b',fa); cp=cp.group(1) if cp else None
    last=re.sub(r'^\d{7}\s*','',parts[-1])
    region=REG.get(last,'Valparaíso' if last=='Valparaíso' else None)
    com=None
    for p in reversed(parts[:-1] if len(parts)>1 else parts):
        p2=re.sub(r'\b\d{7}\b','',p).strip()
        for c in sorted(COMUNAS,key=len,reverse=True):
            if norm(p2)==norm(c) or norm(p2).endswith(norm(c)): com=c;break
        if not com:
            for k,v in ALIAS.items():
                if k.lower() in p2.lower(): com=v;break
        if com: break
    if not com and re.match(r'^\d{7}\s',parts[-1]): com=last
    if not com and region=='Valparaíso': com='Valparaíso'
    if region is None and com: region='Metropolitana' if com in ['Curacaví','San Bernardo','La Cisterna','Lampa','Buin','Pudahuel','Las Condes','Pedro Aguirre Cerda','Santiago'] else 'Valparaíso'
    if re.search(r'\b[23456789CFGHJMPQRVWX]{4}\+[23456789CFGHJMPQRVWX]{2,3}\b',fa): prec='Plus Code (aprox.)'
    elif 'Unnamed Road' in fa: prec='Camino sin nombre'
    elif not re.search(r'\d',parts[0]) and len(parts)<=2: prec='Solo comuna'
    elif re.search(r'\d',fa.replace(cp or '','')): prec='Con número'
    else: prec='Sin número'
    return pd.Series([com,region,cp,prec])
u[['Comuna','Region','Codigo_postal','Precision_direccion']]=u.apply(parse_addr,axis=1)
# corregir comuna desde coords si falta dirección
u['Latitud']=u['Latitude'].astype(float).round(6); u['Longitud']=u['Longitude'].astype(float).round(6)
L('Parsear dirección',len(u),'Extracción de comuna, región, código postal y nivel de precisión (Plus Code, camino sin nombre, solo comuna)')

# 9. Métricas
u['N_resenas']=pd.to_numeric(u['Review Count']).astype('Int64'); u['Rating']=pd.to_numeric(u['Average Rating'])
u['Categorias']=u['Categories'].fillna('').str.replace('Service establishment','Establecimiento de servicios').str.replace(',',', ')
u['Categoria_principal']=u['Categories'].fillna('').str.split(',').str[0].replace('',None)

# 10. Segmento
SEG=[('Portuario / Contenedores',r'portuari|contenedor'),('Química / Energía',r'químic|refiner|pintura'),
 ('Frío / Alimentos',r'frío|frigor|carnicer|alimento|marisco|matadero|supermercado'),
 ('Vitivinícola / Agrícola',r'viñedo|vino|agrícol|vivero|frutería|jardín'),
 ('Transporte de carga',r'transporte|camión|mudanza|lanzadera|envíos|autobus'),
 ('Logística / Almacenaje',r'logístic|almacén|bodega'),
 ('Maquinaria / Industrial',r'maquinaria|herramient|industrial|ferreter|soldadura|fábrica|repuesto|embalaje|contratista|ingenier|construcción|metal|incendio|tierras|limpieza'),
 ]
def seg(r):
    t=(str(r['Categories'])+' '+r['Empresa']).lower()
    if re.search(r'viña|vineyard|bodegas re\b',r['Empresa'].lower()) and 'bodega' in t: return 'Vitivinícola / Agrícola'
    for s,p in SEG:
        if re.search(p,t): return s
    q=r['_busq'].lower()
    for s,k in [('Transporte de carga','transport'),('Logística / Almacenaje','log'),('Frío / Alimentos','frigor'),('Vitivinícola / Agrícola','agr'),('Maquinaria / Industrial','maquin'),('Química / Energía','quím')]:
        if k in q: return s
    return 'Otros / Servicios empresariales'
u['Segmento']=u.apply(seg,axis=1)

# 11. Relevancia B2B
BAJA_CAT=r'escuela|farmacia|tienda de alimentación|bebidas alcohólicas|delicatessen|carnicería|supermercado|frutería|centro de jardinería|oficinas virtuales|consultora de administración|restaurante|empresa de autobuses|tienda de alimentos para animales'
OVR={ # Place Id -> (relevancia, motivo)
}
OVR_NAME_REL={
 'Frigorífico Casa Blanca S.A.':('Descartar','Empresa en Uruguay (homónimo de Casablanca, Chile)'),
 'Bodega Simón Bolívar':('Descartar','Edificio histórico (ficha Wikipedia), no es empresa'),
 'Transporte de carga por carretera':('Descartar','Ficha genérica sin empresa identificable'),
 'Las Bodegas':('Descartar','Ficha genérica sin datos'),
 'San Antonio Port - Chile':('Descartar','Duplicado de Puerto San Antonio sin datos'),
 'portería ENAP':('Descartar','Portería de ENAP Aconcagua (duplicado)'),
 'Centro De Abastecimiento De La Armada':('Baja','Institución militar, no prospecto comercial'),
 'Equipos Gastronomicos Rorland':('Baja','Comercio minorista de equipos gastronómicos'),
 'La Industrial Outlet':('Baja','Ferretería minorista'),
 'Bodega el Campo':('Baja','Comercio minorista'),
 'Viña Casas del Bosque':('Alta','Viña de gran escala (enoturismo + producción)'),
 'CALCA Curacaví':('Media','Sucursal de Calca (insumos agrícolas); categoría Google dudosa'),
}
def rel(r):
    if r['Empresa'] in OVR_NAME_REL: return pd.Series(OVR_NAME_REL[r['Empresa']])
    if r['Name'] in OVR_NAME_REL: return pd.Series(OVR_NAME_REL[r['Name']])
    if re.search(BAJA_CAT,str(r['Categories']).lower()) and not re.search(r'frío|logística|viñedo',str(r['Categories']).lower()):
        return pd.Series(['Baja','Negocio minorista / no industrial ('+r['Categoria_principal']+')'])
    if pd.isna(r['Fulladdress']) and pd.isna(r['Telefono']): return pd.Series(['Baja','Sin dirección ni teléfono'])
    if pd.isna(r['Region']): return pd.Series(['Media','Sin dirección: verificar ubicación'])
    out = r['Region']!='Valparaíso' and r['Comuna']!='Curacaví'
    if out: return pd.Series(['Media','Fuera de la zona objetivo (Región de Valparaíso / Curacaví)'])
    contact = pd.notna(r['Telefono']) or r['Web_propia']=='Sí'
    if not contact: return pd.Series(['Media','Sin teléfono ni web propia: requiere investigación'])
    return pd.Series(['Alta','Empresa industrial/logística en zona objetivo con canal de contacto'])
u[['Relevancia','Motivo_relevancia']]=u.apply(rel,axis=1)

# 12. Grupo empresarial (union-find por dominio propio y teléfono)
par={i:i for i in u.index}
def f(x):
    while par[x]!=x: par[x]=par[par[x]]; x=par[x]
    return x
for key in ['Dominio','Telefono']:
    sub=u[u[key].notna() & ((u['Web_propia']=='Sí') if key=='Dominio' else True)]
    for _,g in sub.groupby(key):
        ids=list(g.index)
        for i in ids[1:]: par[f(i)]=f(ids[0])
u['_root']=[f(i) for i in u.index]
gs=u.groupby('_root')
lead={r:g.sort_values('N_resenas',ascending=False,na_position='last')['Empresa'].iloc[0] for r,g in gs}
size=gs.size()
u['Grupo_empresa']=[lead[r] if size[r]>1 else None for r in u['_root']]
u['N_sedes_grupo']=[int(size[r]) for r in u['_root']]
L('Agrupar sedes de una misma empresa',int((u['N_sedes_grupo']>1).sum()),f'{int((size>1).sum())} grupos detectados por dominio web o teléfono compartido (no se borran: son sedes distintas)')

# 13. Flags y score
def flags(r):
    fl=[]
    if pd.isna(r['Telefono']): fl.append('sin_telefono')
    if r['Web_propia']=='No': fl.append('sin_web_propia')
    if r['Precision_direccion'] not in ('Con número',): fl.append('direccion_imprecisa')
    if pd.isna(r['N_resenas']): fl.append('sin_resenas')
    if pd.isna(r['Opening Hours']): fl.append('sin_horario')
    if pd.isna(r['Categories']): fl.append('sin_categoria')
    if r['Empresa']!=r['Name']: fl.append('nombre_corregido')
    return ', '.join(fl)
u['Flags_calidad']=u.apply(flags,axis=1)
def score(r):
    s=0
    s+=25 if pd.notna(r['Telefono']) else 0
    s+=25 if r['Web_propia']=='Sí' else 0
    s+=15 if r['Precision_direccion']=='Con número' else (5 if r['Precision_direccion'] in ('Sin número','Camino sin nombre','Plus Code (aprox.)') else 0)
    n=r['N_resenas'] if pd.notna(r['N_resenas']) else 0
    s+=15 if n>=50 else (10 if n>=10 else (5 if n>=1 else 0))
    s+=10 if pd.notna(r['Rating']) and r['Rating']>=4 else 0
    s+=10 if pd.notna(r['Opening Hours']) else 0
    return s
u['Score_completitud']=u.apply(score,axis=1)

cols=['Place Id','Empresa','Nombre_original','Grupo_empresa','N_sedes_grupo','Segmento','Relevancia','Motivo_relevancia','Score_completitud',
 'Telefono','Tipo_telefono','Sitio_web','Dominio','Web_propia','Comuna','Region','Codigo_postal','Direccion','Precision_direccion','Latitud','Longitud',
 'Categoria_principal','Categorias','N_resenas','Rating','Horario','Google_Maps_URL','Busquedas_origen','N_apariciones','Flags_calidad']
u=u.rename(columns={'Name':'Nombre_original','Fulladdress':'Direccion','Opening Hours':'Horario','Google Maps URL':'Google_Maps_URL','_busq':'Busquedas_origen','_n':'N_apariciones'})
u['Horario']=u['Horario'].fillna('').str.replace('a.m.','am').str.replace('p.m.','pm').str.replace(',',' | ').replace('',None)
relo={'Alta':0,'Media':1,'Baja':2,'Descartar':3}
u=u[cols].sort_values(['Relevancia','Score_completitud','N_resenas'],key=lambda s: s.map(relo) if s.name=='Relevancia' else s,ascending=[True,False,False])
u=u.rename(columns={'Place Id':'Place_Id'})
u.to_pickle(OUT+'/clean.pkl'); pd.DataFrame(log).to_pickle(OUT+'/log.pkl')
print(u['Relevancia'].value_counts()); print(u['Segmento'].value_counts()); print(u['Comuna'].value_counts()); print(u['Precision_direccion'].value_counts())
print(u[u['Grupo_empresa'].notna()][['Empresa','Grupo_empresa']].to_string())
print(u[u['Relevancia']!='Alta'][['Empresa','Relevancia','Motivo_relevancia']].to_string())
