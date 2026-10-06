import pdfplumber, sys, re, json
pdf=pdfplumber.open(sys.argv[1])
ORANGE=(0.909804, 0.415686, 0.164706)
BOUNDS=[145,270,445,565,685,10000]
def col(x):
    for i,b in enumerate(BOUNDS):
        if x<b: return i
def clean(t):
    t=t.replace('(cid:3)',' ').replace('(cid:193)','fl')
    t=re.sub(r'\(cid:\d+\)','',t)
    t=''.join('fi' if (0xE000<=ord(ch)<=0xF8FF or ord(ch)==0xFFFD or ord(ch)>0xFFFF) else ch for ch in t)
    return re.sub(r'\s+',' ',t).strip()
def isor(c): return c and len(c)==3 and abs(c[0]-ORANGE[0])<0.03 and abs(c[1]-ORANGE[1])<0.03
def iswhite(c): return c and len(c)==3 and min(c)>0.98
entries=[]; rubro=None; cur=None
import unicodedata
CANON=["Agua, análisis, tecnología y productos","Ambulancias y servicios de emergencia","Andamios y trabajos en altura","Campamentos e instalaciones","Capacitación, consultoría y auditoría","Certificación, inspección y peritaje","Detección de gases: equipos y servicios","Electricidad industrial","Equipos de protección personal","Higiene industrial, saneamiento y limpieza","Iluminación industrial","Mantenimiento industrial","Máquinas, herramientas y repuestos","Medioambiente: productos y servicios","Perforación geológica","Protección contra incendios","Reflectivos y señalización","Rescate y salvamento","Ropa industrial","Salud ocupacional","Seguridad vial y prevención de accidentes vehiculares","Seguridad y vigilancia interna","Soldadura industrial","Sostenimiento de rocas y suelos","Tecnología de la información, comunicación y telecomunicaciones","Transporte de combustibles y materiales peligrosos","Vehículos: ventas, alquiler, repuestos y reparaciones","Ventilación industrial y subterránea","Voladura, explosivos y polvorines","Empresas contratistas y conexas","Representaciones y distribuidores"]
def nrm(t): return re.sub(r'[^a-z ]','',unicodedata.normalize('NFKD',t.lower()).encode('ascii','ignore').decode()).split()
def canon(t,cur):
    w=nrm(t)
    if not w: return None
    j=' '.join(w)
    if len(j)<6: return None
    hits=[c for c in CANON if ' '+j+' ' in ' '+' '.join(nrm(c))+' ']
    if cur in hits: return cur
    return hits[0] if hits else None
PH=re.compile(r'0\s?800[\s\d]{5,}\d|(\(\+?\d{2,3}\)\s*)?\+?\d[\d\-]{4,}\d')
for pi,p in enumerate(pdf.pages):
    if pi<4: continue
    words=p.extract_words(extra_attrs=["non_stroking_color","size"],x_tolerance=1.5,keep_blank_chars=False)
    cols={}
    for w in words:
        if w['top']>570: continue  # page numbers
        cols.setdefault(col(w['x0']),[]).append(w)
    for ci in sorted(cols):
        ws=sorted(cols[ci],key=lambda w:(round(w['top']),w['x0']))
        lines=[]
        for w in ws:
            if lines and abs(lines[-1]['top']-w['top'])<2.5:
                lines[-1]['ws'].append(w)
            else: lines.append({'top':w['top'],'ws':[w]})
        cur=None
        for L in lines:
            txt=clean(' '.join(w['text'] for w in L['ws']))
            if not txt: continue
            big=any(w['size']>10 for w in L['ws'])
            if big and all(iswhite(w['non_stroking_color']) or w['size']>10 for w in L['ws']):
                r=canon(txt,rubro)
                if r: rubro=r; cur=None
                continue
            if big: continue
            orange=sum(isor(w['non_stroking_color']) for w in L['ws'])>len(L['ws'])/2
            if rubro is None: continue
            if orange:
                if cur and not cur['lines'] and L['top']-cur['last']<10:
                    cur['name']+=' '+txt
                else:
                    cur={'rubro':rubro,'name':txt,'lines':[],'page':pi+1}
                    entries.append(cur)
                cur['last']=L['top']
            else:
                if cur and L['top']-cur['last']>14: cur=None
                if cur is None:
                    cur={'rubro':rubro,'name':None,'lines':[],'page':pi+1}; entries.append(cur)
                cur['lines'].append(txt); cur['last']=L['top']
out=[]
for e in entries:
    emails=[];webs=[];phones=[];addr=[];contact=[]
    lastc=False
    for t in e['lines']:
        t=re.sub(r'(^|\s)fi(?=\s|$)',' ',t).strip(' /')
        if not t: continue
        if re.match(r'(?i)contacto\s*:',t): contact.append(re.sub(r'(?i)^contacto\s*:?\s*','',t)); lastc=True; continue
        if lastc and not re.search(r'\d|@|www',t) and len(t.split())<=3: contact[-1]+=' '+t; continue
        lastc=False
        em=re.findall(r'[\w.\-+]+@[\w.\-]+\.\w+',t)
        if em:
            emails+=em; continue
        if re.search(r'(?i)(www\.|https?://|ww\.)',t) or re.fullmatch(r'[\w\-]+(\.[\w\-]+)+(/\S*)?',t):
            webs.append(t.split()[0] if ' ' in t and 'www' in t.split()[0] else t); continue
        ph=[m.group(0).strip() for m in PH.finditer(t)]
        rest=PH.sub('',t).strip(' /-,')
        phones+=ph
        if re.search(r'[A-Za-zÁÉÍÓÚáéíóúñÑ]{2,}',rest): addr.append(rest)
    a=' '.join(addr)
    REG=['Lima','Callao','Arequipa','Cusco','Moquegua','Áncash','Ancash','La Libertad','Trujillo','Piura','Ica','Junín','Huancayo','Cajamarca','Puno','Tacna','Lambayeque','Chiclayo','Pasco','Huánuco','Apurímac','Ayacucho','Huancavelica','Ucayali','Loreto','Tumbes','Amazonas','Chimbote','Huaraz','Juliaca','Ilo','Espinar']
    best=(-1,'')
    for r in REG:
        for m in re.finditer(r'\b'+r+r'\b',a):
            if m.start()>best[0]: best=(m.start(),r)
    city=best[1].replace('Ancash','Áncash')
    if city in ('Trujillo',): city='La Libertad'
    name=e['name']
    if not name:
        if not (emails or webs or phones): continue
        dom=(webs[0] if webs else emails[0].split('@')[1] if emails else '')
        dom=re.sub(r'^(https?://)?(www?\.)?','',dom).split('.')[0]
        name=(dom.capitalize() if dom else 'Empresa sin nombre')+' (nombre por validar)'
    out.append({'orphan':not e['name'],'rubro':e['rubro'],'empresa':name,'direccion':a,'ciudad':city,'telefonos':' / '.join(phones),
                'email':'; '.join(dict.fromkeys(emails)),'web':'; '.join(dict.fromkeys(webs)),'contacto':'; '.join(contact),'pagina':e['page']})
json.dump(out,open(sys.argv[2],'w'),ensure_ascii=False,indent=1)
print(len(out))
