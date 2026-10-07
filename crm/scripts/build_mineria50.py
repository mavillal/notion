import csv, re, sys, json, collections, unicodedata
import openpyxl
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
# usage: build_mineria50.py m50.csv Directorio.xlsx ExpoMina.xlsx Exponor.xlsx AquaSur.xlsx FIDAE.xlsx ExpoSanJuan.xlsx out.xlsx
src, out = sys.argv[1], sys.argv[-1]
def norm(s): return unicodedata.normalize('NFKD',s.upper()).encode('ascii','ignore').decode()
def key(n):
    n=norm(n).lower(); n=re.sub(r'\(.*?\)','',n)
    n=re.sub(r'\b(s\.?\s?a\.?\s?a\.?|s\.?\s?a\.?\s?c\.?|s\.?\s?a\.?|s\.?r\.?l\.?|ltda\.?|spa|sac|saa|peru|chile|argentina|del|de|y|cia|compania|minera|limited|ltd|inc|corp|group|grupo)\b','',n)
    return re.sub(r'[^a-z0-9]','',n)
def load(path):
    ws=openpyxl.load_workbook(path)['Contactos']; h=[c.value for c in ws[1]]; m={}
    for r in ws.iter_rows(min_row=2,values_only=True):
        d=dict(zip(h,r))
        if d.get('Empresa') and key(d['Empresa']): m.setdefault(key(d['Empresa']),d)
    return m
NAMES=['Directorio SSO','ExpoMina','Exponor','AquaSur','FIDAE','Expo San Juan']
CRMS=[(n,load(p)) for n,p in zip(NAMES,sys.argv[2:8])]
rows=list(csv.DictReader(open(src,encoding='utf-8-sig')))

def title_of(r):
    m=re.search(r'"(.+)"\s+(?:CARGO:|[A-ZÁÉÍÓÚÑ ]+:)',r['TextoOriginal'])
    return (m.group(1) if m else r['TituloPresentacion']).strip()
def fix_title(t):
    t=t.strip().rstrip('.')
    t=re.sub(r'\s+,',',',t); t=re.sub(r'\s+',' ',t)
    if not t.isupper(): return t
    t=t.lower()
    for a,b in [('iot–ia','IoT–IA'),('iiot','IIoT'),('ia','IA'),('ai','AI'),('sag','SAG'),('cmh','CMH'),('tmc','TMC'),('deeplabv3+','DeepLabv3+'),
                ('poderosa','Poderosa'),('seequent','Seequent'),('buenaventura','Buenaventura'),('optimus','Optimus'),('lean','Lean'),('mine vision','Mine Vision'),('semiautonomo','semiautónomo'),
                ('compañía de minas','Compañía de Minas'),('minería 5.0','Minería 5.0'),('mineria','minería')]:
        t=re.sub(r'(?<![\w])'+re.escape(a)+r'(?![\w])',b,t)
    t=re.sub(r'([./] )([a-záéíóú])',lambda m:m.group(1)+m.group(2).upper(),t)
    return t[:1].upper()+t[1:]
def area(s): return {'MINERÍA 5.0':'Minería 5.0','EXCELENCIA OPERACIONAL':'Excelencia operacional'}.get(s,s.title())

S_MINE='Cliente final – minera'
S_COMP='Competidor / tecnología de seguridad'
S_TECH='Partner tecnológico (IA, IIoT, datos)'
S_OEM='Partner OEM – equipos mineros'
S_CONS='Canal / aliado – consultoría y formación'
S_ACAD='Academia / investigación'
S_IND='Profesional independiente / talento'
PRI={S_MINE:'Alta',S_COMP:'Alta',S_TECH:'Media',S_OEM:'Media',S_CONS:'Media',S_ACAD:'Media',S_IND:'Baja'}
ACC={S_MINE:'Contactar post-ponencia; mapear a Gerente SSO/Operaciones y proponer piloto SafetyMind',
 S_COMP:'Inteligencia competitiva: revisar ponencia y oferta; evaluar si es competidor o integrable',
 S_TECH:'Explorar integración/alianza (datos de planta, IIoT, analítica)',
 S_OEM:'Explorar alianza con representante local (equipos y talleres)',
 S_CONS:'Explorar alianza comercial / canal de formación',
 S_ACAD:'Vincular para I+D, casos de estudio o validación académica',
 S_IND:'Mantener en red (posible champion o talento)'}
# Curación por fila de origen (orden del CSV): [(nombre, cargo, linkedin, ok_linkedin)], empresa, país, segmento, nota
CUR=[
 ([('Thierry de Saint Pierre','Profesor investigador',True)],'Universidad San Sebastián','Chile',S_ACAD,'Dr.; nombre y organización venían invertidos en la fuente.'),
 ([('Carlos Calderón','Director de la Academia de IA, Datos y Ciencias de la Computación',False)],'UTEC','Perú',S_ACAD,''),
 ([('Arturo Vásquez','Gerente Regional de Desarrollo de Negocios para Latinoamérica',False)],'Epiroc','Perú',S_OEM,''),
 ([('Luis Rojas','Gerente de Tecnologías de Información',True)],'Hudbay Perú','Perú',S_MINE,''),
 ([('Alessandra Morante','CEO',True),('Angela Barreda','Co-presentadora',None)],'Innovation Academy & Experience','Perú',S_CONS,'Organización corregida (la fuente ponía a Alessandra Morante como organización).'),
 ([('Danilo Sturiza','COO y Founder',True),('Claudia Castillo','Advisor',None)],'Smart Mining','Perú',S_TECH,'IIoT para mantenimiento y espesamiento.'),
 ([('Vicente Rojas','Gerente Tecnología & IA',True)],'Empírica','Perú',S_TECH,''),
 ([('Martin Miranda','Asistente de Supervisor General de Mantenimiento Eléctrico Mina',None)],'Investigador independiente','Perú',S_IND,'Trabaja en mina; desarrolló visión artificial para conteo de camiones: posible champion interno.'),
 ([('Terry Williams','Director Ejecutivo',None)],'New Project Perú SAC','Perú',S_COMP,'IoT-IA para tracking, seguridad activa y control de acceso en minería subterránea: posible competidor.'),
 ([('Angel Navarro','Sales & Configuration Manager',True)],'Procetradi SAC','Perú',S_TECH,'Centros integrados de operaciones.'),
 ([('Fernando Velásquez','Gerente General',False)],'Feedback Economic Investments and B…','Perú',S_IND,'Nombre de la organización viene truncado en la fuente. Tema blockchain/tokenización (fuera de foco).'),
 ([('Leonel Aguilar','Gerente Posventa Sudamérica',True)],'Watergenics SAC','Perú',S_TECH,'Monitoreo de agua en tiempo real.'),
 ([('Andrea Quevedo','Jefe de Transformación Digital y Proyectos',True)],'Alpayana','Perú',S_MINE,'Lidera scoop semiautónomo y automatización en planta.'),
 ([('Hyder Mamani','Especialista en Circuitos de Molienda',True)],'Weir Minerals Perú','Perú',S_OEM,''),
 ([('Lucas Torres','Team Lead, Technical Solutions, Mining',True)],'Seequent','Perú',S_TECH,''),
 ([('Leonardo Forero','Investigador IA',None),('Eloy Paredes','Consultor Senior',None),('Angela Donaires','Especialista en Deep Learning',None)],'MGM Corporate Resources','Perú',S_TECH,'Caso aplicado en Compañía de Minas Buenaventura (deep learning en imágenes geo-mineras).'),
 ([('Andrew Hulme','General Manager & SSE',True)],'Sibanye-Stillwater','Sudáfrica',S_MINE,''),
 ([('Luis Reynafarge','Experto en Excelencia Operacional',True)],'Independiente','Perú',S_CONS,'Consultor Lean + IA multiagente.'),
 ([('David Bayona','Director de Proyectos',True)],'Hot Chili Limited','Chile',S_MINE,'Proyecto Costa Fuego (Chile).'),
 ([('Marcelo Santillana','Director',None)],'Compañía Minera Poderosa','Perú',S_MINE,'Nombre y organización venían invertidos en la fuente.'),
 ([('Raúl Benavides','Presidente del Directorio',True),('Homar Lozano','Co-autor',None)],'Sociedad Minera El Brocal S.A.A.','Perú',S_MINE,'PRIORIDAD: ponencia sobre IA para prevenir accidentes (encaje directo con SafetyMind). Grupo Buenaventura.'),
 ([('Ruddy Morales','Superintendente Corporativo de Excelencia',True),('Glasson Fonseca','Gerente Corporativo de Transformación',None)],'Consorcio Minero Horizonte','Perú',S_MINE,'Hoja de ruta de transformación (caso CMH).'),
 ([('Jorge Infantes','Superintendente Planta Concentradora',None),('Luis Chacón','Analítica Avanzada',None)],'Marcobre','Perú',S_MINE,'Optimus AI en planta concentradora (Grupo Minsur).'),
 ([('Iván Velásquez','Gerente Corporativo de Control de Gestión y Excelencia Operacional',None)],'Compañía Minera Lincuna S.A.','Perú',S_MINE,''),
 ([('Diego Rey','Investigador independiente',None)],'Investigador independiente','Perú',S_IND,'Predicción de fallas de motor de camión minero.'),
 ([('Juan Pablo Rodriguez','Gerente Fundición/Refinería y Balance Metalúrgico',None)],'Empírica','Perú',S_TECH,'Misma empresa que Vicente Rojas (unificado "EMPIRICA"/"EMPÍRICA").'),
 ([('Ivan Quintanilla','Director',None)],'Supply Talent Consulting','Perú',S_CONS,'Título de la ponencia venía truncado en la fuente; corregido desde el texto original.'),
]
assert len(CUR)==len(rows)
ALIAS={'Weir Minerals Perú':'WEIR'}
out_rows=[]; log=[]
for i,(r,(people,org,pais,seg,nota)) in enumerate(zip(rows,CUR),2):
    t=fix_title(title_of(r)); li=r['LinkedIn'].strip()
    if t.lower()!=fix_title(r['TituloPresentacion']).lower(): log.append((i,'Título completado desde TextoOriginal',r['TituloPresentacion'],t))
    if r['RequiereRevision']=='Sí': log.append((i,'Fila marcada para revisión (RequiereRevision=Sí)',r['ParticipantesDetectados'] or r['Organizacion'],f'Separada en {len(people)} contactos: '+', '.join(p[0] for p in people) if len(people)>1 else 'Corregida: nombre y organización estaban invertidos'))
    if r['Organizacion']!=org: log.append((i,'Organización normalizada',r['Organizacion'],org))
    names=[p[0] for p in people]
    for j,(n,cargo,ok) in enumerate(people):
        link=li if j==0 and li else ''
        ver='Sin LinkedIn' if not link else ('Coincide con el nombre' if ok else 'Verificar (handle genérico)')
        ko=key(ALIAS.get(org,org))
        cruce=[f"{nm} (ID {m[ko]['ID']})" for nm,m in CRMS if ko in m]
        notas='Conferencista Minería 5.0 2026 (Encuentro de Metalurgia).'+(' '+nota if nota else '')
        if len(people)>1: notas+=' Presentación conjunta.'
        if cruce: notas+=' Empresa también en: '+'; '.join(cruce)+'.'
        out_rows.append({'Nombre':n,'Cargo':cargo,'Empresa':org,'País':pais,'Email':'','LinkedIn':link,'Origen':'Minería 5.0 2026 – Encuentro de Metalurgia',
          'Estado':'Prospecto','Prioridad':PRI[seg],'Próxima acción':ACC[seg],'Notas':notas,'Segmento SafetyMind':seg,'Bloque temático':area(r['AreaTematica']),
          'Tipo presentación':r['TipoPresentacion'],'Ponencia':t,'Co-presentadores':', '.join(x for x in names if x!=n),'Verificación LinkedIn':ver,
          'Cruce otros CRM':'; '.join(cruce) or 'No','Fila origen':i})
ORDER=[S_MINE,S_COMP,S_TECH,S_OEM,S_CONS,S_ACAD,S_IND]; po={'Alta':0,'Media':1,'Baja':2}
out_rows.sort(key=lambda x:(po[x['Prioridad']],ORDER.index(x['Segmento SafetyMind']),x['Empresa'],x['Fila origen']))
COLS=['ID','Nombre','Cargo','Empresa','País','Email','LinkedIn','Origen','Estado','Prioridad','Próxima acción','Notas','Segmento SafetyMind','Bloque temático','Tipo presentación','Ponencia','Co-presentadores','Verificación LinkedIn','Cruce otros CRM']
wb=Workbook(); HF=PatternFill('solid',fgColor='1F3A5F'); HFont=Font(bold=True,color='FFFFFF')
def header(ws,cols,widths):
    ws.append(cols)
    for k,cn in enumerate(cols,1):
        c=ws.cell(1,k); c.fill=HF; c.font=HFont; c.alignment=Alignment(wrap_text=True,vertical='center')
        ws.column_dimensions[get_column_letter(k)].width=widths[k-1] if k-1<len(widths) else 16
    ws.freeze_panes='A2'
ws=wb.active; ws.title='Contactos'
header(ws,COLS,[5,22,40,32,12,24,45,30,12,10,50,60,36,20,14,70,30,24,30])
for k,x in enumerate(out_rows,1): ws.append([k]+[x[c] for c in COLS[1:]])
n=len(out_rows)+1
t=Table(displayName='Contactos',ref=f"A1:{get_column_letter(len(COLS))}{n}"); t.tableStyleInfo=TableStyleInfo(name='TableStyleMedium2',showRowStripes=True); ws.add_table(t)
dv1=DataValidation(type='list',formula1='"Prospecto,Contactado,Reunión,Propuesta,Cliente,No interesado"',allow_blank=True); dv1.add(f'I2:I{n+300}')
dv2=DataValidation(type='list',formula1='"Alta,Media,Baja,Por calificar"',allow_blank=True); dv2.add(f'J2:J{n+300}')
ws.add_data_validation(dv1); ws.add_data_validation(dv2)
ws2=wb.create_sheet('Oportunidades'); header(ws2,['ID oportunidad','ID contacto','Empresa','Necesidad','Solución propuesta','Etapa','Valor estimado','Probabilidad','Próxima acción','Fecha'],[14,12,40,40,40,14,14,12,40,12])
ws3=wb.create_sheet('Resumen'); header(ws3,['Segmento SafetyMind','Alta','Media','Baja','Total'],[45,10,10,10,10])
for s in ORDER:
    k=ws3.max_row+1; ws3.append([s]+[f'=COUNTIFS(Contactos!$M:$M,$A{k},Contactos!$J:$J,{c}$1)' for c in 'BCD']+[f'=SUM(B{k}:D{k})'])
k=ws3.max_row+1; ws3.append(['Total']+[f'=SUM({c}2:{c}{k-1})' for c in 'BCDE']); ws3.append([''])
ws3.append(['Indicadores','Cantidad'])
for lab,f in [('Filas en archivo original',len(rows)),('Personas (tras separar presentaciones conjuntas)','=COUNTA(Contactos!$B$2:$B$500)'),('Organizaciones únicas',f'=SUMPRODUCT(1/COUNTIF(Contactos!$D$2:$D${n},Contactos!$D$2:$D${n}))'),
  ('Con LinkedIn','=COUNTIF(Contactos!$G$2:$G$500,"?*")'),('LinkedIn a verificar','=COUNTIF(Contactos!$R$2:$R$500,"Verificar*")'),('Contactos de empresas ya presentes en otros CRM','=COUNTIF(Contactos!$S$2:$S$500,"*(ID*")')]:
    ws3.append([lab,f])
ws4=wb.create_sheet('Log depuración'); header(ws4,['Fila origen','Corrección','Valor original','Valor final'],[12,36,50,70])
for l in log: ws4.append(list(l))
ws5=wb.create_sheet('Eventos'); header(ws5,['Evento','País','Valor para SafetyMind'],[45,12,100])
ws5.append(['Minería 5.0 2026 – Encuentro de Metalurgia','Perú','Conferencistas sobre IA, IIoT y excelencia operacional en minería: tomadores de decisión de mineras peruanas (El Brocal, Marcobre, CMH, Poderosa, Alpayana, Hudbay, Lincuna) y proveedores tecnológicos. El Brocal presenta IA para prevenir accidentes.'])
wb.save(out)
print(len(rows),len(out_rows),collections.Counter(x['Prioridad'] for x in out_rows),collections.Counter(x['Segmento SafetyMind'] for x in out_rows))
print('orgs',len({x['Empresa'] for x in out_rows}),'li',sum(bool(x['LinkedIn']) for x in out_rows))
print('cruce',{x['Empresa']:x['Cruce otros CRM'] for x in out_rows if x['Cruce otros CRM']!='No'})
print('log',len(log)); [print(l) for l in log if l[1].startswith('Título')]
json.dump(out_rows,open(out+'.json','w'),ensure_ascii=False)
